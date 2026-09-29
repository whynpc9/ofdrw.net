#!/usr/bin/env python3
"""Validate the exact SDK/tool package set and bind consumer tests to its hashes."""
import argparse
import hashlib
import json
import os
import re
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

ROOT = Path(__file__).resolve().parents[1]
DEPENDENCY_BASELINE = ROOT / 'docs' / 'third-party-dependency-baseline.json'

PACKAGES = (
    'Ofdrw.Net.Core', 'Ofdrw.Net.Packaging', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Reader',
    'Ofdrw.Net.Converter.Abstractions', 'Ofdrw.Net.Converter.Docx', 'Ofdrw.Net.Converter.Pdf',
    'Ofdrw.Net.Converter.Svg', 'Ofdrw.Net.Signatures', 'Ofdrw.Net.Converter', 'Ofdrw.Net.Cli',
)


def inspect(directory: Path, version: str):
    verify_release_notices()
    if not re.fullmatch(r'\d+\.\d+\.\d+(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?', version):
        raise ValueError('Expected a SemVer package version.')
    expected = {f'{name}.{version}.nupkg' for name in PACKAGES}
    actual = {path.name for path in directory.glob('*.nupkg')}
    if expected != actual:
        raise ValueError(f'Package set mismatch. Missing={sorted(expected-actual)}, unexpected={sorted(actual-expected)}')
    packages = []
    for name in PACKAGES:
        path = directory / f'{name}.{version}.nupkg'
        if path.is_symlink():
            raise ValueError(f'Package must be a regular artifact: {path.name}')
        with zipfile.ZipFile(path) as package:
            if package.testzip() is not None:
                raise ValueError(f'Corrupt ZIP payload: {path.name}')
            specs = [entry for entry in package.namelist() if entry.endswith('.nuspec')]
            if len(specs) != 1:
                raise ValueError(f'Expected one nuspec: {path.name}')
            root = ET.fromstring(package.read(specs[0]))
            values = {node.tag.rsplit('}', 1)[-1]: node.text for node in root.iter()}
            if values.get('id') != name or values.get('version') != version:
                raise ValueError(f'Package identity/version mismatch: {path.name}')
            license_node = next((node for node in root.iter() if node.tag.rsplit('}', 1)[-1] == 'license'), None)
            if license_node is None or license_node.get('type') != 'expression' or license_node.text != 'MIT':
                raise ValueError(f'Missing or mismatched MIT license expression: {path.name}')
            for dependency in root.iter():
                if dependency.tag.rsplit('}', 1)[-1] == 'dependency' and dependency.get('id', '').startswith('Ofdrw.Net.'):
                    if dependency.get('version') not in (version, f'[{version}]'):
                        raise ValueError(f'Mismatched SDK dependency in {path.name}: {dependency.attrib}')
            required = (['tools/net10.0/any/Ofdrw.Net.Cli.dll'] if name == 'Ofdrw.Net.Cli'
                        else [] if name == 'Ofdrw.Net.Converter'
                        else [f'lib/{framework}/{name}{extension}' for framework in ('netstandard2.0', 'netstandard2.1') for extension in ('.dll', '.xml')])
            required += ['README.md', 'THIRD-PARTY-NOTICES.md', 'docs/feature-parity.md', 'docs/conversion-contracts.md']
            for entry in required:
                if entry not in package.namelist():
                    raise ValueError(f'Missing {entry} in {path.name}')
            if package.read('THIRD-PARTY-NOTICES.md') != (ROOT / 'THIRD-PARTY-NOTICES.md').read_bytes():
                raise ValueError(f'Third-party notices differ from the reviewed source: {path.name}')
        packages.append({'id': name, 'file': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
    return {'version': version, 'packages': packages}


def verify_release_notices():
    props = ET.parse(ROOT / 'Directory.Build.props').getroot()
    expression = [node.text for node in props.iter() if node.tag == 'PackageLicenseExpression']
    if expression != ['MIT'] or 'MIT License' not in (ROOT / 'LICENSE').read_text():
        raise ValueError('Repository LICENSE and package license expression must both declare MIT.')
    notices = (ROOT / 'THIRD-PARTY-NOTICES.md').read_text()
    if not notices.startswith('# Third-Party Notices'):
        raise ValueError('Third-party notice heading is missing.')
    direct = set()
    for project in (ROOT / 'src').glob('*/*.csproj'):
        for node in ET.parse(project).iter('PackageReference'):
            direct.add((node.attrib['Include'], node.attrib['Version']))
    for package, version in sorted(direct):
        row = re.search(r'^\| ' + re.escape(package) + r' \| ' + re.escape(version) + r' \| ([^|]+) \| (https://[^| ]+) \|$', notices, re.M)
        if row is None or not row.group(1).strip():
            raise ValueError(f'Third-party declaration missing license/source for {package} {version}.')
    inventory = resolved_dependency_inventory()
    actual = {f"{item['id']}/{item['version']}": item for item in inventory}
    for package, version in sorted(direct):
        item = actual.get(f'{package}/{version}')
        row = re.search(r'^\| ' + re.escape(package) + r' \| ' + re.escape(version) + r' \| ([^|]+) \| (https://[^| ]+) \|$', notices, re.M)
        declared_license, declared_source = row.group(1).strip(), row.group(2).rstrip('/')
        if item is None or item['source'].rstrip('/') != declared_source:
            raise ValueError(f'Third-party source differs from resolved {package} {version} metadata.')
        if item['license_type'] == 'expression':
            matches = item['license'] == declared_license
        elif item['license_type'] == 'file':
            matches = declared_license == 'MIT' and item['license_file_heading'] == 'MIT License'
        else:
            matches = False
        if not matches:
            raise ValueError(f'Third-party license differs from resolved {package} {version} metadata.')
    if json.loads(DEPENDENCY_BASELINE.read_text()) != inventory:
        raise ValueError('Resolved third-party dependency/license graph changed; review and update its baseline.')


def resolved_dependency_inventory():
    libraries = set()
    for project in (ROOT / 'src').glob('*/*.csproj'):
        assets = project.parent / 'obj' / 'project.assets.json'
        if not assets.exists():
            raise ValueError(f'Restore {project.name} before checking release dependency licenses.')
        for identity, details in json.loads(assets.read_text())['libraries'].items():
            if details['type'] == 'package':
                libraries.add(tuple(identity.split('/', 1)))
    package_root = Path(os.environ.get('NUGET_PACKAGES', Path.home() / '.nuget' / 'packages'))
    inventory = []
    for name, version in sorted(libraries):
        directory = package_root / name.lower() / version.lower()
        metadata = ET.parse(directory / f'{name.lower()}.nuspec').getroot()
        values = {node.tag.rsplit('}', 1)[-1]: node for node in metadata.iter()}
        license_node = values.get('license')
        license_type = license_node.get('type') if license_node is not None else 'legacy-url'
        license_value = (license_node.text if license_node is not None else values.get('licenseUrl').text)
        item = {'id': name, 'version': version, 'license_type': license_type,
                'license': license_value, 'source': values.get('projectUrl').text if values.get('projectUrl') is not None else ''}
        if license_type == 'file':
            license_bytes = (directory / license_value).read_bytes()
            item['license_file_sha256'] = hashlib.sha256(license_bytes).hexdigest()
            item['license_file_heading'] = license_bytes.decode('utf-8-sig').splitlines()[0].lstrip('# ').strip()
        inventory.append(item)
    return inventory


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path, nargs='?')
    parser.add_argument('version', nargs='?')
    parser.add_argument('--verify-manifest', action='store_true')
    parser.add_argument('--notices-only', action='store_true')
    parser.add_argument('--update-dependency-baseline', action='store_true')
    args = parser.parse_args()
    if args.update_dependency_baseline:
        inventory = resolved_dependency_inventory()
        DEPENDENCY_BASELINE.write_text(json.dumps(inventory, indent=2) + '\n')
        print(f'Wrote {len(inventory)} resolved package licenses to {DEPENDENCY_BASELINE}.')
        return
    if args.notices_only:
        verify_release_notices()
        print('Repository license, direct notices, and resolved dependency inventory validated.')
        return
    if args.directory is None or args.version is None:
        parser.error('directory and version are required unless --notices-only is used')
    result = inspect(args.directory, args.version)
    manifest = args.directory / 'package-manifest.json'
    if args.verify_manifest:
        if json.loads(manifest.read_text()) != result:
            raise ValueError('Package bytes changed after the consumer-test manifest was created.')
    else:
        manifest.write_text(json.dumps(result, indent=2) + '\n')
    print(f'Validated {len(result["packages"])} packages at {args.version}.')


if __name__ == '__main__':
    try:
        main()
    except (OSError, ValueError, zipfile.BadZipFile, ET.ParseError) as error:
        print(f'Package validation failed: {error}', file=sys.stderr)
        sys.exit(1)
