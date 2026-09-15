#!/usr/bin/env python3
"""Validate the exact SDK/tool package set and bind consumer tests to its hashes."""
import argparse
import hashlib
import json
import re
from pathlib import Path
import sys
import xml.etree.ElementTree as ET
import zipfile

PACKAGES = (
    'Ofdrw.Net.Core', 'Ofdrw.Net.Packaging', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Reader',
    'Ofdrw.Net.Converter.Abstractions', 'Ofdrw.Net.Converter.Docx', 'Ofdrw.Net.Converter.Pdf',
    'Ofdrw.Net.Converter.Svg', 'Ofdrw.Net.Signatures', 'Ofdrw.Net.Converter', 'Ofdrw.Net.Cli',
)


def inspect(directory: Path, version: str):
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
            for dependency in root.iter():
                if dependency.tag.rsplit('}', 1)[-1] == 'dependency' and dependency.get('id', '').startswith('Ofdrw.Net.'):
                    if dependency.get('version') not in (version, f'[{version}]'):
                        raise ValueError(f'Mismatched SDK dependency in {path.name}: {dependency.attrib}')
            required = (['tools/net10.0/any/Ofdrw.Net.Cli.dll'] if name == 'Ofdrw.Net.Cli'
                        else [] if name == 'Ofdrw.Net.Converter'
                        else [f'lib/{framework}/{name}.dll' for framework in ('netstandard2.0', 'netstandard2.1')])
            required += ['README.md', 'THIRD-PARTY-NOTICES.md', 'docs/feature-parity.md', 'docs/conversion-contracts.md']
            for entry in required:
                if entry not in package.namelist():
                    raise ValueError(f'Missing {entry} in {path.name}')
        packages.append({'id': name, 'file': path.name, 'sha256': hashlib.sha256(path.read_bytes()).hexdigest()})
    return {'version': version, 'packages': packages}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('directory', type=Path)
    parser.add_argument('version')
    parser.add_argument('--verify-manifest', action='store_true')
    args = parser.parse_args()
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
