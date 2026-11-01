#!/usr/bin/env python3
"""Run the complete SDK/tool package consumer using native processes on each OS.

--visual-corpus-only limits execution to the four visual fixtures. Capture mode
never writes reviewed goldens and is explicitly reported as unvalidated.
"""
import argparse
import importlib.util
import json
import hashlib
import zipfile
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parents[1]
FLAGS = ['--disable-build-servers', '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false', '--nologo']


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('version', nargs='?', default='0.1.0-preview.8')
    parser.add_argument('--consume-only', action='store_true')
    parser.add_argument('--packages-dir', type=Path)
    parser.add_argument('--output-dir', type=Path)
    parser.add_argument('--visual-corpus-only', action='store_true')
    parser.add_argument('--capture-visual-baseline', action='store_true')
    options = parser.parse_args()
    if not re.fullmatch(r'\d+\.\d+\.\d+(?:-[0-9A-Za-z-]+(?:\.[0-9A-Za-z-]+)*)?', options.version):
        parser.error('Expected a SemVer package version.')
    if options.capture_visual_baseline and not options.visual_corpus_only:
        parser.error('--capture-visual-baseline requires --visual-corpus-only')
    feed = (options.packages_dir or ROOT / 'artifacts/package-e2e' / options.version / 'packages').resolve()
    output = (options.output_dir or ROOT / 'artifacts/package-e2e' / options.version / 'output').resolve()
    output.mkdir(parents=True, exist_ok=True)
    feed.mkdir(parents=True, exist_ok=True)
    spec = importlib.util.spec_from_file_location('package_validator', ROOT / 'scripts/verify-package-artifacts.py')
    validator = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(validator)
    commands = []
    summary = {'status': 'RUNNING', 'platform': sys.platform, 'version': options.version,
               'mode': 'visual-capture' if options.capture_visual_baseline else
                       'visual-only' if options.visual_corpus_only else 'full-package-e2e',
               'sourceSha': subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip(),
               'sourceChanges': subprocess.check_output(['git', 'status', '--porcelain'], cwd=ROOT, text=True).splitlines(),
               'packages': str(feed), 'commands': commands}
    with tempfile.TemporaryDirectory(prefix='ofdrw-package-e2e-') as temporary:
        task = Path(temporary)
        env = os.environ.copy()
        cache = Path(env.get('NUGET_PACKAGES', Path.home() / '.nuget/packages')).resolve()
        cache.mkdir(parents=True, exist_ok=True)
        env.update(DOTNET_CLI_HOME=str(task / 'dotnet-home'), DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',
                   DOTNET_CLI_TELEMETRY_OPTOUT='1', DOTNET_CLI_DO_NOT_USE_MSBUILD_SERVER='1',
                   NUGET_PACKAGES=str(cache), OFDRW_REPO_ROOT=str(ROOT), OFDRW_E2E_OUTPUT_DIR=str(output),
                   OFDRW_SOURCE_SHA=summary['sourceSha'])
        dotnet = env.get('OFDRW_DOTNET') or shutil.which('dotnet')
        if not dotnet:
            raise RuntimeError('dotnet is required')

        def run(command):
            command = list(map(str, command))
            number = len(commands) + 1
            log = output / f'command-{number:02d}.log'
            record = {'argv': command, 'log': log.name}
            commands.append(record)
            print(f'[E2E {number}] {subprocess.list2cmdline(command)}', flush=True)
            start = time.monotonic()
            with log.open('wb') as stream:
                result = subprocess.run(command, cwd=ROOT, env=env, stdout=stream,
                                        stderr=subprocess.STDOUT, timeout=1800)
            record.update(exitCode=result.returncode, seconds=round(time.monotonic() - start, 2))
            if result.returncode:
                print(log.read_text(encoding='utf-8', errors='replace')[-12000:], file=sys.stderr)
                raise RuntimeError(f'Command {number} failed ({result.returncode}); see {log}')

        def dn(*args):
            run([dotnet, *args])

        def consumer(source, name):
            directory = task / name
            directory.mkdir()
            shutil.copyfile(ROOT / f'e2e/{source}/Program.cs', directory / 'Program.cs')
            project = directory / 'Consumer.csproj'
            if source == 'Ofdrw.Net.Text.HarfBuzz.E2E':
                shutil.copyfile(ROOT / f'e2e/{source}/PositionedRunProbe.cs', directory / 'PositionedRunProbe.cs')
                shutil.copyfile(ROOT / f'e2e/{source}/SizedFontSpanProbe.cs', directory / 'SizedFontSpanProbe.cs')
                shutil.copyfile(ROOT / f'e2e/{source}/PositionedConsumerProbe.cs', directory / 'PositionedConsumerProbe.cs')
                items = ('Text.HarfBuzz', 'Packaging', 'Reader', 'Converter.Pdf', 'Converter.Svg')
                references = '\n'.join(f'    <PackageReference Include="Ofdrw.Net.{item}" Version="{options.version}" />'
                                       for item in items)
                project.write_text('<Project Sdk="Microsoft.NET.Sdk">\n'
                                   '  <PropertyGroup><OutputType>Exe</OutputType><TargetFramework>net10.0</TargetFramework>'
                                   '<ImplicitUsings>enable</ImplicitUsings></PropertyGroup>\n'
                                   f'  <ItemGroup>\n{references}\n  </ItemGroup>\n</Project>\n', encoding='utf-8')
            else:
                shutil.copyfile(ROOT / 'e2e/Ofdrw.Net.Converter.Pdf.E2E/Ofdrw.Net.Converter.Pdf.E2E.csproj', project)
            dn('restore', project, '--configfile', config, f'-p:OfdrwPackageVersion={options.version}', *FLAGS)
            dn('build', project, '-c', 'Release', '--no-restore', f'-p:OfdrwPackageVersion={options.version}', *FLAGS)
            return project

        def verify_runtime_identity(project, receipt_name, required):
            bin_dir = project.parent / 'bin/Release/net10.0'
            deps = json.loads((bin_dir / 'Consumer.deps.json').read_text(encoding='utf-8'))
            target = deps['targets'][deps['runtimeTarget']['name']]
            manifest = {item['id']: item for item in
                        json.loads((feed / 'package-manifest.json').read_text(encoding='utf-8'))['packages']}
            items = []
            for key, node in sorted(target.items()):
                if not key.startswith('Ofdrw.Net.'):
                    continue
                package_id, resolved_version = key.split('/', 1)
                if resolved_version != options.version or package_id not in manifest:
                    raise RuntimeError(f'Unexpected resolved Ofdrw package: {key}')
                assets = [path for path in node.get('runtime', {}) if path.endswith('.dll')]
                if not assets:
                    continue  # Meta-packages have no selected runtime assembly.
                if len(assets) != 1:
                    raise RuntimeError(f'Expected one selected runtime DLL for {key}: {assets}')
                asset = assets[0]
                nupkg = feed / manifest[package_id]['file']
                if hashlib.sha256(nupkg.read_bytes()).hexdigest() != manifest[package_id]['sha256']:
                    raise RuntimeError(f'Candidate package changed: {nupkg}')
                with zipfile.ZipFile(nupkg) as archive:
                    payload = archive.read(asset)
                expected = hashlib.sha256(payload).hexdigest()
                extracted = task / 'packages' / package_id.lower() / options.version / asset
                copied = bin_dir / Path(asset).name
                if any(hashlib.sha256(path.read_bytes()).hexdigest() != expected for path in (extracted, copied)):
                    raise RuntimeError(f'Selected runtime DLL differs from candidate: {key} {asset}')
                items.append({'id': package_id, 'version': options.version, 'asset': asset,
                              'nupkgSha256': manifest[package_id]['sha256'], 'dllSha256': expected})
            if not required <= {item['id'] for item in items}:
                raise RuntimeError('Candidate runtime DLL identity set is incomplete')
            (output / receipt_name).write_text(json.dumps({
                'sourceSha': summary['sourceSha'], 'runtimeTarget': deps['runtimeTarget']['name'],
                'packages': items}, indent=2) + '\n', encoding='utf-8')

        try:
            dn('--info')
            if not options.consume_only:
                if any(feed.glob('*.nupkg')):
                    raise RuntimeError('Use a fresh feed directory, or --consume-only for immutable candidate packages.')
                dn('restore', ROOT / 'Ofdrw.Net.sln', *FLAGS)
                dn('build', ROOT / 'Ofdrw.Net.sln', '-c', 'Release', '--no-restore',
                   f'-p:Version={options.version}', f'-p:PackageVersion={options.version}', *FLAGS)
                for product in validator.PACKAGES:
                    dn('pack', ROOT / f'src/{product}/{product}.csproj', '-c', 'Release', '--no-build',
                       '--no-restore', '-o', feed, f'-p:Version={options.version}',
                       f'-p:PackageVersion={options.version}', *FLAGS)
            else:
                # License validation needs restored source dependency metadata, even
                # when the product DLLs come exclusively from a supplied feed.
                dn('restore', ROOT / 'Ofdrw.Net.sln', *FLAGS)
            # Existing manifest is immutable when consuming an already verified feed.
            validation = [sys.executable, ROOT / 'scripts/verify-package-artifacts.py', feed, options.version]
            run(validation + (['--verify-manifest'] if (feed / 'package-manifest.json').exists() else []))
            config = task / 'NuGet.Config'
            configuration = ET.Element('configuration')
            sources = ET.SubElement(configuration, 'packageSources')
            ET.SubElement(sources, 'clear')
            mapping = ET.SubElement(configuration, 'packageSourceMapping')
            for key, location, pattern in [('local-artifacts', str(feed), 'Ofdrw.Net.*'),
                                           ('cached-third-party', str(cache), '*'),
                                           ('nuget.org', 'https://api.nuget.org/v3/index.json', '*')]:
                ET.SubElement(sources, 'add', key=key, value=location)
                ET.SubElement(ET.SubElement(mapping, 'packageSource', key=key), 'package', pattern=pattern)
            ET.ElementTree(configuration).write(config, encoding='utf-8', xml_declaration=True)
            env['NUGET_PACKAGES'] = str(task / 'packages')
            project = consumer('Ofdrw.Net.Converter.Pdf.E2E', 'consumer')
            verify_runtime_identity(project, 'runtime-identities-main.json',
                                    {'Ofdrw.Net.Core', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Converter.Pdf'})
            if not options.visual_corpus_only:
                dn('tool', 'install', 'Ofdrw.Net.Cli', '--tool-path', task / 'tools', '--version',
                   options.version, '--configfile', config)
                cli = task / 'tools' / ('ofdrw.exe' if os.name == 'nt' else 'ofdrw')
                text = output / 'cli-plain-text.txt'
                text.write_bytes('\ufeff中英 iii WWW\r\n\r\n尾行\n'.encode('utf-8'))
                run([cli, 'text-to-ofd', text, output / 'cli-plain-text.ofd'])
                extracted = output / 'cli-plain-text-extracted.txt'
                run([cli, 'extract-text', output / 'cli-plain-text.ofd', extracted])
                actual = extracted.read_text(encoding='utf-8-sig')
                if actual.count('中英 iii WWW') != 1 or actual.count('尾行') != 1:
                    raise RuntimeError('Installed CLI changed or duplicated source text')
                docx_fonts = []
                if os.name == 'nt':
                    # Match the API consumer's explicit Noto directory. The
                    # fixture requests Noto; it need not be installed globally.
                    archive = ROOT / 'tests/fixtures/trusted-fonts/sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip'
                    with zipfile.ZipFile(archive) as bundle:
                        font = bundle.read('fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf')
                    if len(font) != 21659544 or hashlib.sha256(font).hexdigest() != '3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa':
                        raise RuntimeError('Trusted DOCX font hash mismatch')
                    directory = output / 'fixture-fonts'
                    directory.mkdir(exist_ok=True)
                    (directory / 'NotoSansCJKsc-Regular.ttf').write_bytes(font)
                    docx_fonts = ['--font-directory', directory]
                run([cli, 'docx-to-ofd', ROOT / 'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx', output / 'cli-native.ofd', *docx_fonts])
                run([cli, 'ofd-to-image', output / 'cli-native.ofd', output / 'cli-image.png', '--pages', '1', '--ppm', '2'])
                run([cli, 'image-to-ofd', output / 'cli-image.png', output / 'cli-image.ofd', '--ppm', '2', '--page-width', '220', '--page-height', '307'])
            arguments = ['--visual-corpus-only'] if options.visual_corpus_only else []
            if options.capture_visual_baseline:
                arguments.append('--capture-visual-baseline')
            dn('run', '--project', project, '-c', 'Release', '--no-build', '--no-restore', '--', *arguments)
            if not options.visual_corpus_only:
                font_project = consumer('Ofdrw.Net.FontSubset.E2E', 'font-consumer')
                dn('run', '--project', font_project, '-c', 'Release', '--no-build', '--no-restore', '--', ROOT, output / 'font-subset')
                text_project = consumer('Ofdrw.Net.Text.HarfBuzz.E2E', 'text-consumer')
                verify_runtime_identity(text_project, 'runtime-identities-text.json',
                                        {'Ofdrw.Net.Core', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Text.HarfBuzz',
                                         'Ofdrw.Net.Packaging', 'Ofdrw.Net.Reader', 'Ofdrw.Net.Converter.Pdf',
                                         'Ofdrw.Net.Converter.Svg'})
                env['R21_SOURCE_SHA'] = summary['sourceSha']
                dn('run', '--project', text_project, '-c', 'Release', '--no-build', '--no-restore', '--',
                   output / 'text-prelayout')
            env['NUGET_PACKAGES'] = str(cache)
            run(validation + ['--verify-manifest'])
            shutil.copyfile(feed / 'package-manifest.json', output / 'package-manifest.json')
            summary['status'] = 'CAPTURE_ONLY_NOT_VALIDATED' if options.capture_visual_baseline else 'PASS'
        except Exception as failure:
            summary.update(status='FAIL', error=str(failure))
            raise
        finally:
            (output / 'package-e2e-run.json').write_text(json.dumps(summary, indent=2), encoding='utf-8')
    print(f"[E2E] {summary['status']}: {output}")


if __name__ == '__main__':
    main()
