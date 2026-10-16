from pathlib import Path
import subprocess, os, json, datetime, shutil, xml.etree.ElementTree as ET

root = Path(__file__).resolve().parents[2]
out = root / 'artifacts/pr23-review-fixes'
shared = Path('/Users/wanghongyi/Projects/ofdrw.net/.git/codex-followup-20261016')
review_shared = Path('/Users/wanghongyi/Projects/ofdrw.net/.git/codex-pr-babysit-20261016')
state = json.loads((review_shared / 'state.json').read_text())
assert state['stage'] == 'REVIEW_FIXES_READY_FOR_VALIDATION'
for finding in state['findings'].values():
    assert finding['status'] == 'ROOT_ACCEPTED'
    subprocess.run(['git', 'merge-base', '--is-ancestor', finding['fix_head'], 'HEAD'], cwd=root, check=True)
source = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=root, text=True).strip()
lock = shared / 'validation.lock'
lock.mkdir()
owner = {'task': 'pr23-review-integration', 'thread': '01a13e30-bd27-77a1-8bf8-1c2d10ac0717', 'source': source,
         'time': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'operation': 'combined source gates'}
(lock / 'owner.json').write_text(json.dumps(owner))
env = os.environ.copy()
env.update(DOTNET_CLI_HOME=str(out / 'dotnet-home'), DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',
           DOTNET_CLI_TELEMETRY_OPTOUT='1', NUGET_PACKAGES='/Users/wanghongyi/.nuget/packages')
flags = ['--disable-build-servers', '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false']
version = '0.0.0-pr23.20261016.1'
results = []

def run(name, command, run_env=None):
    print('START', name, flush=True)
    entry = {'name': name, 'command': command, 'source': source,
             'started': datetime.datetime.now(datetime.timezone.utc).isoformat()}
    with (out / (name + '.log')).open('w') as log:
        entry['exit_code'] = subprocess.run(command, cwd=root, env=run_env or env,
                                             stdout=log, stderr=subprocess.STDOUT).returncode
    results.append(entry)
    (out / 'gates.json').write_text(json.dumps(results, indent=2) + '\n')
    print('END', name, entry['exit_code'], flush=True)
    if entry['exit_code']:
        raise SystemExit(entry['exit_code'])

try:
    run('prepare-fonts', ['python3', 'scripts/prepare-trusted-fonts.py', '--output-dir', str(out / 'trusted-fonts')])
    fonts = json.loads((out / 'prepare-fonts.log').read_text())
    env.update(fonts['environment'])
    (out / 'environment.json').write_text(json.dumps({'source': source, 'version': version,
        'environment': {k: env[k] for k in ['DOTNET_CLI_HOME', 'DOTNET_SKIP_FIRST_TIME_EXPERIENCE',
        'DOTNET_CLI_TELEMETRY_OPTOUT', 'NUGET_PACKAGES', *fonts['environment'].keys()] }}, indent=2) + '\n')
    run('restore', ['dotnet', 'restore', 'Ofdrw.Net.sln', *flags])
    run('api-docs-rebuild', ['python3', 'scripts/check-public-api-docs.py'])
    run('full-tests', ['dotnet', 'test', 'Ofdrw.Net.sln', '-c', 'Release', '--no-build', '--no-restore',
                      '--logger', 'trx', '--results-directory', str(out / 'full-tests'), *flags])
    ns = {'t': 'http://microsoft.com/schemas/VisualStudio/TeamTest/2010'}
    counts = {k: 0 for k in ['total', 'passed', 'failed', 'notExecuted']}
    files = list((out / 'full-tests').glob('*.trx'))
    assert len(files) == 11
    for file in files:
        doc = ET.parse(file)
        counter = doc.find('.//t:Counters', ns).attrib
        assert all(result.attrib.get('outcome') == 'Passed' for result in doc.findall('.//t:UnitTestResult', ns)), str(file)
        for key in counts:
            counts[key] += int(counter[key])
    assert counts['total'] > 0 and counts['total'] == counts['passed'] and counts['failed'] == counts['notExecuted'] == 0
    (out / 'test-summary.json').write_text(json.dumps({'source': source, 'counts': counts, 'projects': len(files)}, indent=2) + '\n')
    run('trusted-font-tests', ['python3', 'scripts/run-trusted-font-tests.py', '--font-output-dir', str(out / 'trusted-fonts'),
                             '--results-directory', str(out / 'trusted-font-tests')])
    run('python-tests', ['python3', '-m', 'unittest', 'discover', '-s', 'scripts/tests', '-v'])
    run('package11', ['bash', 'scripts/run-converter-package-e2e.sh', version, '--packages-dir', str(out / 'packages'),
                      '--output-dir', str(out / 'package-output')])
    run('html-package', ['bash', 'scripts/run-html-preview-package-e2e.sh', version,
                         str(out / 'packages'), str(out / 'html-package-output')])
    process_consumer = out / 'process-consumer'
    process_config = ET.Element('configuration')
    process_sources = ET.SubElement(process_config, 'packageSources'); ET.SubElement(process_sources, 'clear')
    for key, value in [('candidate', str(out / 'packages')), ('cache', env['NUGET_PACKAGES'])]:
        ET.SubElement(process_sources, 'add', key=key, value=value)
    process_mapping = ET.SubElement(process_config, 'packageSourceMapping')
    for key, pattern in [('candidate', 'Ofdrw.Net.*'), ('cache', '*')]:
        ET.SubElement(ET.SubElement(process_mapping, 'packageSource', key=key), 'package', pattern=pattern)
    ET.ElementTree(process_config).write(process_consumer / 'NuGet.Config', encoding='utf-8', xml_declaration=True)
    process_env = dict(env, NUGET_PACKAGES=str(out / 'process-consumer-cache'))
    process_project = str(process_consumer / 'ProcessConsumer.csproj')
    run('process-consumer-restore', ['dotnet', 'restore', process_project, '--configfile', str(process_consumer / 'NuGet.Config'),
                                    '-p:OfdrwPackageVersion=' + version, *flags], process_env)
    run('process-consumer-build', ['dotnet', 'build', process_project, '-c', 'Release', '--no-restore',
                                  '-p:OfdrwPackageVersion=' + version, *flags], process_env)
    process_dll = str(process_consumer / 'bin/Release/net10.0/ProcessConsumer.dll')
    run('process-package-consume', ['dotnet', process_dll, 'default', str(out / 'process-package-output')], process_env)
    import zipfile
    process_net20 = process_consumer / 'netstandard20'
    process_net20.mkdir(exist_ok=True)
    core_package = out / 'packages' / ('Ofdrw.Net.Core.' + version + '.nupkg')
    with zipfile.ZipFile(core_package) as archive:
        (process_net20 / 'Ofdrw.Net.Core.dll').write_bytes(archive.read('lib/netstandard2.0/Ofdrw.Net.Core.dll'))
    run('process-package-consume-net20', ['dotnet', process_dll, 'default', str(out / 'process-package-output-net20'),
                                        str(process_net20 / 'Ofdrw.Net.Core.dll')], process_env)
    consumer = out / 'preflight-consumer'
    consumer.mkdir(exist_ok=True)
    for name in ['Program.cs', 'Ofdrw.Net.Packaging.Preflight.E2E.csproj']:
        shutil.copyfile(root / 'e2e/Ofdrw.Net.Packaging.Preflight.E2E' / name, consumer / name)
    config = ET.Element('configuration')
    sources = ET.SubElement(config, 'packageSources'); ET.SubElement(sources, 'clear')
    for key, value in [('candidate', str(out / 'packages')), ('cache', env['NUGET_PACKAGES'])]:
        ET.SubElement(sources, 'add', key=key, value=value)
    mapping = ET.SubElement(config, 'packageSourceMapping')
    for key, pattern in [('candidate', 'Ofdrw.Net.*'), ('cache', '*')]:
        ET.SubElement(ET.SubElement(mapping, 'packageSource', key=key), 'package', pattern=pattern)
    ET.ElementTree(config).write(consumer / 'NuGet.Config', encoding='utf-8', xml_declaration=True)
    clean_env = dict(env, NUGET_PACKAGES=str(out / 'preflight-cache'))
    project = str(consumer / 'Ofdrw.Net.Packaging.Preflight.E2E.csproj')
    run('preflight-restore', ['dotnet', 'restore', project, '--configfile', str(consumer / 'NuGet.Config'),
                             '-p:OfdrwPackageVersion=' + version, *flags], clean_env)
    run('preflight-build', ['dotnet', 'build', project, '-c', 'Release', '--no-restore',
                           '-p:OfdrwPackageVersion=' + version, *flags], clean_env)
    run('preflight-consume', ['dotnet', 'run', '--project', project, '-c', 'Release', '--no-build', '--no-restore',
                             '--', str(out / 'preflight-output')], clean_env)
    print('FUNCTIONAL_GATES_PASS', source, counts, flush=True)
finally:
    assert json.loads((lock / 'owner.json').read_text()) == owner
    (lock / 'owner.json').unlink(); lock.rmdir()
    print('VALIDATION_LEASE_RELEASED', flush=True)
