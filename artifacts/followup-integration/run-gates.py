from pathlib import Path
import subprocess, os, json, datetime, shutil, xml.etree.ElementTree as ET

root = Path(__file__).resolve().parents[2]
out = root / 'artifacts/followup-integration'
shared = Path('/Users/wanghongyi/Projects/ofdrw.net/.git/codex-followup-20261016')
state = json.loads((shared / 'state.json').read_text())
assert state['stage'] == 'INTEGRATED_PENDING_VALIDATION'
for task in state['tasks'].values():
    assert task['ready_to_merge']
    subprocess.run(['git', 'merge-base', '--is-ancestor', task['worker_head'], 'HEAD'], cwd=root, check=True)
source = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=root, text=True).strip()
lock = shared / 'validation.lock'
lock.mkdir()
owner = {'task': 'integration', 'thread': state['parent_thread'], 'source': source,
         'time': datetime.datetime.now(datetime.timezone.utc).isoformat(), 'operation': 'combined source gates'}
(lock / 'owner.json').write_text(json.dumps(owner))
env = os.environ.copy()
env.update(DOTNET_CLI_HOME=str(out / 'dotnet-home'), DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',
           DOTNET_CLI_TELEMETRY_OPTOUT='1', NUGET_PACKAGES='/Users/wanghongyi/.nuget/packages')
flags = ['--disable-build-servers', '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false']
version = '0.0.0-followup.20261016.1'
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
        counter = ET.parse(file).find('.//t:Counters', ns).attrib
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
