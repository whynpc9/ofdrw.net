"""Run the exact freshly consumed Core package on native Unix architectures.

Invoke only after the package/source gates succeed. Each container has no network,
a read-only app mount and a private executable tmpfs. Remote cleanup owns one
UUID directory and exact container names; evidence is copied back first.
"""
from pathlib import Path
import subprocess, os, json, shutil, hashlib, uuid, shlex, datetime

ROOT = Path(__file__).resolve().parents[2]
OUT = ROOT / 'artifacts/pr23-review-fixes'
APP = OUT / 'process-matrix-v2-app'
RESULT = OUT / 'process-matrix-v2'
RESULT.mkdir(exist_ok=True)
source = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=ROOT, text=True).strip()
assert json.loads((OUT / 'environment.json').read_text())['source'] == source
assert json.loads((OUT / 'gates.json').read_text())[-1]['exit_code'] == 0
if APP.exists():
    raise SystemExit('Use a new evidence directory for a new matrix; do not overwrite previous outputs')
shutil.copytree(OUT / 'process-consumer/bin/Release/net10.0', APP)
shutil.copytree(OUT / 'process-consumer/netstandard20', APP / 'netstandard20')
MUSL = 'mcr.microsoft.com/dotnet/runtime@sha256:e90a4b984926ed14328afec87e9be930b00dba55fed7e1743fbd585022afbad7'

def inspect(image, remote=False):
    command = ['docker', 'image', 'inspect', image]
    if remote:
        command = ['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(command)]
    return json.loads(subprocess.check_output(command, text=True))[0]

local_glibc = inspect('mcr.microsoft.com/dotnet/sdk:10.0')
remote_glibc = inspect('mcr.microsoft.com/dotnet/runtime:10.0.13-noble', True)
assert local_glibc['Architecture'] == 'arm64' and remote_glibc['Architecture'] == 'amd64'
GLIBC_ARM = local_glibc['RepoDigests'][0]
GLIBC_X64 = remote_glibc['RepoDigests'][0]
nonce = uuid.uuid4().hex[:12]
remote = '/tmp/ofdrw-pr23-package-' + source[:12] + '-' + nonce
records = []
base_env = dict(os.environ, DOTNET_CLI_HOME=str(OUT / 'dotnet-home'), DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1', DOTNET_CLI_TELEMETRY_OPTOUT='1')


def run_cell(name, command, destination, remote_command=False):
    destination.mkdir(parents=True, exist_ok=True)
    started = datetime.datetime.now(datetime.timezone.utc).isoformat()
    try:
        process = subprocess.run(command, env=base_env, cwd=ROOT, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True, timeout=75)
        (destination / 'console.log').write_text(process.stdout)
        record = {'name': name, 'source': source, 'command': command, 'started': started, 'exit_code': process.returncode}
        records.append(record)
        print(name, process.returncode, flush=True)
    except subprocess.TimeoutExpired as error:
        records.append({'name': name, 'source': source, 'command': command, 'started': started, 'timeout': str(error)})
        raise
    finally:
        (RESULT / 'cells.json').write_text(json.dumps(records, indent=2) + '\n')


def container_command(image, arch, mode, app, output, name, tfm):
    command = ['docker', 'run', '--rm', '--name', name, '--platform', 'linux/' + arch,
               '--network', 'none', '--read-only', '--tmpfs', '/tmp:exec,mode=1777']
    if mode != 'pid1':
        command.append('--init')
    for value in ['DOTNET_CLI_HOME=/tmp/dotnet-home', 'DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1', 'DOTNET_CLI_TELEMETRY_OPTOUT=1', 'DOTNET_SYSTEM_GLOBALIZATION_INVARIANT=1']:
        command += ['-e', value]
    command += ['-v', app + ':/app:ro', '-v', output + ':/output', image, 'dotnet', '/app/ProcessConsumer.dll', mode, '/output']
    if tfm == 'net20':
        command += ['/app/netstandard20/Ofdrw.Net.Core.dll']
    return command


for architecture, executable in [('arm64', '/usr/local/share/dotnet/dotnet'), ('x64-rosetta', '/Users/wanghongyi/Projects/ofdrw.net/.git/codex-pr-babysit-20261016/runtime-osx-x64-10.0.13/dotnet')]:
    for tfm in ['net21', 'net20']:
        name = 'macos-' + architecture + '-' + tfm
        destination = RESULT / name
        command = [executable, str(APP / 'ProcessConsumer.dll'), 'default', str(destination)]
        if tfm == 'net20': command += [str(APP / 'netstandard20/Ofdrw.Net.Core.dll')]
        run_cell(name, command, destination)

for libc, image in [('glibc', GLIBC_ARM), ('musl', MUSL)]:
    for tfm in ['net21', 'net20']:
        for mode in ['default', 'pid1', 'ignore']:
            name = 'linux-arm64-' + libc + '-' + tfm + '-' + mode
            destination = RESULT / name; destination.mkdir()
            container = 'ofd-pr23-' + nonce + '-' + name
            command = container_command(image, 'arm64', mode, str(APP), str(destination), container, tfm)
            try: run_cell(name, command, destination)
            finally: subprocess.run(['docker', 'rm', '-f', container], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)

subprocess.run(['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(['mkdir', '-p', remote + '/app', remote + '/outputs'])], check=True)
subprocess.run(['scp', '-r', str(APP) + '/.', 'legion:' + remote + '/app/'], check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
try:
    for libc, image in [('glibc', GLIBC_X64), ('musl', MUSL)]:
        for tfm in ['net21', 'net20']:
            for mode in ['default', 'pid1', 'ignore']:
                name = 'linux-x64-' + libc + '-' + tfm + '-' + mode
                path = remote + '/outputs/' + name
                subprocess.run(['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(['mkdir', path])], check=True)
                container = 'ofd-pr23-' + nonce + '-' + name
                command = container_command(image, 'amd64', mode, remote + '/app', path, container, tfm)
                try: run_cell(name, ['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(command)], RESULT / name)
                finally: subprocess.run(['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(['docker', 'rm', '-f', container])], stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
finally:
    # A failed copy retains the uniquely owned remote directory for recovery.
    import tarfile
    archive_remote = ['docker', 'run', '--rm', '--network', 'none', '--read-only', '-v', remote + ':/owned', MUSL,
                      '/bin/sh', '-c', 'tar -czf /owned/remote-evidence.tar.gz -C /owned outputs && chmod 644 /owned/remote-evidence.tar.gz']
    subprocess.run(['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(archive_remote)], check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    archive = RESULT / 'remote-evidence.tar.gz'
    subprocess.run(['scp', 'legion:' + remote + '/remote-evidence.tar.gz', str(archive)], check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    with tarfile.open(archive) as data:
        data.extractall(RESULT / 'remote-raw', filter='data')
    raw = RESULT / 'remote-raw/outputs'
    assert len(list(raw.glob('*/results.json'))) == 12
    for path in raw.glob('*'):
        for name in ['results.json', 'descendant-observations.json']:
            shutil.copyfile(path / name, RESULT / path.name / name)
    clean_remote = ['docker', 'run', '--rm', '--network', 'none', '--read-only', '-v', remote + ':/owned', MUSL,
                    '/bin/sh', '-c', 'rm -rf /owned/outputs /owned/app /owned/remote-evidence.tar.gz']
    subprocess.run(['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(clean_remote)], check=True, stdout=subprocess.PIPE, stderr=subprocess.STDOUT)
    subprocess.run(['ssh', '-o', 'BatchMode=yes', 'legion', shlex.join(['rmdir', remote])], check=True)

assert len(records) == 28
for cell in records:
    assert cell.get('exit_code') == 0, cell
    data = json.loads((RESULT / cell['name'] / 'results.json').read_text())
    assert len(data['cases']) == 9 and all(case['outcome'] == 'Passed' for case in data['cases']), cell['name']
    assert data['architecture'].lower() == ('x64' if 'x64' in cell['name'] else 'arm64')
    if cell['name'].endswith('-pid1'): assert data['pid'] == 1
summary = {'source': source, 'cells': 28, 'passed_cases': 252, 'core_netstandard': ['2.0','2.1'], 'notes': 'macOS x64 through Rosetta; Linux both architectures are native. Both Linux libc hosts run default/nonPID1, actual PID1 and explicit inherited SIGCHLD ignore.', 'source_inputs': {str(p.relative_to(APP)): hashlib.sha256(p.read_bytes()).hexdigest() for p in APP.rglob('*') if p.is_file()}, 'glibc_arm64_image': GLIBC_ARM, 'glibc_x64_image': GLIBC_X64, 'musl_image': MUSL}
(RESULT / 'summary.json').write_text(json.dumps(summary, indent=2) + '\n')
print('PACKAGE_MATRIX_PASS', summary['passed_cases'], flush=True)
