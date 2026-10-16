from pathlib import Path
import os,json,subprocess,datetime,sys
root=Path.cwd();out=root/'artifacts/plan-028';lock=Path('/Users/wanghongyi/Projects/ofdrw.net/.git/codex-followup-20261016/validation.lock')
try: lock.mkdir()
except FileExistsError:
 print('BUSY', (lock/'owner.json').read_text() if (lock/'owner.json').exists() else 'owner pending');sys.exit(75)
(lock/'owner.json').write_text(json.dumps({'task':'028','thread':'01a14282-823e-70f2-a8b0-34a65983949f','time':datetime.datetime.now(datetime.timezone.utc).isoformat(),'operation':'local package build/consumption'}))
env=os.environ.copy();env.update(DOTNET_CLI_HOME=str(out/'dotnet-home'),DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',DOTNET_CLI_TELEMETRY_OPTOUT='1',NUGET_PACKAGES='/Users/wanghongyi/.nuget/packages');env.update(json.loads((out/'trusted-fonts.json').read_text())['environment'])
try:
 with (out/'package-e2e.log').open('w') as log: code=subprocess.call(['bash','scripts/run-converter-package-e2e.sh','0.0.0-audit028','--output-dir',str(out/'package-output')],env=env,stdout=log,stderr=subprocess.STDOUT)
 print('package_e2e_exit',code,flush=True)
 sys.exit(code)
finally:
 (lock/'owner.json').unlink();lock.rmdir()
