from pathlib import Path
import subprocess, os, json, datetime, sys
root=Path.cwd(); out=root/'artifacts/plan-028'
lock=Path('/Users/wanghongyi/Projects/ofdrw.net/.git/codex-followup-20261016/validation.lock')
try: lock.mkdir()
except FileExistsError:
 print('BUSY', (lock/'owner.json').read_text() if (lock/'owner.json').exists() else 'owner pending'); sys.exit(75)
(lock/'owner.json').write_text(json.dumps({'task':'028','thread':'plain-text implementation','time':datetime.datetime.now(datetime.timezone.utc).isoformat(),'worktree':str(root)}))
env=os.environ.copy(); env.update(DOTNET_CLI_HOME=str(out/'dotnet-home'),DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',DOTNET_CLI_TELEMETRY_OPTOUT='1',NUGET_PACKAGES='/Users/wanghongyi/.nuget/packages')
env.update(json.loads((out/'trusted-fonts.json').read_text())['environment'])
flags=['--disable-build-servers','-m:1','/nodeReuse:false','/p:UseSharedCompilation=false']
steps=[('api-docs',['python3','scripts/check-public-api-docs.py']),('full-tests',['dotnet','test','Ofdrw.Net.sln','-c','Release','--no-build','--no-restore',*flags,'--logger','trx','--results-directory',str(out/'full-tests')]),('trusted-font-tests',['python3','scripts/run-trusted-font-tests.py','--font-output-dir',str(out/'trusted-fonts'),'--results-directory',str(out/'trusted-font-tests')])]
results=[]
try:
 for name,command in steps:
  print('START',name,flush=True)
  with (out/(name+'.log')).open('w') as log: code=subprocess.call(command,env=env,stdout=log,stderr=subprocess.STDOUT)
  results.append({'name':name,'command':command,'exit_code':code})
  (out/'validation-commands.json').write_text(json.dumps(results,indent=2))
  print('END',name,code,flush=True)
  if code: sys.exit(code)
finally:
 (lock/'owner.json').unlink();lock.rmdir()
