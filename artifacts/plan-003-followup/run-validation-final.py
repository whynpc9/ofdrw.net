from pathlib import Path
import subprocess,os,json,datetime,sys,time
root=Path.cwd(); out=root/'artifacts/plan-003-followup'
lock=Path('/Users/wanghongyi/Projects/ofdrw.net/.git/codex-followup-20261016/validation.lock')
while True:
 try:
  lock.mkdir(); break
 except FileExistsError:
  print('WAIT validation lease',flush=True); time.sleep(10)
owner={'task':'003','thread':'01a14281-a666-75f3-b005-a1e2518078c6','time':datetime.datetime.now(datetime.timezone.utc).isoformat(),'worktree':str(root),'operation':'full regression/docs/package gates'}
(lock/'owner.json').write_text(json.dumps(owner))
env=os.environ.copy(); env.update(DOTNET_CLI_HOME=str(out/'dotnet-home'),DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',DOTNET_CLI_TELEMETRY_OPTOUT='1',NUGET_PACKAGES='/Users/wanghongyi/.nuget/packages')
env.update(json.loads((out/'trusted-fonts.json').read_text())['environment'])
flags=['--disable-build-servers','-m:1','/nodeReuse:false','/p:UseSharedCompilation=false']
steps=[('restore',['dotnet','restore','Ofdrw.Net.sln',*flags]),('api-docs',['python3','scripts/check-public-api-docs.py']),('full-tests',['dotnet','test','Ofdrw.Net.sln','-c','Release','--no-build','--no-restore',*flags,'--logger','trx','--results-directory',str(out/'full-tests-final')]),('trusted-font-tests',['python3','scripts/run-trusted-font-tests.py','--font-output-dir',str(out/'trusted-fonts'),'--results-directory',str(out/'trusted-font-tests-final')]),('python-tests',['python3','-m','unittest','discover','-s','scripts/tests','-v']),('package-e2e',['bash','scripts/run-converter-package-e2e.sh','0.0.0-audit003final','--output-dir',str(out/'package-output-final')])]
results=[]
try:
 for name,command in steps:
  print('START',name,flush=True)
  with (out/(name+'-final.log')).open('w') as log: code=subprocess.call(command,env=env,stdout=log,stderr=subprocess.STDOUT)
  results.append({'name':name,'command':command,'exit_code':code})
  (out/'validation-commands-final.json').write_text(json.dumps(results,indent=2))
  print('END',name,code,flush=True)
  if code: sys.exit(code)
finally:
 (lock/'owner.json').unlink();lock.rmdir()
