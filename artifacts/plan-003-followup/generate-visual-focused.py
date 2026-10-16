from pathlib import Path
import subprocess,os,json,datetime,sys,hashlib
root=Path.cwd();out=root/'artifacts/plan-003-followup';visual=out/'visual'
env=os.environ.copy();env.update(DOTNET_CLI_HOME=str(out/'dotnet-home'),DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',DOTNET_CLI_TELEMETRY_OPTOUT='1',NUGET_PACKAGES='/Users/wanghongyi/.nuget/packages');env.update(json.loads((out/'trusted-fonts.json').read_text())['environment'])
flags=['--disable-build-servers','-m:1','/nodeReuse:false','/p:UseSharedCompilation=false']
commands=[]
def run(name,cmd):
 print(name,flush=True)
 with (out/(name+'.log')).open('w') as log: subprocess.run(cmd,env=env,stdout=log,stderr=subprocess.STDOUT,check=True)
 commands.append({'name':name,'command':list(map(str,cmd))});(out/'visual-commands.json').write_text(json.dumps(commands,indent=2))
cli=root/'src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll'
source=root/'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx'
if True:
 run('current-cli-build',['dotnet','build',str(root/'src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj'),'-c','Release',*flags])
 baseline=out/'baseline-src'
 if not (baseline/'src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj').exists():
  baseline.mkdir(parents=True,exist_ok=True)
  archive=subprocess.Popen(['git','archive','3247c76f95c9a3312cc26144e24c4984767af26f'],stdout=subprocess.PIPE)
  subprocess.run(['tar','-x','-C',str(baseline)],stdin=archive.stdout,check=True)
  archive.stdout.close()
  if archive.wait()!=0: raise RuntimeError('Frozen source archive failed')
 fontdir=visual/'real-italic-fonts';fontdir.mkdir(exist_ok=True)
 for name,digest in [('Arial.ttf','525979822591a3447cfc49d943d6f7683508e25543407871c0ed8fed05fd2bd9'),('Arial Italic.ttf','ce1d2f1ab89db45f9796100eee960f5702a40e84c225c2b48c3ec3e81d153f98')]:
  physical=Path('/System/Library/Fonts/Supplemental')/name
  if hashlib.sha256(physical.read_bytes()).hexdigest()!=digest: raise RuntimeError('Recorded system face changed: '+name)
  target=fontdir/name
  if not target.exists(): target.symlink_to(physical)
  if hashlib.sha256(target.read_bytes()).hexdigest()!=digest: raise RuntimeError('Unexpected control face: '+name)
 run('baseline-build',['dotnet','build',str(baseline/'src/Ofdrw.Net.Cli/Ofdrw.Net.Cli.csproj'),'-c','Release',*flags])
 baseline_cli=baseline/'src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll'
 identity=[]
 for binary in [cli,cli.parent/'Ofdrw.Net.Converter.Docx.dll',baseline_cli,baseline_cli.parent/'Ofdrw.Net.Converter.Docx.dll']:
  identity.append({'path':str(binary),'sha256':hashlib.sha256(binary.read_bytes()).hexdigest(),'bytes':binary.stat().st_size})
 (out/'generator-binary-identity.json').write_text(json.dumps({'current_source_head':subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip(),'baseline_source_head':'3247c76f95c9a3312cc26144e24c4984767af26f','binaries':identity},indent=2))
 for mode,args in [('native2',['--docx-ofd-mode','native']),('default2',[]),('selected1',['--docx-ofd-mode','native','--pages','2'])]:
  directory=visual/mode;directory.mkdir(exist_ok=True)
  ofd=directory/(mode+'.ofd');pdf=directory/(mode+'.pdf')
  run(mode+'-ofd',['dotnet',str(cli),'docx-to-ofd',str(source),str(ofd),*args])
  run(mode+'-pdf',['dotnet',str(cli),'ofd-to-pdf',str(ofd),str(pdf)])
 for mode,exe,docx in [('directBuiltIn2',cli,source),('baseline-direct2',baseline_cli,source),('italic-boundaries',cli,visual/'italic-boundaries.docx')]:
  directory=visual/mode;directory.mkdir(exist_ok=True)
  run(mode,['dotnet',str(exe),'docx-to-pdf',str(docx),str(directory/(mode+'.pdf')),'--docx-engine','built-in',*(['--font-directory',str(visual/'real-italic-fonts')] if mode=='italic-boundaries' else [])])
 # Emit an unchanged physical italic control directly from the actual system font via the same DOCX path.
 for pdf in visual.rglob('*.pdf'):
  run(pdf.stem+'-render',['pdftoppm','-png','-r','144',str(pdf),str(pdf.parent/pdf.stem)])
  run(pdf.stem+'-text',['pdftotext','-layout',str(pdf),str(pdf.with_suffix('.txt'))])
 records=[]
 for p in sorted(visual.rglob('*')):
  if p.is_file():records.append({'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
 (out/'visual-manifest.json').write_text(json.dumps({'source_commit':subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip(),'base':'3247c76f95c9a3312cc26144e24c4984767af26f','files':records},indent=2))
