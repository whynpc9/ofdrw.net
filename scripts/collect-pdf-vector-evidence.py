#!/usr/bin/env python3
"""Persist ticket21 source-bound artifacts, including failures; never infer visual acceptance."""
from pathlib import Path
import argparse, hashlib, json, shutil, subprocess, tarfile, tempfile

root=Path(__file__).resolve().parents[1]
parser=argparse.ArgumentParser(description=__doc__)
parser.add_argument('runtime_head');parser.add_argument('version');parser.add_argument('--round',default='r2');parser.add_argument('--low',default='low3');parser.add_argument('--regression',default='r3')
args=parser.parse_args(); destination=root/'docs/evidence/pdf-vectors'; destination.mkdir(parents=True,exist_ok=True)
tracked=[str(p.relative_to(root)) for folder in ('src/Ofdrw.Net.Converter.Pdf.Vector','src/Ofdrw.Net.Core','src/Ofdrw.Net.Converter.Pdf','src/Ofdrw.Net.Converter.Svg','src/Ofdrw.Net.Graphics.SkiaSharp','e2e/Ofdrw.Net.Pdf.Vector.E2E','tests/Ofdrw.Net.Converter.Pdf.Vector.Tests') for p in (root/folder).rglob('*') if p.is_file() and 'bin' not in p.parts and 'obj' not in p.parts]
if subprocess.check_output(['git','diff',args.runtime_head,'--',*tracked],cwd=root):raise ValueError('Runtime/sample changed since freeze; regenerate gates first')
sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
excluded={'bin','obj','packages','feed','dotnet-home','dotnet-home-low2','__pycache__','native','node_modules','venv','.venv'}
with tempfile.TemporaryDirectory(prefix='ofd-vector-evidence-') as temporary:
 stage=Path(temporary)
 def copy(source,target):
  source=Path(source);source=source if source.is_absolute() else root/source;target=stage/target;target.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(source,target)
 def tree(source,target):
  source=root/source
  for p in sorted(source.rglob('*')):
   rel=p.relative_to(source)
   if p.is_file() and not any(part in excluded for part in rel.parts):copy(p,Path(target)/rel)
 candidate=f'artifacts/pdf-vectors/package-{args.round}'
 tree(candidate+'/output','candidate')
 copy(candidate+'/package-manifest.json','packages/primary-optional.json');copy(candidate+'/consumer.assets.json','packages/primary-consumer.assets.json')
 for product in ('Ofdrw.Net.Converter.Pdf.Vector','Ofdrw.Net.Graphics.SkiaSharp'):
  copy(candidate+f'/feed/{product}.{args.version}.nupkg',f'packages/{product}.{args.version}.nupkg')
 copy(f'artifacts/package-e2e/{args.version}/packages/package-manifest.json','packages/primary-default.json')
 tree('artifacts/pdf-vectors/probe','design-probe')
 tree('artifacts/pdf-vectors/image-fallback-design','image-design')
 tree('artifacts/pdf-vectors/interpolation-diagnostic','interpolation-diagnostic')
 copy('artifacts/pdf-vectors/interpolation-experiment.patch','interpolation-diagnostic/reverted-fixture-experiment.patch')
 tree('artifacts/pdf-vectors/image-page-probe','image-before-repair')
 tree('artifacts/pdf-vectors/golden-r2-observer','golden-original-source-repair')
 # The first round's actual Preview failure stays immutable and independently reviewable.
 for name in ('fallback-source.pdf','fallback.pdf','fallback.ofd','fallback-source-3.png','fallback-3.png','sample-report.json','verified-evidence.json'):
  copy('artifacts/pdf-vectors/package-r1/output/'+name,'previous-r1/'+name)
 for p in sorted((root/'artifacts/pdf-vectors').glob('*.log')):copy(p,'logs/'+p.name)
 for folder in ('full-r2','full-image-r2-final'):
  tree('artifacts/pdf-vectors/'+folder,'logs/'+folder)
 for filename in ('environment.json','preview-r1.json',f'preview-{args.round}.json'):
  copy('artifacts/pdf-vectors/'+filename,'acceptance/'+filename)
 regression=root/f'artifacts/pdf-vectors/graphics-regression-{args.regression}'
 for prefix in ('baseline-native','baseline-default'):
  for p in regression.glob(prefix+'*'):
   if p.is_file():copy(p,'docx-regression/'+p.name)
  for p in (regression/'pages').glob(prefix+'*'):copy(p,'docx-regression/pages/'+p.name)
 copy(regression/'licensed-layout.docx','docx-regression/licensed-layout.docx');copy(regression/'manifest.json','docx-regression/manifest.json')
 for p in (root/'artifacts/graphics-fonts').iterdir():
  if p.is_file():copy(p,'licenses-and-fonts/'+p.name)
 independent=root/'artifacts/pdf-vectors/independent'
 for p in independent.glob('verification-*.json'):copy(p,'independent/'+p.name)
 for folder in ('logs','logs-low2',f'logs-{args.low}',f'observer-{args.low}'):
  if (independent/folder).exists():tree(str((independent/folder).relative_to(root)),'independent/'+folder)
 source=independent/f'source-{args.low}'
 optional=independent/f'optional13-{args.low}'
 for filename in ('package-manifest.json','consumer.assets.json'):
  copy(optional/filename,'independent/'+filename)
 for filename in ('sample-report.json','image-hint-report.json','verified-evidence.json'):
  copy(optional/'output'/filename,'independent/'+filename)
 copy(independent/f'default-feed-{args.low}'/'package-manifest.json','independent/default-packages.json')
 copy(root/'LICENSE','licenses-and-fonts/repository-MIT.txt');copy(root/'THIRD-PARTY-NOTICES.md','licenses-and-fonts/THIRD-PARTY-NOTICES.md')
 copy('artifacts/pdf-vectors/pdfpig-source/LICENSE','licenses-and-fonts/PdfPig-Apache-2.0.txt')
 copy('artifacts/pdf-vectors/pdfpig-source/NOTICES.txt','licenses-and-fonts/PdfPig-NOTICES.txt')
 for p in independent.glob('visual-low3*.png'):copy(p,'independent/'+p.name)
 for package in ('skiasharp.nativeassets.macos','skiasharp'):
  folder=Path.home()/'.nuget/packages'/package/'3.119.1'
  for filename in ('LICENSE.txt','THIRD-PARTY-NOTICES.txt'):
   if (folder/filename).exists():copy(folder/filename,'licenses-and-fonts/'+package+'-'+filename)
 files=[{'file':str(p.relative_to(stage)),'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(stage.rglob('*')) if p.is_file()]
 with tarfile.open(stage/'evidence.tar','w') as archive:
  for item in files:archive.add(stage/item['file'],arcname=item['file'])
 archive=destination/'evidence.tar.zst'
 subprocess.run(['zstd','-q','-19','--long=27','-f',str(stage/'evidence.tar'),'-o',str(archive)],check=True)
 if archive.stat().st_size>95_000_000:raise ValueError('Evidence archive exceeds bounded public Git artifact size; preserve and narrow before committing')
 manifest={'baseline':'d878aeda57c1da79917e11d0a35bcc028bc4c1fd','runtime_head':args.runtime_head,'version':args.version,
           'archive':archive.name,'archive_bytes':archive.stat().st_size,'archive_sha256':sha(archive),'files':files}
 (destination/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
print(f'Persisted {len(files)} source-bound evidence files; archive {archive.stat().st_size} bytes. Preview status remains explicit.')
