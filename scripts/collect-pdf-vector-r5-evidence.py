#!/usr/bin/env python3
"""Retain R5 palette repair, original R4 color diagnosis and bounded package/Preview evidence."""
import argparse,hashlib,json,shutil,subprocess,tarfile,tempfile
from pathlib import Path
root=Path(__file__).resolve().parents[1]
p=argparse.ArgumentParser(description=__doc__);p.add_argument('runtime_head');p.add_argument('version');a=p.parse_args()
subprocess.run(['git','diff','--exit-code',a.runtime_head,'--','src','tests','e2e','.github/workflows/ci.yml','scripts/run-pdf-vector-package-e2e.sh','scripts/verify-pdf-vector-evidence.py'],cwd=root,check=True)
b=root/'artifacts/pdf-vectors';dest=root/'docs/evidence/pdf-vectors';low=b/'independent'
preview=json.loads((b/'preview-r5.json').read_text());validation=json.loads((low/'verification-low7.json').read_text())
assert preview['runtime_head']==a.runtime_head and preview['checked_pages']==8 and preview['gui_released']
assert validation['source']['commit']==a.runtime_head and validation['functional']['dotnet_total']==602 and validation['r4_observer']['result']=='passed'
sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
for row in preview['docs']:assert sha(Path(row['absolute_path']))==row['sha256']
excluded={'bin','obj','packages','feed','__pycache__','venv','.venv','dotnet-home','native','swift-module-cache','tmp','font-cache'}
with tempfile.TemporaryDirectory(prefix='ofd-vector-palette-evidence-') as tmp:
 stage=Path(tmp)
 def copy(source,target):
  target=stage/target;target.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(source,target)
 def tree(source,target,color=False):
  for path in sorted(source.rglob('*')):
   rel=path.relative_to(source)
   if not path.is_file() or any(part in excluded for part in rel.parts) or path.suffix=='.svg' or path.name=='quartz-render':continue
   # Reproducible controlled PDF copies repeat the full source font; their scripts/hashes,
   # actual R4 source/target in the prior archive, RGB data and actual rendered PNGs are retained.
   if color and path.suffix=='.pdf' and path.stat().st_size>128000:continue
   copy(path,Path(target)/rel)
 for path in sorted((b/'package-r5/output').iterdir()):
  if path.is_file() and (path.name.startswith('indexed-') or path.name.endswith('-report.json') or path.name=='verified-evidence.json'):copy(path,'candidate/'+path.name)
 for folder in ['indexed-color-r4-design','indexed-color-r5-actual']:tree(b/folder,folder,color=True)
 for name in ['package-manifest.json','consumer.assets.json']:copy(b/'package-r5'/name,'packages/primary-'+name)
 copy(root/f'artifacts/package-e2e/{a.version}/packages/package-manifest.json','packages/default11.json')
 for product in ['Ofdrw.Net.Converter.Pdf.Vector','Ofdrw.Net.Graphics.SkiaSharp']:copy(b/f'package-r5/feed/{product}.{a.version}.nupkg',f'packages/{product}.{a.version}.nupkg')
 for name in ['preview-r5.json','unchanged-baseline-r5.json','tests-indexed-r5-preflight.log','tests-indexed-r5-expanded.log','tests-indexed-r5-controls.log','full-r5.log','python-r5.log','python-r5-corrected.log','default-package-r5.log','optional-package-r5.log','pr18-readable-r5-initial.json','pr18-readable-r5-before-text.json']:
  copy(b/name,'acceptance/'+name)
 tree(b/'full-r5-trx','acceptance/trx')
 for name in ['verification-low7.json','source-hashes-low7.json']:copy(low/name,'independent/'+name)
 for folder in ['logs-low7','observer-low7','observer-output-low7']:tree(low/folder,'independent/'+folder)
 for name in ['package-manifest.json','consumer.assets.json']:copy(low/'optional13-low7'/name,'independent/'+name)
 for name in ['sample-report.json','image-hint-report.json','nonpainting-report.json','precision-resource-report.json','indexed-palette-report.json','verified-evidence.json']:copy(low/'optional13-low7/output'/name,'independent/'+name)
 copy(low/'default-low7/packages/package-manifest.json','independent/default11.json')
 for source,target in [(root/'LICENSE','MIT.txt'),(root/'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt','Noto-OFL.txt'),(root/'THIRD-PARTY-NOTICES.md','THIRD-PARTY-NOTICES.md')]:copy(source,'licenses/'+target)
 files=[{'file':str(path.relative_to(stage)),'bytes':path.stat().st_size,'sha256':sha(path)} for path in sorted(stage.rglob('*')) if path.is_file()]
 with tarfile.open(stage/'evidence.tar','w') as archive:
  for row in files:archive.add(stage/row['file'],arcname=row['file'])
 archive=dest/'review-r5.tar.zst';subprocess.run(['zstd','-q','-19','--long=27','-f',str(stage/'evidence.tar'),'-o',str(archive)],check=True)
 assert archive.stat().st_size<95000000
 manifest={'runtime_head':a.runtime_head,'version':a.version,'archive':archive.name,'archive_bytes':archive.stat().st_size,'archive_sha256':sha(archive),'files':files,'prior_archives':[{'file':name,'sha256':sha(dest/name)} for name in ['evidence.tar.zst','review-r3.tar.zst','review-r4.tar.zst']],'excluded_controlled_pdf_copies':'Large diagnostic variants reproducible from retained scripts and prior R4 immutable source; no product outputs replaced.'}
 (dest/'review-r5-manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(f'Archived{len(files)} files,{archive.stat().st_size}bytes; prior archives unchanged.')
