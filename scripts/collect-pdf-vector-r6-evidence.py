#!/usr/bin/env python3
"""Retain text producer repair, final-runtime scoped Preview and the failed Linux frozen-input gate."""
import argparse,hashlib,json,shutil,subprocess,tarfile,tempfile
from pathlib import Path
root=Path(__file__).resolve().parents[1]
p=argparse.ArgumentParser(description=__doc__);p.add_argument('runtime_head');p.add_argument('version');a=p.parse_args()
subprocess.run(['git','diff','--exit-code',a.runtime_head,'--','src','tests','e2e','.github/workflows/ci.yml','scripts/run-pdf-vector-package-e2e.sh','scripts/verify-pdf-vector-evidence.py'],cwd=root,check=True)
b=root/'artifacts/pdf-vectors';dest=root/'docs/evidence/pdf-vectors';low=b/'independent';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
preview=json.loads((b/'preview-r6.json').read_text());validation=json.loads((low/'verification-low8.json').read_text())
assert preview['runtime_head']==a.runtime_head and preview['checked_pages']==20 and preview['gui_released']
assert validation['source']['commit']==a.runtime_head and validation['functional']['total']==625 and validation['observer']['control_counts']['malformed_palette']==5
for d in preview['docs']:assert sha(Path(d['absolute_path']))==d['sha256']
excluded={'bin','obj','packages','feed','__pycache__','venv','.venv','dotnet-home','native','tmp','font-cache'}
with tempfile.TemporaryDirectory(prefix='ofd-vector-text-evidence-') as temporary:
 stage=Path(temporary)
 def copy(source,target):
  target=stage/target;target.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(source,target)
 def tree(source,target):
  for path in sorted(source.rglob('*')):
   rel=path.relative_to(source)
   if path.is_file() and path.suffix!='.svg' and not any(part in excluded for part in rel.parts):copy(path,Path(target)/rel)
 for path in sorted((b/'package-r6/output').iterdir()):
  if path.is_file() and (path.name.startswith(('text-precision','indexed-image-clipped')) or path.name.endswith('-report.json') or path.name=='verified-evidence.json'):copy(path,'candidate/'+path.name)
 tree(b/'text-precision-r5-design','text-precision-r5-design')
 tree(b/'text-r6-aux','auxiliary')
 for name in ['package-manifest.json','consumer.assets.json']:copy(b/'package-r6'/name,'packages/primary-'+name)
 copy(root/f'artifacts/package-e2e/{a.version}/packages/package-manifest.json','packages/default11.json')
 for product in ['Ofdrw.Net.Converter.Pdf.Vector','Ofdrw.Net.Graphics.SkiaSharp']:copy(b/f'package-r6/feed/{product}.{a.version}.nupkg',f'packages/{product}.{a.version}.nupkg')
 for name in ['preview-r6.json','unchanged-baseline-r6.json','full-r6.log','python-r6.log','default-package-r6.log','optional-package-r6.log','tests-text-r6-preflight.log','source-r6-preflight.log','source-r6-build.log','source-r6-direct.log','pr18-readable-r6-initial.json','pr18-readable-r6-postgui.json','pr18-package-ci-r6-failed.log']:copy(b/name,'acceptance/'+name)
 tree(b/'full-r6-trx','acceptance/main-trx')
 for name in ['verification-low8.json','source-hashes-low8.json']:copy(low/name,'independent/'+name)
 for folder in ['logs-low8','observer-low8','observer-original-low8','observer-output-low8']:tree(low/folder,'independent/'+folder)
 for name in ['package-manifest.json','consumer.assets.json']:copy(low/'optional13-low8'/name,'independent/'+name)
 for name in ['sample-report.json','image-hint-report.json','nonpainting-report.json','precision-resource-report.json','indexed-palette-report.json','text-precision-report.json','verified-evidence.json']:copy(low/'optional13-low8/output'/name,'independent/'+name)
 copy(low/'default-low8/packages/package-manifest.json','independent/default11.json')
 for source,target in [(root/'LICENSE','MIT.txt'),(root/'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt','Noto-OFL.txt'),(root/'THIRD-PARTY-NOTICES.md','THIRD-PARTY-NOTICES.md')]:copy(source,'licenses/'+target)
 files=[{'file':str(path.relative_to(stage)),'bytes':path.stat().st_size,'sha256':sha(path)} for path in sorted(stage.rglob('*')) if path.is_file()]
 with tarfile.open(stage/'evidence.tar','w') as tar:
  for row in files:tar.add(stage/row['file'],arcname=row['file'])
 archive=dest/'review-r6.tar.zst';subprocess.run(['zstd','-q','-19','--long=27','-f',str(stage/'evidence.tar'),'-o',str(archive)],check=True)
 assert archive.stat().st_size<95000000
 manifest={'runtime_head':a.runtime_head,'version':a.version,'archive':archive.name,'archive_bytes':archive.stat().st_size,'archive_sha256':sha(archive),'files':files,'prior_archives':[{'file':name,'sha256':sha(dest/name)} for name in ['evidence.tar.zst','review-r3.tar.zst','review-r4.tar.zst','review-r5.tar.zst']],'status':'bounded functional/Preview passed; Linux frozen-input gate failed and overall PR open'}
 (dest/'review-r6-manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
print(f'Archived{len(files)}files,{archive.stat().st_size}bytes; previous archives unchanged.')
