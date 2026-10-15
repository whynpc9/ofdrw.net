#!/usr/bin/env python3
"""Archive immutable R15 package/input/readback and eight-page Preview evidence."""
from pathlib import Path
import hashlib,json,shutil,subprocess,tarfile,tempfile
root=Path(__file__).resolve().parents[1]; b=root/'artifacts/pdf-vectors'; d=root/'docs/evidence/pdf-vectors'
runtime='f5914e6d41d63da541da5d8e031e7d97090296cc';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
subprocess.run(['git','diff','--exit-code',runtime,'--','src','tests','e2e'],cwd=root,check=True)
freeze=json.loads((b/'source-r15-delivery-freeze.json').read_text())
for row in freeze['files']:
 p=root/row['path'];assert p.stat().st_size==row['bytes'] and sha(p)==row['sha256']
for row in freeze['packages']:
 p=root/row['file'];assert p.stat().st_size==row['bytes'] and sha(p)==row['sha256']
low=b/'independent/low17';verification=json.loads((low/'verification.json').read_text())
assert verification['runtimeHead']==runtime and verification['status']=='functional_pass_preview_pending'
for row in verification['artifactInventory']:
 p=Path(row['path']);assert p.stat().st_size==row['bytes'] and sha(p)==row['sha256']
preview=json.loads((b/'preview-r15.json').read_text())
assert preview['runtime_head']==runtime and preview['status']=='passed_bounded_visual_inspection' and preview['total_pages']==8
for row in preview['docs']:
 p=Path(row['absolute_path']);assert row['checked_pages']==list(range(1,row['pages']+1))
 assert p.stat().st_size==row['bytes'] and sha(p)==row['sha256']
excluded={'bin','obj','packages','package-cache','cache','feed','__pycache__','dotnet-home','tmp','dependency-feed','native','font-cache','.git','venv','source','nuget-cache','cli-home'}
with tempfile.TemporaryDirectory(prefix='ofd-r15-evidence-') as tmp:
 stage=Path(tmp)
 def copy(p,name):
  q=stage/name;q.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(p,q)
 def tree(src,dest):
  for p in sorted(src.rglob('*')):
   rel=p.relative_to(src)
   if not p.is_file() or any(x in excluded for x in rel.parts):continue
   # Keep actual OFDs and evidence. Old unchanged page PDFs/PNGs remain in R13.
   if p.suffix in {'.svg','.png','.pdf'} and not p.name.startswith(('writer-accuracy','writer-tiny')):continue
   copy(p,Path(dest)/rel)
 for pattern in ['*-r15*.log','*-r15*.json','r15*.log','r15*.json','preview-r15*.json']:
  for p in b.glob(pattern):
   if p.is_file():copy(p,'main/'+p.name)
 for name in ['r15-targeted-trx','r15-baseline-red-trx','full-r15-first-trx','full-r15-trx','preview-r15-ui','replay-r15','replay-r14-on-r15','replay-r13-on-r15','replay-r12-on-r15']:
  tree(b/name,'main/'+name)
 # Original saved corpus and raw failed probe history, including its input PDFs.
 for p in sorted((b/'design-r15').rglob('*')):
  rel=p.relative_to(b/'design-r15')
  if p.is_file() and not any(x in excluded for x in rel.parts):copy(p,Path('design-r15')/rel)
 for name in ['verify-r15-replays.py','run-r15-replays.sh','analyze-r15-colors.py']:
  copy(b/name,'main/'+name)
 tree(low,'independent/low17')
 copy(b/'independent/r15-low-review.json','design-r15-independent-review.json')
 copy(low/'default/feed/package-manifest.json','independent/low17/default11-manifest.json')
 for name in ['package-manifest.json','consumer.assets.json']:
  copy(b/'package-r15'/name,'main/optional13/'+name)
 copy(root/'artifacts/package-e2e/0.1.0-pdfvector.20261013.r15/packages/package-manifest.json','main/default11.json')
 for p in sorted((b/'package-r15/output').iterdir()):
  if p.suffix in {'.ofd','.json'} or (p.name.startswith(('writer-accuracy','writer-tiny')) and p.suffix in {'.pdf','.png'}):copy(p,'candidate/'+p.name)
 for label,feed in [('main',b/'package-r15/feed'),('independent',low/'optional/feed')]:
  for product in ['Ofdrw.Net.Core','Ofdrw.Net.Layout','Ofdrw.Net.Converter.Pdf.Vector','Ofdrw.Net.Graphics.SkiaSharp']:
   matches=list(feed.glob(product+'.*.nupkg'));assert len(matches)==1
   copy(matches[0],f'packages/{label}/'+matches[0].name)
 for p,name in [(root/'LICENSE','MIT.txt'),(root/'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt','Noto-OFL.txt'),(root/'THIRD-PARTY-NOTICES.md','THIRD-PARTY-NOTICES.md')]:copy(p,'licenses/'+name)
 files=[{'file':str(p.relative_to(stage)),'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(stage.rglob('*')) if p.is_file()]
 with tarfile.open(stage/'evidence.tar','w') as archive:
  for row in files:archive.add(stage/row['file'],arcname=row['file'])
 local=b/'r15-final.tar.zst';subprocess.run(['zstd','-q','-19','--long=27','-f',str(stage/'evidence.tar'),'-o',str(local)],check=True)
 parts=[]
 with local.open('rb') as stream:
  index=1
  while chunk:=stream.read(90_000_000):
   p=d/f'review-r15.tar.zst.part{index:02}';p.write_bytes(chunk)
   parts.append({'order':index,'file':p.name,'bytes':len(chunk),'sha256':sha(p)});index+=1
 manifest={'runtime_head':runtime,'archive_bytes':local.stat().st_size,'archive_sha256':sha(local),'parts':parts,'files':files,
  'reconstruction':'Concatenate numbered parts; verify whole SHA256; zstd decompress and verify every file using verify-r15.py.',
  'bounded_scope':'Actual main/Low OFDs, new four-page source/export PDFs and tiny failure-input PDFs and PNG aids, saved 312-conversion corpus / 112 unique input PDFs and failed probe history, regression logs and current package/source identities. Derived SVG font copies and caches omitted. Unchanged older input/PDF/PNG evidence remains in immutable R12/R13/R14 archives.',
  'prior_parts':[{'file':p.name,'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(d.glob('review-r1[234]*.part*'))]}
 (d/'review-r15-manifest.json').write_text(json.dumps(manifest,indent=2)+'\n')
 acceptance={'runtime_head':runtime,'main':json.loads((b/'r15-main-settlement.json').read_text()),'independent_low17':verification,'preview':preview,
 'scope':'Producer total written scalar error certificate using actual double mm transform, old float gates retained; typed non-rasterizable page classification only after successful processing. Core/04/19/Writer unchanged; earlier tiny branches and default XML Invisible limitations retained.',
 'historical_visual':'R14 4 pages remain35fe; R13 20 pages remain7afd; R11 8 and R10 28 keep original identities. Only R15 new source/export eight pages checked here.',
 'final_head_reviews_and_ci':'Tracked separately on PR18, not inherited','release_accepted':False,'merged':False,'published':False}
 (d/'review-r15-acceptance.json').write_text(json.dumps(acceptance,indent=2)+'\n')
 print(json.dumps({'archive_bytes':manifest['archive_bytes'],'archive_sha256':manifest['archive_sha256'],'files':len(files),'parts':parts}))
