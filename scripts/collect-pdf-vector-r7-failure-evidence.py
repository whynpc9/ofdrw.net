#!/usr/bin/env python3
"""Retain the immutable R7 fixture-copy failure and actual Linux source-width forensics."""
import hashlib,json,shutil,subprocess,tarfile,tempfile
from pathlib import Path
root=Path(__file__).resolve().parents[1];b=root/'artifacts/pdf-vectors';d=root/'docs/evidence/pdf-vectors';h='abd85f3b613e96fdd1aa2bc27c71b51c029bd243';sha=lambda p:hashlib.sha256(p.read_bytes()).hexdigest()
subprocess.run(['git','diff','--exit-code',h,'--','src','tests','e2e','scripts/run-pdf-vector-package-e2e.sh'],cwd=root,check=True)
low=b/'independent';r=json.loads((low/'verification-low9.json').read_text());assert r['exact_head']==h and r['status']=='failed_first_actual_gate'
with tempfile.TemporaryDirectory(prefix='ofd-frozen-input-failure-') as tmp:
 stage=Path(tmp)
 def copy(src,dst):
  dst=stage/dst;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dst)
 for name in ['source-r7-build.log','source-r7-direct.log','full-r7.log','python-r7.log','default-package-r7.log','optional-package-r7.log','pr18-package-ci-r6-failed.log']:copy(b/name,'main/'+name)
 for name in ['verification-low9.json','source-hashes-low9.json']:copy(low/name,'independent/'+name)
 for path in (low/'logs-low9').glob('*'):copy(path,'independent/logs/'+path.name)
 for path in (b/'full-r7-trx').glob('*'):copy(path,'main/trx/'+path.name)
 for name in ['extract.py','extract.log','extract-explicit-range.py','extract-explicit-range.log','download.json','ci-same-source.pdf','font-width-differences.json','full-download-attempt.json']:copy(b/'ci-r6-input-forensics'/name,'ci-forensics/'+name)
 for name in ['evaluation.json','external-evaluation.json','structural.log','diagnose.sh']:copy(b/'fixture-copy-r7-diagnosis'/name,'copy-diagnosis/'+name)
 for name in ['file-inventory.txt','post-restore-none.json','task-directory.txt']:copy(b/'fixture-copy-r7-diagnosis/structural'/name,'copy-diagnosis/'+name)
 for name in ['package-manifest.json']:copy(root/'artifacts/package-e2e/0.1.0-pdfvector.20261009.r7/packages'/name,'packages/main-default11.json')
 for product in ['Ofdrw.Net.Converter.Pdf.Vector','Ofdrw.Net.Graphics.SkiaSharp']:copy(b/f'package-r7/feed/{product}.0.1.0-pdfvector.20261009.r7.nupkg',f'packages/{product}.0.1.0-pdfvector.20261009.r7.nupkg')
 copy(root/'e2e/Ofdrw.Net.Pdf.Vector.E2E/testdata/README.md','input/README.md');copy(root/'e2e/Ofdrw.Net.Pdf.Vector.E2E/testdata/Noto-OFL.txt','licenses/Noto-OFL.txt');copy(root/'LICENSE','licenses/MIT.txt')
 files=[{'file':str(p.relative_to(stage)),'bytes':p.stat().st_size,'sha256':sha(p)} for p in sorted(stage.rglob('*')) if p.is_file()]
 with tarfile.open(stage/'evidence.tar','w') as t:
  for item in files:t.add(stage/item['file'],arcname=item['file'])
 archive=d/'review-r7-failed.tar.zst';subprocess.run(['zstd','-q','-19','--long=27','-f',str(stage/'evidence.tar'),'-o',str(archive)],check=True);assert archive.stat().st_size<95000000
 (d/'review-r7-failed-manifest.json').write_text(json.dumps({'runtime_head':h,'production_src_same_as':'deda6d2928c2db0e05640d2df96e50b79131ccfa','status':'failed fixture-copy gate; no new Preview accepted','archive':archive.name,'archive_bytes':archive.stat().st_size,'archive_sha256':sha(archive),'files':files,'golden_gzip_checked_in':'e2e/Ofdrw.Net.Pdf.Vector.E2E/testdata/text-precision-source.pdf.gz','golden_gzip_sha256':'024e4dbd0f735b90c9cf50f91d1e245cfab1caca9e2e73af243afa8fb4def3fa','prior_archives':[{'file':n,'sha256':sha(d/n)} for n in ['evidence.tar.zst','review-r3.tar.zst','review-r4.tar.zst','review-r5.tar.zst','review-r6.tar.zst']]},indent=2)+'\n')
print('Retained R7 failed gate and actual CI source artifact',len(files),archive.stat().st_size)
