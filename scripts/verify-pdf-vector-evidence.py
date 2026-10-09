#!/usr/bin/env python3
"""Verify actual same-PDF structure/text and rendered pages; never infer Preview."""
from pathlib import Path
from xml.etree import ElementTree as ET
import collections, hashlib, json, re, shutil, subprocess, sys, zipfile

def imagemagick(tool):
 magick=shutil.which('magick')
 if magick: return [magick,tool]
 executable=shutil.which(tool)
 if not executable: raise RuntimeError(f'ImageMagick {tool} is required (magick or ImageMagick 6 standalone tool).')
 return [executable]

compare=imagemagick('compare'); identify=imagemagick('identify')

root=Path(sys.argv[1]).resolve()
report=json.loads((root/'sample-report.json').read_text())
local=lambda node:node.tag.rsplit('}',1)[-1]
structures={}
for mode in ('dual','vector','fallback'):
 with zipfile.ZipFile(root/(mode+'.ofd')) as z:
  pages=sorted(n for n in z.namelist() if re.fullmatch(r'Doc_0/Pages/Page_\d+/Content.xml',n))
  objects=[]; original=[]
  for page in pages:
   xml=ET.fromstring(z.read(page)); objects.extend(local(n) for n in xml.iter() if local(n).endswith('Object'))
   original.extend(n.text or '' for n in xml.iter() if local(n)=='TextCode')
  structures[mode]={'pages':len(pages),'objects':dict(collections.Counter(objects)),'original':''.join(original)}
  if mode=='vector':
   assert len(pages)==2 and 'ImageObject' not in objects and 'PathObject' in objects and 'TextObject' in objects
   assert 'A  B' in ''.join(original) and '中文' in ''.join(original)
  if mode=='dual': assert len(pages)==2 and objects.count('ImageObject')==2 and 'PathObject' not in objects
  if mode=='fallback': assert len(pages)==5 and objects.count('ImageObject')==4 and objects.count('PathObject')==0
metrics={}
for page,source,native in [('1',root/'same-source-1.png',root/'vector-1.png'),('2',root/'same-source-2.png',root/'vector-2.png'),('no-op',root/'no-op-source-1.png',root/'no-op-1.png')]:
 result=subprocess.run(compare+['-metric','RMSE',str(source),str(native),'null:'],capture_output=True,text=True)
 assert result.returncode in (0,1),result.stderr
 match=re.search(r'\(([0-9.eE+-]+)\)',result.stderr); assert match,result.stderr
 error=float(match[1]); assert error<0.015,('source/native rendered mismatch',page,error)
 metrics[str(page)]={'normalized_rmse':error,'maximum':0.015,'role':'automated regression only, Preview separate'}
nonpainting=json.loads((root/'nonpainting-report.json').read_text())
expected={'no-op':(True,0),'no-content':(False,1),'closed-point':(False,1),'singular':(False,1),'no-op-scan':(False,1)}
assert {row['name'] for row in nonpainting}==set(expected)
for row in nonpainting:
 page=row['result']['Pages'][0]; native,images=expected[row['name']]
 assert page['IsNative']==native and page['ImageObjects']==images,(row['name'],page)
 if not native: assert page['PathObjects']==0,(row['name'],page)
precision_resources=json.loads((root/'precision-resource-report.json').read_text())
expected_precision={'precision':[False,False,True],'encoding-stream':[True,False],'cid-map-stream':[True,False],'indexed-image':[False]}
assert {row['name'] for row in precision_resources}==set(expected_precision)
for row in precision_resources:
 pages=row['result']['Pages']; assert [page['IsNative'] for page in pages]==expected_precision[row['name']]
 for page in pages:
  if not page['IsNative']: assert page['ImageObjects']==1 and page['PathObjects']==0,(row['name'],page)
image_metrics={}
for name in ('absent','false','true'):
 paths=[root/f'image-{name}-source-1.png',root/f'image-{name}-1.png']
 sizes=[tuple(map(int,subprocess.check_output(identify+['-format','%w %h',str(p)],text=True).split())) for p in paths]
 # OFD's existing0.001mm physical-page serialization can add one ceil raster pixel.
 # Keep the original canvases; compare only the common viewport without resampling.
 assert all(abs(sizes[0][i]-sizes[1][i])<=1 for i in (0,1)),('unexpected page extent change',name,sizes)
 extent=f'{min(s[0] for s in sizes)}x{min(s[1] for s in sizes)}+0+0'
 result=subprocess.run(compare+['-metric','RMSE','(',str(paths[0]),'-crop',extent,'+repage',')','(',str(paths[1]),'-crop',extent,'+repage',')','null:'],capture_output=True,text=True)
 assert result.returncode in (0,1),result.stderr
 match=re.search(r'\(([0-9.eE+-]+)\)',result.stderr); assert match,result.stderr
 error=float(match[1]); assert error<0.001,('original image source/export mismatch',name,error)
 image_metrics[name]={'normalized_rmse':error,'maximum':0.001,'source_and_output_canvas':sizes,'common_viewport':extent,'same_renderer':'Poppler144DPI; no resampling; existing physical-page mm rounding recorded; across-renderer algorithms not compared'}
text=subprocess.check_output(['pdftotext','-raw',str(root/'vector.pdf'),'-'],text=True)
for phrase in ('No. 2026-021','192.00','Restored context - page 2 / 2'):
 assert text.count(phrase)==1,(phrase,text.count(phrase))
files=[{'file':p.name,'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(root.iterdir()) if p.is_file() and p.name!='verified-evidence.json']
(root/'verified-evidence.json').write_text(json.dumps({'status':'passed','structures':structures,'source_vs_native_png':metrics,'source_vs_preserved_image_png':image_metrics,'image_tools':{'compare':compare,'identify':identify,'version':subprocess.check_output(compare+['-version'],text=True)},'files':files,'preview':'not inferred'},ensure_ascii=False,indent=2)+'\n')
print('PDF vector structure, original text, no duplicate export text and same-source PNG gates passed.')
