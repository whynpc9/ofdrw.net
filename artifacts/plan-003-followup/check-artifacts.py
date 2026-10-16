from pathlib import Path
from zipfile import ZipFile
import xml.etree.ElementTree as E
import subprocess,json,hashlib,re
root=Path.cwd();out=root/'artifacts/plan-003-followup';visual=out/'visual'
ns={'w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
def normalize(s):return ''.join(s.split())
with ZipFile(root/'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx') as z:
 xml=E.fromstring(z.read('word/document.xml')); source=''.join(t.text or '' for t in xml.findall('.//w:t',ns))
checks=[]
for mode in ['native2','default2','directBuiltIn2','baseline-direct2']:
 pdf=visual/mode/(mode+'.pdf');text=subprocess.check_output(['pdftotext','-raw',str(pdf),'-'],text=True)
 assert normalize(source)==normalize(text),(mode,normalize(source),normalize(text))
 checks.append({'mode':mode,'complete_text_exactly_once':True,'normalized_characters':len(normalize(text))})
with ZipFile(visual/'italic-boundaries.docx') as z:
 xml=E.fromstring(z.read('word/document.xml')); boundary_source=''.join(t.text or '' for t in xml.findall('.//w:t',ns))
text=subprocess.check_output(['pdftotext','-raw',str(visual/'italic-boundaries/italic-boundaries.pdf'),'-'],text=True)
assert normalize(boundary_source)==normalize(text)
checks.append({'mode':'italic-boundaries','complete_text_exactly_once':True,'normalized_characters':len(normalize(text))})
selected_text=subprocess.check_output(['pdftotext','-raw',str(visual/'selected1/selected1.pdf'),'-'],text=True)
assert normalize(selected_text)==normalize(source[source.index('2. 第二页'):])
checks.append({'mode':'selected1','source_page2_text_exactly_once':True})
for mode in ['native2','default2','selected1']:
 ofd=visual/mode/(mode+'.ofd')
 with ZipFile(ofd) as z:
  names=sorted(n for n in z.namelist() if n.endswith('/Content.xml') and '/Pages/' in n)
  text=''.join(t.text or '' for n in names for t in E.fromstring(z.read(n)).iter() if t.tag.endswith('}TextCode'))
 expected=source if mode!='selected1' else source[source.index('2. 第二页'):]
 assert text==expected,(mode,text,expected)
 checks.append({'mode':mode,'native_source_text_exactly_once':True,'raw_characters':len(text),'raw_source_equality':True,'pages':len(names)})
# Same source Native/default page PNGs match. Unaffected direct page2 is byte-identical to baseline.
for page in [1,2]:
 assert (visual/'native2'/f'native2-{page}.png').read_bytes()==(visual/'default2'/f'default2-{page}.png').read_bytes()
assert (visual/'directBuiltIn2'/'directBuiltIn2-2.png').read_bytes()==(visual/'baseline-direct2'/'baseline-direct2-2.png').read_bytes()
checks.append({'native_default_pngs_equal':[1,2],'direct_unaffected_page2_equal_baseline':True})
fontreport=(out/'italic-boundaries-fonts.txt').read_text()
assert 'ce1d2f1ab89db45f9796100eee960f5702a40e84c225c2b48c3ec3e81d153f98,Italic' in fontreport
checks.append({'real_italic_arial_physical_sha256':'ce1d2f1ab89db45f9796100eee960f5702a40e84c225c2b48c3ec3e81d153f98','regular_fallback_sha256':'876af2cd4854644e7f3e7feb2f688997fdb3343c6df6693611209c9dfb47ccec'})
files=[]
for p in sorted(visual.rglob('*')):
 if p.is_file():files.append({'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
for p in [root/'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx',root/'src/Ofdrw.Net.Converter.Docx/bin/Release/netstandard2.1/Ofdrw.Net.Converter.Docx.dll']:
 files.append({'path':str(p),'bytes':p.stat().st_size,'sha256':hashlib.sha256(p.read_bytes()).hexdigest()})
(out/'visual-manifest.json').write_text(json.dumps({'source_commit':subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip(),'checks':checks,'files':files},indent=2))
(out/'text-integrity.json').write_text(json.dumps(checks,indent=2));print(json.dumps(checks,indent=2))
