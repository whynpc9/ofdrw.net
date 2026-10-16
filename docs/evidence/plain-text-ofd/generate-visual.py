from pathlib import Path
import subprocess,os,json,hashlib,zipfile,xml.etree.ElementTree as ET,re
root=Path.cwd();out=root/'artifacts/plan-028';visual=out/'visual';env=os.environ.copy();env.update(DOTNET_CLI_HOME=str(out/'dotnet-home'),DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1',DOTNET_CLI_TELEMETRY_OPTOUT='1')
cli=['dotnet',str(root/'src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll')];commands=[]
def run(args):
 result=subprocess.run(args,env=env,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True)
 commands.append({'command':args,'exit':result.returncode,'output':result.stdout})
 if result.returncode: raise RuntimeError(result.stdout)
 return result.stdout
samples=[('notice',['--page-width','150','--page-height','130','--margin-left','10','--margin-right','10','--margin-top','10','--margin-bottom','10','--font-size-mm','4']),('control',['--page-width','40','--page-height','30','--margin-left','5','--margin-right','5','--margin-top','5','--margin-bottom','5','--font-size-mm','10']),('empty',[])]
summary=[]
expect=json.loads((visual/'source-expectations.json').read_text())
for name,options in samples:
 src=visual/(name+'.txt');ofd=visual/(name+'.ofd');pdf=visual/(name+'.pdf')
 run(cli+['text-to-ofd',str(src),str(ofd),*options]);run(cli+['ofd-to-pdf',str(ofd),str(pdf)])
 with zipfile.ZipFile(ofd) as z:
  files=sorted([f for f in z.namelist() if re.search(r'/Pages/[^/]+/Content.xml$',f)],key=lambda f:[int(n) for n in re.findall(r'\d+',f)])
  # Packaging currently uses Page_N/Content.xml. Refuse a missing readback instead of vacuous success.
  if not files: files=sorted([f for f in z.namelist() if '/Pages/' in f and f.endswith('.xml') and 'Content' in f],key=lambda f:[int(n) for n in re.findall(r'\d+',f)])
  assert files,z.namelist()
  pages=[ET.fromstring(z.read(f)) for f in files];native=''.join(''.join(t.itertext()) for page in pages for t in page.iter() if t.tag.endswith('}TextCode'))
  assert native==expect[name]['expected_native_text'],(name,repr(native))
  assert all(not any(t.tag.endswith('}ImageObject') for t in page.iter()) for page in pages)
  if 'pages' in expect[name]: assert len(pages)==expect[name]['pages']
  if name=='control': assert not any(t.tag.endswith('}TextObject') for t in pages[1].iter())
 textfile=visual/(name+'-pdf-text.txt');run(['pdftotext','-raw',str(pdf),str(textfile)])
 plain=textfile.read_text();compact=lambda s:''.join(c for c in s if not c.isspace())
 assert compact(plain)==compact(expect[name]['expected_native_text']),(name,'PDF lost or repeated text',repr(plain))
 run(['pdftoppm','-png','-r','144',str(pdf),str(visual/(name+'-page'))])
 summary.append({'sample':name,'pages':len(pages),'source':str(src),'ofd':str(ofd),'pdf':str(pdf),'text_ordinal':'passed','pdf_text_no_duplicates':'passed','fonts':run(['pdffonts',str(pdf)])})
for name in ['generated-docx-native','generated-docx-default']:
 pdf=out/'package-output'/(name+'.pdf')
 assert pdf.is_file(),pdf
 summary.append({'sample':name,'pages':2,'pdf':str(pdf),'png_source':'package-e2e outputs','fonts':run(['pdffonts',str(pdf)])})
(out/'visual-commands.json').write_text(json.dumps(commands,ensure_ascii=False,indent=2));(out/'visual-summary.json').write_text(json.dumps(summary,ensure_ascii=False,indent=2));print(json.dumps(summary,ensure_ascii=False,indent=2))
