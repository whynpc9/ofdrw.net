from pathlib import Path
import json
from pypdf import PdfReader
root=Path.cwd();out=root/'artifacts/plan-003-followup'
source=out/'visual/italic-boundaries/italic-boundaries.pdf'
reader=PdfReader(source);records=[]
def visit(text,cm,tm,font,size):
 if text.strip():records.append(dict(text=text,basefont=str(font.get('/BaseFont')) if font else None,cm=list(cm),tm=list(tm),size=size))
reader.pages[0].extract_text(visitor_text=visit)
real='ce1d2f1ab89db45f9796100eee960f5702a40e84c225c2b48c3ec3e81d153f98'
fallback='876af2cd4854644e7f3e7feb2f688997fdb3343c6df6693611209c9dfb47ccec'
regular='525979822591a3447cfc49d943d6f7683508e25543407871c0ed8fed05fd2bd9'
actual=[r for r in records if real in (r['basefont'] or '')]
simulated=[r for r in records if fallback in (r['basefont'] or '')]
upright=[r for r in records if regular in (r['basefont'] or '')]
assert actual and simulated and upright
assert all(abs(r['cm'][2])<1e-6 and abs(r['tm'][2])<1e-6 for r in actual)
assert all(abs(r['cm'][2]-.3420201)<1e-6 for r in simulated)
assert all(abs(r['cm'][2])<1e-6 and abs(r['tm'][2])<1e-6 for r in upright)
report={'source':str(source),'independent_parser':'pypdf '+__import__('pypdf').__version__,'physical_real_italic_no_added_skew':True,'physical_regular_fallback_has_skew':True,'adjacent_regular_no_skew':True,'records':records}
(out/'pdf-style-proof.json').write_text(json.dumps(report,indent=2,ensure_ascii=False));print({k:v for k,v in report.items() if k!='records'})
