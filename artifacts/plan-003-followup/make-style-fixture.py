from pathlib import Path
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
from xml.sax.saxutils import escape
root=Path.cwd(); out=root/'artifacts/plan-003-followup/visual'
out.mkdir(parents=True, exist_ok=True)
source=root/'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx'
def run(t,italic=False,font='Arial',color='000000'):
 return '<w:r><w:rPr><w:rFonts w:ascii="'+font+'" w:hAnsi="'+font+'"/><w:sz w:val="28"/>'+('<w:i/>' if italic else '')+'<w:color w:val="'+color+'"/></w:rPr><w:t xml:space="preserve">'+escape(t)+'</w:t></w:r>'
body='<w:p>'+run('BuiltIn italic boundaries: exact run / exact anchor')+'</w:p>'
body+='<w:p>'+run('Same words ')+run('Same words ',True,color='B32020')+run('Same words')+'</w:p>'
body+='<w:p>'+run('Genuine Arial italic: ')+run('Slanted source face, no second shear.',True)+'</w:p>'
body+='<w:p>'+run('Host regular fallback: ',font='Missing-Ofdrw-003')+run('Synthetic italic from original selected face.',True,'Missing-Ofdrw-003')+'</w:p>'
body+='<w:p>'+run('Wrap begins ')+run('Repeated mixed italic words continue across a narrow paragraph, ',True)+run('then regular resumes ')+run('and italic finishes.',True)+'</w:p>'
body+='<w:p>'+run('中文保持完整，English spacing; adjacent regular after italic.',font='Arial Unicode MS')+'</w:p>'
body+='<w:sectPr><w:pgSz w:w="7600" w:h="9000"/><w:pgMar w:top="700" w:right="700" w:bottom="700" w:left="700"/></w:sectPr>'
doc='<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>'+body+'</w:body></w:document>'
with ZipFile(source) as z, ZipFile(out/'italic-boundaries.docx','w',ZIP_DEFLATED) as target:
 for name in z.namelist():
  info=ZipInfo(name,(2026,1,1,0,0,0)); info.compress_type=ZIP_DEFLATED
  target.writestr(info,doc.encode() if name=='word/document.xml' else z.read(name))
print(out/'italic-boundaries.docx')
