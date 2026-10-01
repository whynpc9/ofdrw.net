#!/usr/bin/env python3
"""Verify public tool samples, render every PDF page, and record reproducible hashes."""
import argparse, collections, hashlib, json, pathlib, subprocess, xml.etree.ElementTree as ET, zipfile
parser = argparse.ArgumentParser()
parser.add_argument('directory', type=pathlib.Path)
parser.add_argument('--render', action='store_true')
args = parser.parse_args()
directory = args.directory.resolve()
local = lambda node: node.tag.rsplit('}', 1)[-1]
def contents(name):
    with zipfile.ZipFile(directory / (name + '.ofd')) as archive:
        assert archive.testzip() is None
        return {path: archive.read(path) for path in archive.namelist()}
def pages(entries):
    document = ET.fromstring(entries['Doc_0/Document.xml'])
    return [node for parent in document.iter() if local(parent) == 'Pages' for node in parent if local(node) == 'Page']
def texts(entries, page):
    path = 'Doc_0/' + page.attrib['BaseLoc']
    return ''.join(node.text or '' for node in ET.fromstring(entries[path]).iter() if local(node) == 'TextCode')
expected = {'baseline-native':2,'baseline-default':2,'rich':2,'signed':2,'watermark':2,'watermark-merged':3,'split':2,'mix':1,'clean':2,'overlay':1,'cli-watermark':2,'cli-split':2,'cli-mix':1,'cli-clean':2,'cli-merged':3}
with zipfile.ZipFile(directory/'licensed-layout.docx') as archive:
    source_text = ''.join(node.text or '' for node in ET.fromstring(archive.read('word/document.xml')).iter() if local(node)=='t')
source = contents('signed')
source_pages = pages(source)
checks = {}
for name, count in expected.items():
    data = contents(name)
    selected = pages(data)
    assert len(selected) == count, (name, 'page count')
    if name in ('baseline-native','baseline-default'):
        assert ''.join(texts(data,page) for page in selected) == source_text, (name,'DOCX text changed')
    root = ET.fromstring(data['OFD.xml'])
    declarations = [child for body in root if local(body) == 'DocBody' for child in body if local(child) == 'Signatures']
    if name not in ('signed',): assert not declarations, (name, 'residual signatures')
    if name in ('clean','cli-clean'):
        for path, value in source.items():
            if path == 'OFD.xml' or '/Signs/' in path: continue
            assert data.get(path) == value, (name, 'body changed', path)
        assert not any('/Signs/' in path for path in data), (name, 'signature payload retained')
    if name in ('split','cli-split'):
        assert [texts(data, page) for page in selected] == [texts(source, source_pages[i]) for i in (1,0)], (name, 'selection text')
    if name in ('watermark','watermark-merged','cli-watermark','cli-merged'):
        assert sum(texts(data, page).count('DRAFT 草稿') for page in selected) == 1, (name, 'duplicate watermark')
        assert texts(source, source_pages[0]) in texts(data, selected[0]), (name, 'body text lost')
    if name in ('mix','cli-mix'):
        assert texts(source, source_pages[0]) in texts(data, selected[0]), (name, 'mixed body lost')
        assert texts(data, selected[0]).count('TOP LAYER 上层') == 1
    pdf = directory / (name + '.pdf')
    info = subprocess.check_output(['pdfinfo', str(pdf)], text=True)
    assert int(next(line.split(':')[1] for line in info.splitlines() if line.startswith('Pages:'))) == count
    subprocess.run(['pdftotext','-layout',str(pdf),str(directory/(name+'.pdf.txt'))],check=True)
    subprocess.run(['pdftotext','-raw',str(pdf),str(directory/(name+'.pdf-raw.txt'))],check=True)
    pdf_text = (directory/(name+'.pdf-raw.txt')).read_text()
    if name in ('watermark','watermark-merged','cli-watermark','cli-merged'): assert pdf_text.count('DRAFT 草稿') == 1
    if name in ('mix','cli-mix'): assert pdf_text.count('TOP LAYER 上层') == 1 and pdf_text.count('UNDER LAYER 下层') == 1
    if name in ('rich','signed','watermark','watermark-merged','split','mix','clean','cli-watermark','cli-split','cli-mix','cli-clean','cli-merged'):
        assert pdf_text.count('NOTE 注释') == 1 and pdf_text.count('TEMPLATE 模板') == 1
    if args.render:
        (directory/'pages').mkdir(exist_ok=True)
        subprocess.run(['pdftoppm','-scale-to','1200','-png',str(pdf),str(directory/'pages'/name)],check=True,stdout=subprocess.DEVNULL,stderr=subprocess.DEVNULL)
    checks[name] = {'pages':count,'signatures':len(declarations),'pdf_text_checked':True}
for name in ('watermark','watermark-merged','cli-watermark','cli-merged'):
    svg = ET.parse(directory/(name+'-1.svg')).getroot()
    assert sum(''.join(node.itertext()).count('DRAFT 草稿') for node in svg.iter() if local(node)=='text') == 1
    assert any(local(node)=='image' and any(value.startswith('data:image/png;base64,') for value in node.attrib.values()) for node in svg.iter())
    if args.render:
        subprocess.run(['rsvg-convert','--background-color','white','-w','849','-h','1200','-o',str(directory/'pages'/(name+'-svg.png')),str(directory/(name+'-1.svg'))],check=True)
files = {str(path.relative_to(directory)): {'bytes':path.stat().st_size,'sha256':hashlib.sha256(path.read_bytes()).hexdigest()} for path in sorted(directory.rglob('*')) if path.is_file() and path.name != 'manifest.json'}
manifest = {'checks':checks,'files':files,'render_tool':subprocess.check_output(['pdftoppm','-v'],stderr=subprocess.STDOUT,text=True).splitlines()[0]}
(directory/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2)+'\n')
print(f'Verified {len(checks)} API/CLI samples, {sum(item["pages"] for item in checks.values())} PDF pages, OFD/PDF watermark text count and embedded SVG images; hashes saved.')
