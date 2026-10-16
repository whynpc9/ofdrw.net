from pathlib import Path
import os, subprocess, json, hashlib, re, shutil, zipfile, xml.etree.ElementTree as ET

root = Path(__file__).resolve().parents[2]
out = root / 'artifacts/pr23-review-fixes'
visual = out / 'visual'
visual.mkdir(exist_ok=True)
env = os.environ.copy(); env.update(json.loads((out / 'environment.json').read_text())['environment'])
source = subprocess.check_output(['git', 'rev-parse', 'HEAD'], cwd=root, text=True).strip()
cli_path = root / 'src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll'
cli = ['dotnet', str(cli_path)]
commands = []

def run(args, extra=None):
    result = subprocess.run(args, cwd=root, env=dict(env, **(extra or {})), stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True)
    commands.append({'command': args, 'exit_code': result.returncode, 'output': result.stdout})
    (out / 'visual-commands.json').write_text(json.dumps(commands, indent=2) + '\n')
    if result.returncode:
        raise RuntimeError(result.stdout)
    return result.stdout

docx = root / 'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx'
run(cli + ['docx-to-pdf', str(docx), str(visual / 'direct-built-in.pdf'), '--docx-engine', 'built-in'])
run(cli + ['docx-to-pdf', str(docx), str(visual / 'external-libreoffice.pdf'), '--docx-engine', 'libreoffice'])
run(cli + ['docx-to-ofd', str(docx), str(visual / 'selected.ofd'), '--docx-ofd-mode', 'native', '--pages', '2'])
run(cli + ['ofd-to-pdf', str(visual / 'selected.ofd'), str(visual / 'selected.pdf')])
shutil.copyfile(root / 'docs/evidence/plain-text-ofd/visual/notice.txt', visual / 'notice.txt')
run(cli + ['text-to-ofd', str(visual / 'notice.txt'), str(visual / 'notice.ofd'), '--page-width', '150', '--page-height', '130',
           '--margin-left', '10', '--margin-right', '10', '--margin-top', '10', '--margin-bottom', '10', '--font-size-mm', '4'])
run(cli + ['ofd-to-pdf', str(visual / 'notice.ofd'), str(visual / 'notice.pdf')])
fontdir = visual / 'real-italic-fonts'; fontdir.mkdir(exist_ok=True)
for name in ['Arial.ttf', 'Arial Italic.ttf']:
    target = fontdir / name
    if not target.exists():
        target.symlink_to(Path('/System/Library/Fonts/Supplemental') / name)
run(cli + ['docx-to-pdf', str(root / 'artifacts/plan-003-followup/visual/italic-boundaries.docx'),
           str(visual / 'italic-boundaries.pdf'), '--docx-engine', 'built-in', '--font-directory', str(fontdir)])
flags = ['--disable-build-servers', '-m:1', '/nodeReuse:false', '/p:UseSharedCompilation=false']
run(['dotnet', 'test', 'tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj', '-c', 'Release',
     '--no-build', '--no-restore', '--filter', 'FullyQualifiedName~OneRead_IdIsolation_FontScope_DuplicateOrder_AndEscaping',
     '--logger', 'trx', '--results-directory', str(out / 'fixture-test'), *flags], {'OFDRW_TEST_ARTIFACTS': str(visual / 'html-fixtures')})
short_text = '纯文本 HTML 联合验证 Plain text iii WWW\n第二行保留 <literal> & text。\n'
(visual / 'notice-short.txt').write_text(short_text)
run(cli + ['text-to-ofd', str(visual / 'notice-short.txt'), str(visual / 'notice-short.ofd')])
for name, ofd in [('native', out / 'package-output/generated-docx-native.ofd'), ('notice-short', visual / 'notice-short.ofd')]:
    run(cli + ['ofd-to-html', str(ofd), str(visual / (name + '.html'))])

pdfs = [out / 'package-output' / (name + '.pdf') for name in ['generated-docx-native', 'generated-docx-default', 'cli-native']]
pdfs += [visual / (name + '.pdf') for name in ['direct-built-in', 'selected', 'notice', 'italic-boundaries', 'external-libreoffice']]
entries = []
for pdf in pdfs:
    count = int(re.search(r'^Pages:\s+(\d+)', run(['pdfinfo', str(pdf)]), re.M).group(1))
    prefix = visual / (pdf.stem + '-page')
    run(['pdftoppm', '-png', '-r', '144', str(pdf), str(prefix)])
    text = run(['pdftotext', '-layout', str(pdf), '-'])
    (visual / (pdf.stem + '-text.txt')).write_text(text)
    images = [{'path': str(p), 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in sorted(visual.glob(pdf.stem + '-page-*.png'))]
    assert len(images) == count
    entries.append({'pdf': str(pdf), 'pages': count, 'bytes': pdf.stat().st_size,
                    'sha256': hashlib.sha256(pdf.read_bytes()).hexdigest(), 'images': images})

with zipfile.ZipFile(docx) as archive:
    document = ET.fromstring(archive.read('word/document.xml'))
    original = ''.join(e.text or '' for e in document.iter() if e.tag.endswith('}t'))
compact = lambda value: ''.join(value.split())
for name in ['generated-docx-native', 'generated-docx-default', 'cli-native', 'direct-built-in']:
    assert compact((visual / (name + '-text.txt')).read_text()) == compact(original), name

htmls = [visual / 'native.html', visual / 'notice-short.html', visual / 'html-fixtures/mixed.html', out / 'html-package-output/package-api.html']
manifest = {'source_head': source, 'chain': 'DOCX/native or default OFD/PDF/PNG/Preview; direct BuiltIn separately; actual external LibreOffice separately; plain text/native OFD/PDF; HTML current API/CLI',
            'entries': entries, 'pdf_count': len(entries), 'page_count': sum(e['pages'] for e in entries),
            'html': [{'path': str(p), 'sha256': hashlib.sha256(p.read_bytes()).hexdigest()} for p in htmls],
            'cli_sha256': hashlib.sha256(cli_path.read_bytes()).hexdigest(),
            'docx_dll_sha256': hashlib.sha256((cli_path.parent / 'Ofdrw.Net.Converter.Docx.dll').read_bytes()).hexdigest(),
            'core_dll_sha256': hashlib.sha256((cli_path.parent / 'Ofdrw.Net.Core.dll').read_bytes()).hexdigest(),
            'native_source_sha256': hashlib.sha256((root / 'native/Ofdrw.ProcessSupervisor/supervisor.c').read_bytes()).hexdigest(),
            'native_manifest_sha256': hashlib.sha256((root / 'native/Ofdrw.ProcessSupervisor/assets/manifest.json').read_bytes()).hexdigest(),
            'automatic_render': 'PASS', 'png_visual_review': 'PENDING', 'preview': 'PENDING', 'browser': 'PENDING'}
(out / 'visual-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
print(json.dumps({'source': source, 'PDFs': len(entries), 'pages': manifest['page_count'], 'HTML_files': len(htmls)}))
