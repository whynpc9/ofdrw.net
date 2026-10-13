#!/usr/bin/env python3
"""Preserve R13's bounded producer contract and final R12+R13 page inspection."""
from pathlib import Path
import hashlib
import json
import shutil
import subprocess
import tarfile
import tempfile

root = Path(__file__).resolve().parents[1]
b = root / 'artifacts/pdf-vectors'
d = root / 'docs/evidence/pdf-vectors'
runtime = '7afd12042e9099093d9974472d1431686304a162'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
subprocess.run(['git', 'diff', '--exit-code', runtime, '--', 'src', 'tests', 'e2e'], cwd=root, check=True)
freeze = json.loads((b / 'source-r13-delivery-freeze.json').read_text())
for row in freeze['files']:
    p = root / row['path']
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
for row in freeze['packages']:
    p = root / row['file']
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
low = b / 'independent/low15'
verification = json.loads((low / 'verification.json').read_text())
assert verification['runtimeHead'] == runtime and verification['status'] == 'functional_pass_preview_pending'
for row in verification['artifactInventory']:
    p = Path(row['path'])
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
preview = json.loads((b / 'preview-r13.json').read_text())
assert preview['runtime_head'] == runtime and preview['total_pages'] == 20
assert preview['status'] == 'passed_bounded_visual_inspection'
for row in preview['docs']:
    p = Path(row['absolute_path'])
    assert row['checked_pages'] == list(range(1, row['pages'] + 1))
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
excluded = {'bin', 'obj', 'packages', 'package-cache', 'cache', 'feed', '__pycache__', 'dotnet-home',
            'tmp', 'dependency-feed', 'native', 'font-cache', '.git', 'venv'}
with tempfile.TemporaryDirectory(prefix='ofd-r13-evidence-') as tmp:
    stage = Path(tmp)
    def copy(p, name):
        q = stage / name
        q.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, q)
    def tree(source, target, skip_source=False):
        for p in sorted(source.rglob('*')):
            rel = p.relative_to(source)
            if p.is_file() and p.suffix != '.svg' and not any(x in excluded or (skip_source and x == 'source') for x in rel.parts):
                copy(p, Path(target) / rel)
    for pattern in ['*-r13*.log', '*-r13*.json', 'r13*.log', 'r13*.json', 'preview-r13*.json']:
        for p in b.glob(pattern):
            if p.is_file(): copy(p, 'main/' + p.name)
    for name in ['r12-main-settlement.json', 'low14-parent-settlement.json', 'low15-parent-settlement.json',
                 'r12-functional-archive-readback.json', 'r12-public-parts-reconstruction.log']:
        copy(b / name, 'main/' + name)
    for name in ['r13-targeted-trx', 'r13-targeted-corrected-trx', 'full-r13-trx']:
        tree(b / name, 'main/' + name)
    tree(b / 'design-r13', 'design-r13')
    tree(b / 'preview-r13-ui', 'main/preview-r13-ui')
    tree(b / 'replay-r13', 'main/replay-r13')
    tree(b / 'replay-r12-on-r13', 'main/replay-r12-on-r13')
    tree(low, 'independent/low15', skip_source=True)
    copy(low / 'default/feed/package-manifest.json', 'independent/low15/default11-manifest.json')
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(b / 'package-r13' / name, 'main/optional13/' + name)
    copy(root / 'artifacts/package-e2e/0.1.0-pdfvector.20261013.r13/packages/package-manifest.json', 'main/default11.json')
    new_pdfs = {name + '.pdf' for name in ['font-usage-source', 'font-usage', 'xml-ufffe-source', 'xml-ufffe', 'xml-uffff-source', 'xml-uffff']}
    for p in sorted((b / 'package-r13/output').iterdir()):
        if p.suffix in {'.ofd', '.json', '.png'} or p.name in new_pdfs: copy(p, 'candidate/' + p.name)
    for label, feed, ver in [('main', b / 'package-r13/feed', '0.1.0-pdfvector.20261013.r13'),
                             ('independent', low / 'optional/feed', '0.1.0-pdfvector.20261013.low15')]:
        for product in ['Ofdrw.Net.Core', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Converter.Pdf.Vector', 'Ofdrw.Net.Graphics.SkiaSharp']:
            copy(feed / f'{product}.{ver}.nupkg', f'packages/{label}/{product}.{ver}.nupkg')
    for p, name in [(root / 'LICENSE', 'MIT.txt'), (root / 'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt', 'Noto-OFL.txt'),
                    (root / 'THIRD-PARTY-NOTICES.md', 'THIRD-PARTY-NOTICES.md')]: copy(p, 'licenses/' + name)
    files = [{'file': str(p.relative_to(stage)), 'bytes': p.stat().st_size, 'sha256': sha(p)} for p in sorted(stage.rglob('*')) if p.is_file()]
    with tarfile.open(stage / 'evidence.tar', 'w') as archive:
        for row in files: archive.add(stage / row['file'], arcname=row['file'])
    local = b / 'r13-final.tar.zst'
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'evidence.tar'), '-o', str(local)], check=True)
    parts = []
    with local.open('rb') as stream:
        index = 1
        while chunk := stream.read(90_000_000):
            p = d / f'review-r13.tar.zst.part{index:02}'
            p.write_bytes(chunk)
            parts.append({'order': index, 'file': p.name, 'bytes': len(chunk), 'sha256': sha(p)})
            index += 1
    manifest = {'runtime_head': runtime, 'archive_bytes': local.stat().st_size, 'archive_sha256': sha(local),
                'parts': parts, 'files': files,
                'reconstruction': 'Concatenate parts in order into r13-final.tar.zst; verify archive SHA256, then zstd -dc | tar -tf -. Each part is below 90 MiB.',
                'bounded_scope': 'Final actual OFD/PNG and six inspected PDFs, actual replay OFDs and original independent logs retained. Derived SVG font-base64 copies and build/cache trees omitted; R12 original inputs and historical failures remain in the separate immutable R12 functional archive.',
                'prior_archives': [{'file': p.name, 'sha256': sha(p)} for p in sorted(d.glob('*.tar.zst'))]}
    (d / 'review-r13-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    acceptance = {'runtime_head': runtime, 'main': {'dotnet': 715, 'trx_files': 9, 'python': 5, 'default_packages': 11,
                  'optional_packages': 13, 'saved_R12_replay': 62, 'BOM': 2, 'saved_R13_replay': 80},
                  'independent_low15': verification, 'preview': preview,
                  'R12': 'Functional archive at 967d, visual pending/superseded. The final 7afd font-usage pair is inspected here; do not relabel 967d as visually accepted.',
                  'XML_scope': 'Producer TEXT_XML_CHARACTER classification and explicit RasterizePage+None full-page success; default Invisible cannot represent FFFE/FFFF and actual ArgumentException/sentinel failure is retained. Raster page has no searchable original text.',
                  'historical_visual': 'R10 28 pages runtime28fc; R11 8 pages runtime612. Unchanged historical pages not reopened or relabelled.',
                  'final_head_reviews_and_ci': 'Tracked separately on PR18; not inherited', 'release_accepted': False, 'merged': False, 'published': False}
    (d / 'review-r13-acceptance.json').write_text(json.dumps(acceptance, indent=2) + '\n')
    print(json.dumps({'archive_bytes': manifest['archive_bytes'], 'archive_sha256': manifest['archive_sha256'], 'files': len(files), 'parts': parts}))
