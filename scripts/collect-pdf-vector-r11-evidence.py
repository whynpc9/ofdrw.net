#!/usr/bin/env python3
"""Archive R11 fixes, immutable R10 reproductions and bounded new-page inspection."""
import hashlib
import json
import shutil
import subprocess
import tarfile
import tempfile
from pathlib import Path

root = Path(__file__).resolve().parents[1]
b = root / 'artifacts/pdf-vectors'
d = root / 'docs/evidence/pdf-vectors'
runtime = '61204987565bcb38821f877e0010d2c79f8c7730'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
subprocess.run(['git', 'diff', '--exit-code', runtime, '--', 'src', 'tests', 'e2e'], cwd=root, check=True)
freeze = json.loads((b / 'source-r11-delivery-freeze.json').read_text())
for row in freeze['packages']:
    p = root / row['file']
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
low = b / 'independent/low13'
verification = json.loads((low / 'verification.json').read_text())
assert verification['head'] == runtime
assert verification['status'] == 'functional_pass_visual_pending'
for row in verification['artifactInventory']:
    p = Path(row['path'])
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
preview = json.loads((b / 'preview-r11.json').read_text())
assert preview['runtime_head'] == runtime and preview['total_pages'] == 8
assert preview['status'] == 'passed_bounded_visual_inspection'
for row in preview['docs']:
    p = Path(row['absolute_path'])
    assert row['checked_pages'] == list(range(1, row['pages'] + 1))
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256']
excluded = {'bin', 'obj', 'packages', 'package-cache', 'cache', 'feed', '__pycache__',
            'dotnet-home', 'tmp', 'dependency-feed', 'native', 'font-cache', '.git'}
with tempfile.TemporaryDirectory(prefix='ofd-r11-evidence-') as tmp:
    stage = Path(tmp)

    def copy(p, target):
        q = stage / target
        q.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, q)

    def tree(source, target, skip_source=False):
        for p in sorted(source.rglob('*')):
            rel = p.relative_to(source)
            if p.is_file() and p.suffix != '.svg' and not any(x in excluded or (skip_source and x == 'source') for x in rel.parts):
                copy(p, Path(target) / rel)

    for pattern in ['*-r11*.log', '*-r11*.json', 'preview-r11*.json']:
        for p in b.glob(pattern):
            if p.is_file():
                copy(p, 'main/' + p.name)
    copy(b / 'low13-parent-settlement.json', 'main/low13-parent-settlement.json')
    copy(low / 'default/feed/package-manifest.json', 'independent/low13/default11-manifest.json')
    for name in ['r11-failing-trx', 'r11-failing-corrected-trx', 'r11-targeted-trx',
                 'full-r11-first-test-assumption-trx', 'full-r11-trx']:
        tree(b / name, 'main/' + name)
    tree(b / 'design-r11', 'design-r11')
    tree(low, 'independent/low13', skip_source=True)
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(b / 'package-r11' / name, 'main/optional13/' + name)
    copy(root / 'artifacts/package-e2e/0.1.0-pdfvector.20261013.r11/packages/package-manifest.json', 'main/default11.json')
    for p in sorted((b / 'package-r11/output').iterdir()):
        if p.suffix in {'.ofd', '.json', '.png'} or p.name in {'writer-precision-source.pdf', 'writer-precision.pdf'}:
            copy(p, 'candidate/' + p.name)
    for label, feed, ver in [('main', b / 'package-r11/feed', '0.1.0-pdfvector.20261013.r11'),
                             ('independent', low / 'optional/feed', '0.1.0-pdfvector.20261013.low13')]:
        for product in ['Ofdrw.Net.Core', 'Ofdrw.Net.Layout', 'Ofdrw.Net.Converter.Pdf.Vector', 'Ofdrw.Net.Graphics.SkiaSharp']:
            copy(feed / f'{product}.{ver}.nupkg', f'packages/{label}/{product}.{ver}.nupkg')
    for p, name in [(root / 'LICENSE', 'MIT.txt'),
                    (root / 'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt', 'Noto-OFL.txt'),
                    (root / 'THIRD-PARTY-NOTICES.md', 'THIRD-PARTY-NOTICES.md')]:
        copy(p, 'licenses/' + name)
    files = [{'file': str(p.relative_to(stage)), 'bytes': p.stat().st_size, 'sha256': sha(p)}
             for p in sorted(stage.rglob('*')) if p.is_file()]
    with tarfile.open(stage / 'evidence.tar', 'w') as tar:
        for row in files:
            tar.add(stage / row['file'], arcname=row['file'])
    archive = d / 'review-r11.tar.zst'
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'evidence.tar'), '-o', str(archive)], check=True)
    assert archive.stat().st_size < 95_000_000
    manifest = {'runtime_head': runtime, 'archive': archive.name, 'archive_bytes': archive.stat().st_size,
                'archive_sha256': sha(archive), 'files': files,
                'prior_archives': [{'file': p.name, 'sha256': sha(p)} for p in sorted(d.glob('*.tar.zst')) if p != archive]}
    (d / 'review-r11-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    acceptance = {'runtime_head': runtime, 'main': {'dotnet': 661, 'trx_files': 9, 'python': 5,
                  'default_packages': 11, 'optional_packages': 13}, 'independent_low13': verification,
                  'preview': preview, 'prior_r10_visual': '28 pages remain historical at 28fc29d; unchanged old pages not reopened',
                  'preserved_failures': ['unmodified R10 actual package P2 exceptions under both policies',
                      'two initial test-fixture float-boundary assumptions corrected',
                      'first Core test assumed MaxValue remained finite after fixed formatting; corrected',
                      'design review found and repaired budget precedence before final acceptance',
                      'Low13 first shell wrapper assigned readonly zsh status after tests; corrected wrapper reran no-build tests to fresh nine TRX and confirmed exit0'],
                  'final_head_reviews_and_ci': 'Tracked separately in PR18; not inherited',
                  'open_followup': {'thread': 'PRRT_kwDORQ4ZKs6rsyhK', 'reviewed_head': runtime,
                      'issue': 'Tf-selected fonts with no rendered glyph still load/embed and consume output font budget',
                      'status': 'open_pending_R12_design_fix_and_validation'},
                  'release_accepted': False, 'merged': False, 'published': False}
    (d / 'review-r11-acceptance.json').write_text(json.dumps(acceptance, ensure_ascii=False, indent=2) + '\n')
print(f'Archived {len(files)} files, {archive.stat().st_size} bytes; runtime {runtime}.')
