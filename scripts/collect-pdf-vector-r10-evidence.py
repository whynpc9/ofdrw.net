#!/usr/bin/env python3
"""Preserve R9 failures, R10 package evidence, Low12 and actual Preview inspection."""
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
runtime = '28fc29daeab5e93b0afa5da2af504a1534ee8116'
version = '0.1.0-pdfvector.20261013.r10'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
subprocess.run(['git', 'diff', '--exit-code', runtime, '--', 'src', 'tests', 'e2e'], cwd=root, check=True)
low = b / 'independent/low12'
verification = json.loads((low / 'verification.json').read_text())
assert verification['runtimeHead'] == runtime
assert verification['status'] == 'passed_with_explicit_visual_limit'
for row in verification['artifactInventory']:
    p = Path(row['path'])
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256'], row
preview = json.loads((b / 'preview-r10.json').read_text())
assert preview['status'] == 'passed_bounded_visual_inspection'
assert preview['runtime_head'] == runtime and preview['total_pages'] == 28
for row in preview['docs']:
    p = Path(row['absolute_path'])
    assert row['checked_pages'] == list(range(1, row['pages'] + 1))
    assert p.stat().st_size == row['bytes'] and sha(p) == row['sha256'], row
excluded = {'bin', 'obj', 'packages', 'package-cache', 'cache', 'feed', '__pycache__',
            'dotnet-home', 'tmp', 'dependency-feed', 'native', 'font-cache', '.git', 'source'}
with tempfile.TemporaryDirectory(prefix='ofd-r10-evidence-') as tmp:
    stage = Path(tmp)

    def copy(p, target):
        q = stage / target
        q.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(p, q)

    def tree(source, target):
        for p in sorted(source.rglob('*')):
            rel = p.relative_to(source)
            if p.is_file() and p.suffix != '.svg' and not any(x in excluded for x in rel.parts):
                copy(p, Path(target) / rel)

    for pattern in ['*-r9.log', '*-r10.log', 'source-r9-*.log', 'source-r9-freeze.json',
                    'source-r10-*.json', 'model-comparison-r10-first-error.json',
                    'unchanged-baseline-r10.json', 'preview-r10*.json',
                    'low12-parent-settlement.json', 'pr18-state-r10-before-gui.json']:
        for p in b.glob(pattern):
            if p.is_file():
                copy(p, 'main/' + p.name)
    tree(b / 'full-r9-trx', 'main/preliminary-r9-trx')
    tree(b / 'full-r10-trx', 'main/r10-trx')
    tree(b / 'font-identity-r9-measure', 'measurement/preliminary-r9')
    tree(b / 'font-identity-r10-measure', 'measurement/r10')
    tree(low, 'independent/low12')
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(b / 'package-r10' / name, 'main/optional13/' + name)
    copy(root / f'artifacts/package-e2e/{version}/packages/package-manifest.json', 'main/default11.json')
    copy(low / 'source/artifacts/package-e2e/0.1.0-pdfvector.20261013.low12/packages/package-manifest.json',
         'independent/default11.json')
    for p in sorted((b / 'package-r10/output').iterdir()):
        if p.suffix in {'.ofd', '.json', '.png'} or any(p == Path(r['absolute_path']) for r in preview['docs']):
            copy(p, 'candidate/' + p.name)
    for label, feed, ver in [('main', b / 'package-r10/feed', version),
                             ('independent', low / 'optional/feed', '0.1.0-pdfvector.20261013.low12')]:
        for product in ['Ofdrw.Net.Converter.Pdf.Vector', 'Ofdrw.Net.Graphics.SkiaSharp']:
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
    archive = d / 'review-r10.tar.zst'
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'evidence.tar'), '-o', str(archive)], check=True)
    assert archive.stat().st_size < 95_000_000
    manifest = {'runtime_head': runtime, 'delivery_source_freeze': verification['head'],
                'version': version, 'archive': archive.name, 'archive_bytes': archive.stat().st_size,
                'archive_sha256': sha(archive), 'files': files,
                'prior_archives': [{'file': p.name, 'sha256': sha(p)} for p in sorted(d.glob('*.tar.zst')) if p != archive],
                'status': 'Main and independent functional gates passed; actual 28-page Preview passed within stated scope; final-head CI/reviews tracked separately in PR18'}
    (d / 'review-r10-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    acceptance = {'runtime_head': runtime, 'main': {'dotnet': 641, 'trx_files': 9, 'python': 5,
                  'default_packages': 11, 'optional_packages': 13}, 'independent_low12': verification,
                  'preview': preview, 'measurement_report': 'measurement/r10/report.md',
                  'preserved_failures': ['R9 initial CS0122 build', 'R9 unnecessary public stream comparison slowdown',
                                         'R10 first model comparison expected 20 OFDs instead of actual 21',
                                         'Preview open-panel misselection, cleanup URLs and unproven temporary PostScript ownership'],
                  'final_head_reviews_and_ci': 'Tracked separately in PR18; not inherited from an older commit',
                  'release_accepted': False, 'merged': False, 'published': False}
    (d / 'review-r10-acceptance.json').write_text(json.dumps(acceptance, ensure_ascii=False, indent=2) + '\n')
print(f'Archived {len(files)} files, {archive.stat().st_size} bytes; runtime {runtime}.')
