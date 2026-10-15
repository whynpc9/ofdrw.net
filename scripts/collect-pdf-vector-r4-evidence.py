#!/usr/bin/env python3
"""Append frozen PR18 R4 code and unresolved visual evidence without replacing the first acceptance archive."""
import argparse
import hashlib
import json
import shutil
import subprocess
import tarfile
import tempfile
from pathlib import Path

root = Path(__file__).resolve().parents[1]
p = argparse.ArgumentParser(description=__doc__)
p.add_argument('runtime_head')
p.add_argument('version')
a = p.parse_args()
subprocess.run(['git', 'diff', '--exit-code', a.runtime_head, '--', 'src', 'tests', 'e2e',
                '.github/workflows/ci.yml', 'scripts/run-pdf-vector-package-e2e.sh',
                'scripts/verify-pdf-vector-evidence.py'], cwd=root, check=True)
base = root / 'artifacts/pdf-vectors'
destination = root / 'docs/evidence/pdf-vectors'
low = base / 'independent'
preview = json.loads((base / 'preview-r4.json').read_text())
validation = json.loads((low / 'verification-low6.json').read_text())
assert preview['runtime_head'] == a.runtime_head and preview['gui_released'] and preview['checked_pages'] == 16
assert validation['source']['commit'] == a.runtime_head and validation['functional']['full_solution']['passed'] == 578
assert validation['functional']['observer']['status'] == 'passed'
sha = lambda path: hashlib.sha256(path.read_bytes()).hexdigest()
for document in preview['docs']:
    assert sha(Path(document['absolute_path'])) == document['sha256']
excluded = {'bin', 'obj', 'packages', 'feed', '__pycache__', 'venv', '.venv', 'dotnet-home', 'native'}
with tempfile.TemporaryDirectory(prefix='ofd-vector-review-') as temporary:
    stage = Path(temporary)
    def copy(source, target):
        source = Path(source)
        target = stage / target
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    def tree(source, target):
        source = Path(source)
        assert source.is_dir(), source
        for path in sorted(source.rglob('*')):
            relative = path.relative_to(source)
            if 'output-with-font-base' in relative.parts and path.name != 'font-base.pdf':
                continue  # These are duplicate recreations; original outputs and logs remain.
            if path.suffix == '.svg':
                continue  # Large base64 font duplication; actual OFD/PDF/PNG retain the geometry.
            if path.is_file() and not any(part in excluded for part in relative.parts):
                copy(path, Path(target) / relative)
    # Retain new affected fixtures; unchanged baseline payloads remain in the prior archives.
    for path in sorted((base / 'package-r4/output').iterdir()):
        if path.is_file() and (path.name.startswith(('precision', 'encoding-stream', 'cid-map-stream', 'indexed-image')) or path.name.endswith('-report.json') or path.name == 'verified-evidence.json'):
            copy(path, 'candidate/' + path.name)
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(base / 'package-r4' / name, 'packages/primary-' + name)
    copy(root / f'artifacts/package-e2e/{a.version}/packages/package-manifest.json', 'packages/default11.json')
    for product in ['Ofdrw.Net.Converter.Pdf.Vector', 'Ofdrw.Net.Graphics.SkiaSharp']:
        copy(base / f'package-r4/feed/{product}.{a.version}.nupkg', f'packages/{product}.{a.version}.nupkg')
    for directory in ['review-r4-design', 'review-r4-probe']:
        tree(base / directory, directory)
    for name in ['preview-r4.json', 'unchanged-baseline-r4.json', 'pr18-second-complete.json', 'pr18-package-ci-r3.log',
                 'full-review-r4.log', 'python-review-r4.log', 'default-package-r4.log', 'optional-package-r4.log',
                 'tests-review-r4-precision.log', 'tests-review-r4-precision-b.log', 'tests-review-r4-resources.log', 'tests-review-r4-c.log', 'standalone-tools-r4.log']:
        copy(base / name, 'acceptance/' + name)
    tree(base / 'full-review-r4', 'acceptance/trx')
    copy(base / 'standalone-r4/verified-evidence.json', 'acceptance/standalone-tools.json')
    for round in ['low6']:
        copy(low / f'verification-{round}.json', 'independent/' + f'verification-{round}.json')
        tree(low / f'logs-{round}', 'independent/' + f'logs-{round}')
        tree(low / f'observer-{round}', 'independent/' + f'observer-{round}')
    copy(low / 'source-hashes-low6.json', 'independent/source-hashes-low6.json')
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(low / 'optional13-low6' / name, 'independent/' + name)
    for name in ['sample-report.json', 'nonpainting-report.json', 'verified-evidence.json', 'image-hint-report.json', 'precision-resource-report.json']:
        copy(low / 'optional13-low6/output' / name, 'independent/' + name)
    copy(low / 'default-low6/packages/package-manifest.json', 'independent/default11.json')
    copy(root / 'LICENSE', 'licenses/MIT.txt')
    copy(root / 'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt', 'licenses/Noto-OFL.txt')
    copy(root / 'THIRD-PARTY-NOTICES.md', 'licenses/THIRD-PARTY-NOTICES.md')
    files = [{'file': str(path.relative_to(stage)), 'bytes': path.stat().st_size, 'sha256': sha(path)}
             for path in sorted(stage.rglob('*')) if path.is_file()]
    with tarfile.open(stage / 'review.tar', 'w') as archive:
        for item in files:
            archive.add(stage / item['file'], arcname=item['file'])
    archive = destination / 'review-r4.tar.zst'
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'review.tar'), '-o', str(archive)], check=True)
    assert archive.stat().st_size < 95_000_000
    manifest = {'runtime_head': a.runtime_head, 'version': a.version, 'archive': archive.name,
                'archive_bytes': archive.stat().st_size, 'archive_sha256': sha(archive), 'files': files,
                'prior_archive': 'evidence.tar.zst', 'prior_archive_sha256': sha(destination / 'evidence.tar.zst')}
    (destination / 'review-r4-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Archived {len(files)} repair evidence files, {archive.stat().st_size} bytes; first archive unchanged.')
