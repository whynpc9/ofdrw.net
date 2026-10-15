#!/usr/bin/env python3
"""Append frozen PR18 repair evidence without replacing the first acceptance archive."""
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
preview = json.loads((base / 'preview-r3.json').read_text())
validation = json.loads((low / 'verification-low5.json').read_text())
assert preview['runtime_head'] == a.runtime_head and preview['gui_released'] and preview['checked_pages'] == 8
assert validation['source_commit'] == a.runtime_head and validation['gates']['full_solution']['passed'] == 542
assert validation['gates']['observer'] == 'passed'
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
            if path.is_file() and not any(part in excluded for part in relative.parts):
                copy(path, Path(target) / relative)
    tree(base / 'package-r3/output', 'candidate')
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(base / 'package-r3' / name, 'packages/primary-' + name)
    copy(root / f'artifacts/package-e2e/{a.version}/packages/package-manifest.json', 'packages/default11.json')
    for product in ['Ofdrw.Net.Converter.Pdf.Vector', 'Ofdrw.Net.Graphics.SkiaSharp']:
        copy(base / f'package-r3/feed/{product}.{a.version}.nupkg', f'packages/{product}.{a.version}.nupkg')
    for directory in ['review-r3-design', 'review-r3-probe', 'review-r3-tiny-triage']:
        tree(base / directory, directory)
    for name in ['preview-r3.json', 'unchanged-baseline-r3.json', 'pr18-first-complete.json', 'pr18-ci-first-failed.log',
                 'full-review-r3.log', 'python-review-r3.log', 'default-package-r3.log', 'optional-package-r3.log',
                 'tests-review-r3.log', 'tests-review-r3b.log', 'standalone-tools-r3.log']:
        copy(base / name, 'acceptance/' + name)
    tree(base / 'full-review-r3', 'acceptance/trx')
    copy(base / 'standalone-r3/verified-evidence.json', 'acceptance/standalone-tools.json')
    for round in ['low4', 'low5']:
        copy(low / f'verification-{round}.json', 'independent/' + f'verification-{round}.json')
        tree(low / f'logs-{round}', 'independent/' + f'logs-{round}')
        tree(low / f'observer-{round}', 'independent/' + f'observer-{round}')
    copy(low / 'source-hashes-low5.json', 'independent/source-hashes-low5.json')
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(low / 'optional13-low5' / name, 'independent/' + name)
    for name in ['sample-report.json', 'nonpainting-report.json', 'verified-evidence.json', 'image-hint-report.json', 'no-op-scan-payload-low5.png']:
        copy(low / 'optional13-low5/output' / name, 'independent/' + name)
    copy(low / 'default-low5/packages/package-manifest.json', 'independent/default11.json')
    copy(root / 'LICENSE', 'licenses/MIT.txt')
    copy(root / 'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt', 'licenses/Noto-OFL.txt')
    copy(root / 'THIRD-PARTY-NOTICES.md', 'licenses/THIRD-PARTY-NOTICES.md')
    files = [{'file': str(path.relative_to(stage)), 'bytes': path.stat().st_size, 'sha256': sha(path)}
             for path in sorted(stage.rglob('*')) if path.is_file()]
    with tarfile.open(stage / 'review.tar', 'w') as archive:
        for item in files:
            archive.add(stage / item['file'], arcname=item['file'])
    archive = destination / 'review-r3.tar.zst'
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'review.tar'), '-o', str(archive)], check=True)
    assert archive.stat().st_size < 95_000_000
    manifest = {'runtime_head': a.runtime_head, 'version': a.version, 'archive': archive.name,
                'archive_bytes': archive.stat().st_size, 'archive_sha256': sha(archive), 'files': files,
                'prior_archive': 'evidence.tar.zst', 'prior_archive_sha256': sha(destination / 'evidence.tar.zst')}
    (destination / 'review-r3-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
print(f'Archived {len(files)} repair evidence files, {archive.stat().st_size} bytes; first archive unchanged.')
