#!/usr/bin/env python3
"""Archive R8 gates, Low10 invocation failure, Low11 completion and R9 design only."""
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
low = b / 'independent'
head = 'd8d197ec448c5dad690a72a1bb5a776c2e7019e2'
version = '0.1.0-pdfvector.20261009.r8'
sha = lambda p: hashlib.sha256(p.read_bytes()).hexdigest()
subprocess.run(['git', 'diff', '--exit-code', head, '--', 'src', 'tests', 'e2e',
                'scripts/run-pdf-vector-package-e2e.sh'], cwd=root, check=True)
failed = json.loads((low / 'verification-low10.json').read_text())
completed = json.loads((low / 'verification-low11.json').read_text())
assert failed['exact_head'] == completed['exact_head'] == head
assert failed['status'] == 'failed_first_actual_gate'
assert failed['gates']['dotnet_total'] == 625
assert completed['status'] == 'passed_bounded_completion'
assert completed['immutability']['reused_inputs_sources_binaries_nupkgs_changes'] == []
design = b / 'font-identity-r9-design'
for row in json.loads((design / 'evidence-manifest.json').read_text()):
    path = design / row['path']
    assert path.stat().st_size == row['bytes'] and sha(path) == row['sha256'], row
excluded = {'bin', 'obj', 'packages', 'feed', '__pycache__', 'dotnet-home', 'tmp',
            'dependency-feed', 'native', 'font-cache'}
with tempfile.TemporaryDirectory(prefix='ofd-r8-history-') as tmp:
    stage = Path(tmp)

    def copy(source, target):
        target = stage / target
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)

    def tree(source, target):
        for path in sorted(source.rglob('*')):
            rel = path.relative_to(source)
            if path.is_file() and path.suffix != '.svg' and not any(p in excluded for p in rel.parts):
                copy(path, Path(target) / rel)

    for name in ['full-r8.log', 'python-r8.log', 'default-package-r8.log',
                 'optional-package-r8.log', 'unchanged-baseline-r8.json',
                 'preview-r8-ready.json', 'pr18-state-resume-oct13.json']:
        copy(b / name, 'main/' + name)
    tree(b / 'full-r8-trx', 'main/trx')
    for name in ['package-manifest.json', 'consumer.assets.json']:
        copy(b / 'package-r8' / name, 'main/' + name)
        copy(low / 'optional13-low10' / name, 'independent/' + name)
    copy(root / f'artifacts/package-e2e/{version}/packages/package-manifest.json', 'main/default11.json')
    copy(low / 'default-low10/packages/package-manifest.json', 'independent/default11.json')
    for path in sorted((b / 'package-r8/output').iterdir()):
        if path.is_file() and (path.name.startswith(('text-precision', 'indexed-image-clipped'))
                              or path.suffix == '.json'):
            copy(path, 'candidate/' + path.name)
    for product in ['Ofdrw.Net.Converter.Pdf.Vector', 'Ofdrw.Net.Graphics.SkiaSharp']:
        copy(b / f'package-r8/feed/{product}.{version}.nupkg', f'packages/{product}.{version}.nupkg')
    for name in ['verification-low10.json', 'verification-low11.json', 'source-hashes-low10.json',
                 'hashes-low11-before.json', 'model-compare-low11.json']:
        copy(low / name, 'independent/' + name)
    for name in ['logs-low10', 'test-results-low10', 'observer-low10', 'observer-original-low10',
                 'observer-output-low10', 'observer-output-low11', 'auxiliary-low11']:
        tree(low / name, 'independent/' + name)
    tree(design, 'font-identity-r9-design')
    for source, name in [(root / 'LICENSE', 'MIT.txt'),
                         (root / 'artifacts/graphics-fonts/Ofdrw-CI-Noto-OFL.txt', 'Noto-OFL.txt'),
                         (root / 'THIRD-PARTY-NOTICES.md', 'THIRD-PARTY-NOTICES.md')]:
        copy(source, 'licenses/' + name)
    files = [{'file': str(p.relative_to(stage)), 'bytes': p.stat().st_size, 'sha256': sha(p)}
             for p in sorted(stage.rglob('*')) if p.is_file()]
    with tarfile.open(stage / 'evidence.tar', 'w') as tar:
        for row in files:
            tar.add(stage / row['file'], arcname=row['file'])
    archive = d / 'review-r8.tar.zst'
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'evidence.tar'),
                    '-o', str(archive)], check=True)
    assert archive.stat().st_size < 95_000_000
    manifest = {'runtime_head': head, 'production_src_same_as': 'deda6d2928c2db0e05640d2df96e50b79131ccfa',
                'version': version, 'archive': archive.name, 'archive_bytes': archive.stat().st_size,
                'archive_sha256': sha(archive), 'files': files,
                'status': 'R8 functional gates passed; Low10 pre-entry failure retained; Low11 bounded completion; font P1 open; no R8 Preview',
                'prior_archives': [{'file': p.name, 'sha256': sha(p)} for p in sorted(d.glob('*.tar.zst'))
                                   if p != archive]}
    (d / 'review-r8-manifest.json').write_text(json.dumps(manifest, indent=2) + '\n')
    acceptance = {'head': head, 'main': {'dotnet': 625, 'python': 5, 'default_packages': 11, 'optional_packages': 13},
                  'independent_low10_failure': failed, 'independent_low11_completion': completed,
                  'visual': {'r8_preview_performed': False, 'r6_original_observations_retained': True,
                             'new_pdf_bytes_are_not_previously_viewed': True},
                  'font_performance': {'status': 'confirmed finding, design only; repair not implemented',
                                       'font_bytes': 21659544, 'bounded_runs': [1, 10, 50, 100],
                                       'actual_r8_package_timings': True,
                                       'stream_counts': 'artifact-only diagnostic source copy; actual package counts derived from loaded IL',
                                       'report': 'font-identity-r9-design/report.md'},
                  'release_accepted': False, 'merged': False, 'published': False}
    (d / 'review-r8-acceptance.json').write_text(json.dumps(acceptance, indent=2) + '\n')
print(f'Archived {len(files)} files, {archive.stat().st_size} bytes; R9 is design only.')
