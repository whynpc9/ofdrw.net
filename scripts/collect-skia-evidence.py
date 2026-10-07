#!/usr/bin/env python3
"""Persist the bounded ticket 19 candidate evidence; never infer Preview acceptance."""
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tarfile
import tempfile

root = Path(__file__).resolve().parents[1]
source_head = sys.argv[1]
version = sys.argv[2]
sha = lambda data: hashlib.sha256(data).hexdigest()
destination = root / 'docs/evidence/skia'
destination.mkdir(parents=True, exist_ok=True)
with tempfile.TemporaryDirectory(prefix='ofd-skia-evidence-') as temporary:
    stage = Path(temporary)
    def copy(source, target):
        source = root / source
        target = stage / target
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(source, target)
    candidate = 'artifacts/skia/package-r2'
    for path in (root / candidate / 'output').iterdir():
        if path.is_file() and '-crop' not in path.name: copy(path.relative_to(root), 'candidate/' + path.name)
    copy(candidate + '/package-manifest.json', 'packages/primary-optional.json')
    copy(candidate + '/consumer.assets.json', 'packages/primary-consumer.assets.json')
    copy(f'artifacts/package-e2e/{version}/packages/package-manifest.json', 'packages/primary-default.json')
    copy('artifacts/skia/independent/independent-verification.json', 'independent/verification.json')
    for path in (root / 'artifacts/skia/independent/logs').iterdir():
        if path.is_file(): copy(path.relative_to(root), 'independent/logs/' + path.name)
    independent = root / 'artifacts/skia/independent/source-low2'
    for source, target in [
        ('artifacts/package-e2e/0.1.0-skia.20261008.low2/packages/package-manifest.json', 'independent/default-packages.json'),
        ('artifacts/skia/independent-package/package-manifest.json', 'independent/optional-packages.json'),
        ('artifacts/skia/independent-package/output/sample-report.json', 'independent/sample-report.json')]:
        copy((independent / source).relative_to(root), target)
    regression = root / 'artifacts/skia/graphics-regression-r2'
    for path in regression.rglob('*'):
        if path.is_file(): copy(path.relative_to(root), 'regression/' + str(path.relative_to(regression)))
    for filename in ['full-suite-frozen.log', 'python-tests.log', 'default-package-r1.log', 'optional-package-r1.log', 'default-package-r2.log', 'optional-package-r2.log', 'graphics-regression-r2.log']:
        copy('artifacts/skia/' + filename, 'logs/' + filename)
    for path in (root / 'artifacts/skia/full-suite-frozen').glob('*.trx'):
        copy(path.relative_to(root), 'logs/trx/' + path.name)
    notices = Path.home() / '.nuget/packages/skiasharp.nativeassets.macos/3.119.1'
    for filename in ['LICENSE.txt', 'THIRD-PARTY-NOTICES.txt']:
        copy(notices / filename, 'licenses/skia-native-' + filename)
    fonts = root / 'artifacts/graphics-fonts'
    copy(fonts / 'Ofdrw-CI-Noto-OFL.txt', 'licenses/Noto-OFL.txt')
    # Original source probe remains reviewable; SVG/PNG alone are not its code.
    copy('artifacts/skia/probe/Program.cs', 'probe/Program.cs')
    copy('artifacts/skia/probe/Probe.csproj', 'probe/Probe.csproj')
    for filename in ['src/Ofdrw.Net.Graphics.SkiaSharp/SkiaDrawEvent.cs', 'src/Ofdrw.Net.Graphics.SkiaSharp/OfdSkiaAdapter.cs']:
        if subprocess.check_output(['git', 'diff', source_head, '--', filename], cwd=root):
            raise ValueError('Candidate runtime has changed; regenerate acceptance artifacts first')
    files = [{'file': str(p.relative_to(stage)), 'bytes': p.stat().st_size, 'sha256': sha(p.read_bytes())}
             for p in sorted(stage.rglob('*')) if p.is_file()]
    archive = destination / 'evidence.tar.zst'
    with tarfile.open(stage / 'evidence.tar', 'w') as bundle:
        for file in files: bundle.add(stage / file['file'], arcname=file['file'])
    subprocess.run(['zstd', '-q', '-19', '--long=27', '-f', str(stage / 'evidence.tar'), '-o', str(archive)], check=True)
    manifest = {'source_head': source_head, 'runtime_head': '8bf97cfbf9ae16ccea5338d7a3a61fc6b1889644',
                'baseline': 'ab87bcade2ec5df6fa1f7ea0b86c590d13225653', 'version': version,
                'archive': archive.name, 'archive_bytes': archive.stat().st_size, 'archive_sha256': sha(archive.read_bytes()),
                'files': files}
    (destination / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
print(f"Persisted {len(files)} files, archive {archive.stat().st_size} bytes; Preview status is independently recorded.")
