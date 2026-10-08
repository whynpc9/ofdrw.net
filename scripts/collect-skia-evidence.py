#!/usr/bin/env python3
"""Persist the bounded ticket 19 candidate evidence; never infer Preview acceptance."""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import subprocess
import sys
import tarfile
import tempfile

root = Path(__file__).resolve().parents[1]
parser = argparse.ArgumentParser(description=__doc__)
parser.add_argument('source_head'); parser.add_argument('version')
parser.add_argument('--round', default='r2'); parser.add_argument('--independent-round', default='low2')
parser.add_argument('--regression-round', default=None)
args = parser.parse_args()
source_head = args.source_head; version = args.version
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
    candidate = f'artifacts/skia/package-{args.round}'
    for path in (root / candidate / 'output').iterdir():
        if path.is_file() and '-crop' not in path.name: copy(path.relative_to(root), 'candidate/' + path.name)
    copy(candidate + '/package-manifest.json', 'packages/primary-optional.json')
    copy(candidate + '/consumer.assets.json', 'packages/primary-consumer.assets.json')
    copy(f'artifacts/package-e2e/{version}/packages/package-manifest.json', 'packages/primary-default.json')
    copy('artifacts/skia/independent/independent-verification.json', 'independent/verification.json')
    for path in (root / 'artifacts/skia/independent/logs').iterdir():
        if path.is_file(): copy(path.relative_to(root), 'independent/logs/' + path.name)
    independent = root / f'artifacts/skia/independent/source-{args.independent_round}'
    for source, target in [
        (f'artifacts/package-e2e/0.1.0-skia.20261008.{args.independent_round}/packages/package-manifest.json', 'independent/default-packages.json'),
        ('artifacts/skia/independent-package/package-manifest.json', 'independent/optional-packages.json'),
        ('artifacts/skia/independent-package/output/sample-report.json', 'independent/sample-report.json')]:
        copy((independent / source).relative_to(root), target)
    regression = root / f'artifacts/skia/graphics-regression-{args.regression_round or args.round}'
    for path in regression.rglob('*'):
        if path.is_file(): copy(path.relative_to(root), 'regression/' + str(path.relative_to(regression)))
    copy(f'artifacts/skia/preview-acceptance-{args.round}.json', 'preview/acceptance.json')
    copy('artifacts/skia/preview-acceptance-r3.json', 'preview/previous-r3.json')
    copy('artifacts/skia/independent/independent-verification-low3.json', 'independent/previous-low3.json')
    for filename in ['Program.cs', 'Probe.csproj', 'observed.txt']:
        copy('artifacts/skia/blender-probe/' + filename, 'blender-before/' + filename)
        copy('artifacts/skia/blender-probe-fixed/' + filename, 'blender-after/' + filename)
    copy('artifacts/skia/blender-probe/obj/project.assets.json', 'blender-before/consumer.assets.json')
    copy('artifacts/skia/blender-probe/build.log', 'blender-before/build.log')
    copy('artifacts/skia/pr17-review3-complete.json', 'reviews/pr17-review3-complete.json')
    copy('artifacts/skia/pr17-review2.json', 'reviews/pr17-review2.json')
    for filename in ['full-suite-blender-fixed.log', 'full-suite-review1-fixed.log', 'pr17-review1-complete.json', 'full-suite-frozen.log', 'python-tests.log', 'default-package-r1.log', 'optional-package-r1.log', 'default-package-r2.log', 'optional-package-r2.log', 'graphics-regression-r2.log', f'default-package-{args.round}.log', f'optional-package-{args.round}.log', f'graphics-regression-{args.regression_round or args.round}.log']:
        copy('artifacts/skia/' + filename, 'logs/' + filename)
    for path in (root / 'artifacts/skia/full-suite-blender-fixed').glob('*.trx'):
        copy(path.relative_to(root), 'logs/trx/' + path.name)
    notices = Path.home() / '.nuget/packages/skiasharp.nativeassets.macos/3.119.1'
    for filename in ['LICENSE.txt', 'THIRD-PARTY-NOTICES.txt']:
        copy(notices / filename, 'licenses/skia-native-' + filename)
    fonts = root / 'artifacts/graphics-fonts'
    copy(fonts / 'Ofdrw-CI-Noto-OFL.txt', 'licenses/Noto-OFL.txt')
    copy('artifacts/skia/independent/independent-verification-low2.json', 'independent/previous-low2.json')
    copy(f'artifacts/skia/independent/observer-{args.independent_round}/Program.cs', 'independent/observer/Program.cs')
    copy(f'artifacts/skia/independent/observer-{args.independent_round}/Observer.csproj', 'independent/observer/Observer.csproj')
    copy(f'artifacts/skia/independent/observer-{args.independent_round}/obj/project.assets.json', 'independent/observer/consumer.assets.json')
    copy((independent / 'artifacts/skia/independent-package/consumer.assets.json').relative_to(root), 'independent/consumer.assets.json')
    for filename in ['observed.txt', 'Program.cs', 'Probe.csproj']:
        copy('artifacts/skia/cancellation-probe/' + filename, 'review1-cancellation-before/' + filename)
        copy('artifacts/skia/cancellation-probe-fixed/' + filename, 'review1-cancellation-after/' + filename)
    if args.round == 'r5':
        copy('artifacts/skia/preview-acceptance-r4.json', 'preview/previous-r4.json')
        copy('artifacts/skia/independent/independent-verification-low4.json', 'independent/previous-low4.json')
        copy('artifacts/skia/pr17-review4-complete.json', 'reviews/pr17-review4-complete.json')
        for filename in ['full-suite-font-fixed.log', 'style-tests.log', 'style-tests-unique-fixtures.log', 'style-source-probe.log', 'style-source-probe2.log', 'optional-package-r5-roundtrip-id-failure.log']:
            copy('artifacts/skia/' + filename, 'font-style/logs/' + filename)
        for path in (root / 'artifacts/skia/full-suite-font-fixed').glob('*.trx'):
            copy(path.relative_to(root), 'font-style/trx/' + path.name)
        for path in (root / 'e2e/Ofdrw.Net.SkiaSharp.E2E/testdata/fonts').iterdir():
            if path.is_file(): copy(path.relative_to(root), 'font-style/fixtures/' + path.name)
        copy('scripts/generate-skia-style-fixtures.py', 'font-style/generate-skia-style-fixtures.py')
        copy('e2e/Ofdrw.Net.SkiaSharp.E2E/FontStyleProbe.cs', 'font-style/FontStyleProbe.cs')
        for filename in ['Program.cs', 'Probe.csproj', 'observed.txt', 'obj/project.assets.json']:
            copy('artifacts/skia/font-emphasis-before/' + filename, 'font-style/before/' + filename)
        for path in (root / 'artifacts/skia/font-emphasis-before/output').iterdir():
            if path.is_file(): copy(path.relative_to(root), 'font-style/before/output/' + path.name)
        for path in (independent / 'artifacts/skia/independent-package/output').glob('font-style*'):
            if path.is_file(): copy(path.relative_to(root), 'independent/' + path.name)
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
    manifest = {'source_head': source_head, 'runtime_head': source_head,
                'baseline': 'ab87bcade2ec5df6fa1f7ea0b86c590d13225653', 'version': version,
                'archive': archive.name, 'archive_bytes': archive.stat().st_size, 'archive_sha256': sha(archive.read_bytes()),
                'files': files}
    (destination / 'manifest.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n')
print(f"Persisted {len(files)} files, archive {archive.stat().st_size} bytes; Preview status is independently recorded.")
