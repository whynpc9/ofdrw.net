"""Preflight of run-ses-interop-e2e.sh candidate mode: it must reject feeds before any dotnet work."""
import importlib.util
import json
import subprocess
import tempfile
import unittest
import zipfile
from pathlib import Path

SCRIPTS = Path(__file__).parents[1]
SCRIPT = SCRIPTS / 'run-ses-interop-e2e.sh'
spec = importlib.util.spec_from_file_location('packages', SCRIPTS / 'verify-package-artifacts.py')
packages = importlib.util.module_from_spec(spec)
spec.loader.exec_module(packages)
VERSION = '0.1.0-fixture.1'


def nupkg(path, package_id, version, dependencies=(), extra=()):
    deps = ''.join(f'<dependency id="{i}" version="{v}" />' for i, v in dependencies)
    with zipfile.ZipFile(path, 'w') as archive:
        archive.writestr(package_id + '.nuspec', f'<package><metadata><id>{package_id}</id><version>{version}</version>'
                         f'<license type="expression">MIT</license><dependencies>{deps}</dependencies></metadata></package>')
        for name, data in extra:
            archive.writestr(name, data)


class SesInteropCandidateTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.sdk, self.optional, self.output = self.root / 'sdk', self.root / 'optional', self.root / 'evidence'
        self.sdk.mkdir()
        self.optional.mkdir()
        for name in packages.PACKAGES:
            files = [(f, (packages.ROOT / f).read_bytes() if f == 'THIRD-PARTY-NOTICES.md' else b'fixture')
                     for f in ('README.md', 'THIRD-PARTY-NOTICES.md', 'docs/feature-parity.md', 'docs/conversion-contracts.md')]
            if name in packages.NATIVE_SUPERVISOR_PACKAGES:
                files.append((packages.MUSL_COPYRIGHT_ENTRY, (packages.ROOT / packages.MUSL_COPYRIGHT_ENTRY).read_bytes()))
            if name == 'Ofdrw.Net.Cli':
                files.append(('tools/net10.0/any/Ofdrw.Net.Cli.dll', b'fixture'))
            elif name != 'Ofdrw.Net.Converter':
                files += [(f'lib/{tfm}/{name}{ext}', b'fixture') for tfm in ('netstandard2.0', 'netstandard2.1') for ext in ('.dll', '.xml')]
            nupkg(self.sdk / f'{name}.{VERSION}.nupkg', name, VERSION, extra=files)

    def tearDown(self):
        self.temp.cleanup()

    def run_script(self, *options):
        return subprocess.run(['bash', str(SCRIPT), str(self.output), VERSION, *options],
                              capture_output=True, text=True, timeout=120)

    def candidate(self):
        return ['--packages-dir', str(self.sdk), '--optional-packages-dir', str(self.optional)]

    def write_manifest(self):
        (self.sdk / 'package-manifest.json').write_text(json.dumps(packages.inspect(self.sdk, VERSION), indent=2) + '\n')

    def test_both_candidate_feeds_are_required(self):
        result = self.run_script('--packages-dir', str(self.sdk))
        self.assertEqual(1, result.returncode)
        self.assertIn('Candidate mode requires', result.stderr)
        self.assertFalse(self.output.exists())

    def test_candidate_without_manifest_is_rejected(self):
        result = self.run_script(*self.candidate())
        self.assertEqual(1, result.returncode)
        self.assertIn('lacks package-manifest.json', result.stderr)
        self.assertFalse(self.output.exists())

    def test_optional_feed_must_hold_the_exact_ses_package(self):
        self.write_manifest()
        ses = 'Ofdrw.Net.Signatures.SesInterop'
        cases = {
            'missing': [],
            'other version': [(f'{ses}.{VERSION}.nupkg', '0.0.1', f'[{VERSION}]')],
            'loose sibling range': [(f'{ses}.{VERSION}.nupkg', VERSION, VERSION)],
            'unexpected package': [(f'{ses}.{VERSION}.nupkg', VERSION, f'[{VERSION}]'), (f'Ofdrw.Net.Core.{VERSION}.nupkg', VERSION, None)],
        }
        for label, files in cases.items():
            with self.subTest(label):
                for stale in self.optional.iterdir():
                    stale.unlink()
                for file, version, sibling in files:
                    package_id = file[:-len(f'.{VERSION}.nupkg')]
                    nupkg(self.optional / file, package_id, version,
                          [('Ofdrw.Net.Signatures', sibling)] if sibling else [])
                result = self.run_script(*self.candidate())
                self.assertEqual(1, result.returncode, result.stdout + result.stderr)
                self.assertRegex(result.stderr, 'Optional (feed|package)')
                self.assertEqual([], [p for p in (self.output / 'packages').rglob('*') if p.is_file()])


if __name__ == '__main__':
    unittest.main()
