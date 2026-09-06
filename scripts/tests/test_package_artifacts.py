import importlib.util
from pathlib import Path
import tempfile
import unittest
import zipfile

spec = importlib.util.spec_from_file_location('packages', Path(__file__).parents[1] / 'verify-package-artifacts.py')
packages = importlib.util.module_from_spec(spec)
spec.loader.exec_module(packages)


class PackageArtifactsTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.version = '0.1.0-fixture.1'
        for name in packages.PACKAGES:
            path = self.root / f'{name}.{self.version}.nupkg'
            with zipfile.ZipFile(path, 'w') as archive:
                archive.writestr(name + '.nuspec', f'<package><metadata><id>{name}</id><version>{self.version}</version></metadata></package>')
                for file in ('README.md', 'THIRD-PARTY-NOTICES.md', 'docs/feature-parity.md', 'docs/conversion-contracts.md'):
                    archive.writestr(file, 'fixture')
                if name == 'Ofdrw.Net.Cli':
                    archive.writestr('tools/net10.0/any/Ofdrw.Net.Cli.dll', b'fixture')
                elif name != 'Ofdrw.Net.Converter':
                    for framework in ('netstandard2.0', 'netstandard2.1'):
                        archive.writestr(f'lib/{framework}/{name}.dll', b'fixture')

    def tearDown(self):
        self.temp.cleanup()

    def test_complete_set_has_eleven_bound_hashes(self):
        manifest = packages.inspect(self.root, self.version)
        self.assertEqual(11, len(manifest['packages']))
        self.assertTrue(all(len(item['sha256']) == 64 for item in manifest['packages']))

    def test_missing_package_is_rejected(self):
        next(self.root.glob('*.nupkg')).unlink()
        with self.assertRaises(ValueError):
            packages.inspect(self.root, self.version)

    def test_stale_extra_version_is_rejected(self):
        (self.root / 'Ofdrw.Net.Core.0.0.0.nupkg').write_bytes(b'stale')
        with self.assertRaises(ValueError):
            packages.inspect(self.root, self.version)

    def test_altered_artifact_changes_bound_hash(self):
        before = packages.inspect(self.root, self.version)
        with zipfile.ZipFile(next(self.root.glob('*.nupkg')), 'a') as archive:
            archive.writestr('extra.txt', b'tampered')
        self.assertNotEqual(before, packages.inspect(self.root, self.version))

    def test_wrong_sdk_dependency_is_rejected(self):
        path = self.root / f'Ofdrw.Net.Core.{self.version}.nupkg'
        with zipfile.ZipFile(path) as archive:
            contents = {name: archive.read(name) for name in archive.namelist()}
        contents['Ofdrw.Net.Core.nuspec'] = f'<package><metadata><id>Ofdrw.Net.Core</id><version>{self.version}</version><dependencies><dependency id="Ofdrw.Net.Reader" version="0.0.0" /></dependencies></metadata></package>'.encode()
        with zipfile.ZipFile(path, 'w') as archive:
            for name, data in contents.items(): archive.writestr(name, data)
        with self.assertRaises(ValueError): packages.inspect(self.root, self.version)


if __name__ == '__main__': unittest.main()
