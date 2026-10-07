import importlib.util
import os
import shutil
import subprocess
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

    def test_product_list_failure_does_not_accept_existing_feed(self):
        # This feed is valid before the injected generator failure. It must not
        # be accepted as proof that the present pack operation succeeded.
        self.assertEqual(11, len(packages.inspect(self.root, self.version)['packages']))
        repo = self.root / 'repo'
        scripts = repo / 'scripts'
        scripts.mkdir(parents=True)
        source = Path(__file__).parents[1]
        shutil.copyfile(source / 'run-converter-package-e2e.sh', scripts / 'run-converter-package-e2e.sh')
        marker = self.root / 'validator-ran'
        validator = (source / 'verify-package-artifacts.py').read_text()
        validator += "\nif __name__ != '__main__':\n    raise RuntimeError('injected product list failure')\n"
        validator = validator.replace("if __name__ == '__main__':", "if __name__ == '__main__':\n    Path(" + repr(str(marker)) + ").write_text('called')")
        (scripts / 'verify-package-artifacts.py').write_text(validator)
        tools = self.root / 'tools'
        tools.mkdir()
        dotnet = tools / 'dotnet'
        calls = self.root / 'dotnet-calls'
        dotnet.write_text("#!/usr/bin/env bash\nprintf '%s\\n' \"$*\" >> \"$TASK_CALLS\"\nexit 0\n")
        dotnet.chmod(0o755)
        env = dict(os.environ, PATH=str(tools) + os.pathsep + os.environ['PATH'],
                   TASK_CALLS=str(calls), NUGET_PACKAGES=str(self.root / 'cache'),
                   DOTNET_CLI_HOME=str(self.root / 'dotnet-home'),
                   DOTNET_SKIP_FIRST_TIME_EXPERIENCE='1', DOTNET_CLI_TELEMETRY_OPTOUT='1')
        result = subprocess.run(['bash', str(scripts / 'run-converter-package-e2e.sh'), self.version,
                                 '--packages-dir', str(self.root), '--output-dir', str(self.root / 'output')],
                                env=env, capture_output=True, text=True)
        self.assertNotEqual(0, result.returncode)
        self.assertIn('injected product list failure', result.stderr)
        self.assertFalse(marker.exists(), 'Generator failure must stop before accepting stale packages')
        self.assertEqual(['restore', 'build'], [line.split()[0] for line in calls.read_text().splitlines()])


if __name__ == '__main__': unittest.main()
