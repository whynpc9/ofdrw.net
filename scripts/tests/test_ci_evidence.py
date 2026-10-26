import hashlib
import importlib.util
from pathlib import Path
import shutil
import subprocess
import tarfile
import tempfile
import unittest

SCRIPTS = Path(__file__).parents[1]


def load(name):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / f'{name}.py')
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


@unittest.skipUnless(shutil.which('zstd'), 'zstd required for evidence archive roundtrip')
class EvidenceArchiveTests(unittest.TestCase):
    def test_archive_preserves_nested_binary_hidden_and_empty_files(self):
        module = load('archive-ci-evidence')
        with tempfile.TemporaryDirectory() as temporary:
            root = Path(temporary)
            source = root / 'evidence'
            (source / 'pages').mkdir(parents=True)
            expected = {'pages/中文.png': bytes(range(256)) * 100, '.metadata': b'identity', 'empty': b''}
            for name, data in expected.items():
                (source / name).write_bytes(data)
            target = root / 'evidence.tar.zst'
            manifest = module.archive(source, target)
            tar_path = root / 'evidence.tar'
            subprocess.run(['zstd', '-q', '-d', str(target), '-o', str(tar_path)], check=True)
            with tarfile.open(tar_path) as archive:
                actual = {m.name.removeprefix('evidence/'): archive.extractfile(m).read()
                          for m in archive.getmembers() if m.isfile()}
            self.assertEqual(expected, actual)
            self.assertEqual(set(expected), {r['path'] for r in manifest['files']})
            for r in manifest['files']:
                self.assertEqual(hashlib.sha256(expected[r['path']]).hexdigest(), r['sha256'])

    def test_empty_or_recursive_output_is_rejected(self):
        module = load('archive-ci-evidence')
        with tempfile.TemporaryDirectory() as temporary:
            source = Path(temporary)
            with self.assertRaises(ValueError):
                module.archive(source, source / 'nested.tar.zst')
            with self.assertRaises(ValueError):
                module.archive(source, source.parent / 'empty.tar.zst')
