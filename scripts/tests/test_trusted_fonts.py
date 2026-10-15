import copy
import hashlib
import importlib.util
import io
import json
import os
from pathlib import Path
import shutil
import stat
import struct
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import patch
import zipfile

ROOT = Path(__file__).resolve().parents[2]
loader = importlib.util.spec_from_file_location('prepare_fonts', ROOT / 'scripts/prepare-trusted-fonts.py')
prepare_fonts = importlib.util.module_from_spec(loader)
loader.loader.exec_module(prepare_fonts)
MANIFEST = ROOT / 'tests/fixtures/trusted-fonts/bundle.json'
SPEC = prepare_fonts.read_spec(MANIFEST)
BUNDLE = MANIFEST.parent / SPEC['archive']['file']


class TrustedFontPreparationTests(unittest.TestCase):
    def test_source_free_standalone_preparation_and_verified_reuse(self):
        with tempfile.TemporaryDirectory(prefix='trusted-font-independent-') as name:
            root = Path(name)
            shutil.copy2(ROOT / 'scripts/prepare-trusted-fonts.py', root / 'prepare.py')
            shutil.copy2(MANIFEST, root / 'bundle.json')
            shutil.copy2(BUNDLE, root / 'bundle.zip')
            command = [sys.executable, str(root / 'prepare.py'), '--manifest', str(root / 'bundle.json'),
                       '--bundle', str(root / 'bundle.zip'), '--output-dir', str(root / 'out')]
            env = {k: v for k, v in os.environ.items() if not k.startswith('OFDRW_')}
            first = json.loads(subprocess.check_output(command, cwd=root, env=env, text=True))
            second = json.loads(subprocess.check_output(command, cwd=root, env=env, text=True))
            self.assertEqual(first, second)
            self.assertEqual(first['environment']['OFDRW_REQUIRE_TRUSTED_FONTS'], '1')
            prepare_fonts.verify_prepared(Path(first['directory']), SPEC)
            self.assertTrue(all(str(root) in value for key, value in first['environment'].items() if key.endswith('_FONT')))

    def test_corrupt_archive_or_manifest_fails_without_published_directory(self):
        with tempfile.TemporaryDirectory() as name:
            root = Path(name)
            bad = root / 'bad.zip'
            data = bytearray(BUNDLE.read_bytes()); data[100] ^= 1; bad.write_bytes(data)
            with self.assertRaises(ValueError):
                prepare_fonts.prepare(MANIFEST, bad, root / 'out')
            self.assertEqual(list((root / 'out').iterdir()), [])
            manifest = root / 'bad.json'; manifest.write_bytes(MANIFEST.read_bytes() + b' ')
            with self.assertRaises(ValueError):
                prepare_fonts.prepare(manifest, BUNDLE, root / 'not-created')
            self.assertFalse((root / 'not-created').exists())

    def test_actual_central_record_count_is_bounded_before_zipfile(self):
        data = bytearray(BUNDLE.read_bytes())
        at = data.rfind(b'PK\x05\x06')
        struct.pack_into('<HH', data, at + 8, 7, 7)
        with self.assertRaises(ValueError):
            prepare_fonts.preflight_directory(io.BytesIO(data), len(data))
        # Declaring six but containing a seventh actual record is also refused.
        with tempfile.TemporaryFile() as stream:
            with zipfile.ZipFile(stream, 'w') as z:
                for i in range(7): z.writestr(str(i), b'x')
            stream.seek(0); data = bytearray(stream.read())
        at = data.rfind(b'PK\x05\x06'); struct.pack_into('<HH', data, at + 8, 6, 6)
        with self.assertRaises(ValueError):
            prepare_fonts.preflight_directory(io.BytesIO(data), len(data))

    def test_member_path_kind_size_ratio_and_duplicate_guards(self):
        with zipfile.ZipFile(BUNDLE) as source:
            originals = source.infolist()
        class Directory:
            def __init__(self, infos): self.infos = infos
            def infolist(self): return self.infos
        for mutation in ('traversal', 'backslash', 'absolute', 'duplicate', 'symlink', 'size', 'ratio', 'encrypted', 'count'):
            with self.subTest(mutation=mutation):
                infos = [copy.copy(i) for i in originals]
                if mutation in ('traversal', 'backslash', 'absolute'):
                    value = {'traversal': '../escape', 'backslash': 'fonts\\NotoSans-Regular.ttf', 'absolute': '/escape'}[mutation]
                    infos[0].filename = infos[0].orig_filename = value
                if mutation == 'duplicate': infos[1] = copy.copy(infos[0])
                if mutation == 'symlink': infos[0].external_attr = (stat.S_IFLNK | 0o777) << 16
                if mutation == 'size': infos[0].file_size += 1
                if mutation == 'ratio': infos[0].compress_size = 1
                if mutation == 'encrypted': infos[0].flag_bits |= 1
                if mutation == 'count': infos.append(copy.copy(infos[0]))
                with self.assertRaises(ValueError): prepare_fonts.validate_members(Directory(infos), SPEC)

    def test_partial_copy_error_and_interrupt_never_publish(self):
        for error in (OSError('injected disk failure'), KeyboardInterrupt()):
            with self.subTest(error=type(error).__name__), tempfile.TemporaryDirectory() as name:
                output = Path(name) / 'out'
                original = prepare_fonts.copy_member
                def fail(*args):
                    original(*args)
                    raise error
                with patch.object(prepare_fonts, 'copy_member', side_effect=fail):
                    with self.assertRaises(type(error)):
                        prepare_fonts.prepare(MANIFEST, BUNDLE, output)
                self.assertEqual(list(output.iterdir()), [])

    def test_corrupt_existing_cache_is_refused_and_not_overwritten(self):
        with tempfile.TemporaryDirectory() as name:
            output = Path(name)
            prepared = prepare_fonts.prepare(MANIFEST, BUNDLE, output)
            target = Path(prepared['directory'])
            file = target / 'NOTICE.md'; file.write_bytes(b'corrupt existing cache')
            with self.assertRaises(ValueError): prepare_fonts.prepare(MANIFEST, BUNDLE, output)
            self.assertEqual(file.read_bytes(), b'corrupt existing cache')
            self.assertFalse(any(p.name.startswith('.trusted-fonts-') for p in output.iterdir()))

    def test_streamed_payload_sha_guard_is_not_only_archive_metadata(self):
        with tempfile.TemporaryDirectory() as name, zipfile.ZipFile(BUNDLE) as source:
            info = source.getinfo('NOTICE.md')
            wrong = dict(next(m for m in SPEC['members'] if m['path'] == 'NOTICE.md'), sha256='0' * 64)
            with self.assertRaises(ValueError):
                prepare_fonts.copy_member(source, info, Path(name) / 'NOTICE.md', wrong)


if __name__ == '__main__':
    unittest.main()
