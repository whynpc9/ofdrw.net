import importlib.util
import hashlib
import subprocess
import tempfile
from pathlib import Path
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).parents[1]


def load(name, file):
    spec = importlib.util.spec_from_file_location(name, SCRIPTS / file)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class ReleaseGatesTests(unittest.TestCase):
    def test_public_api_warning_uses_symbol_not_line_number(self):
        gate = load('public_docs', 'check-public-api-docs.py')
        file = gate.ROOT / 'src/Ofdrw.Net.Core/Models/NewMember.cs'
        lines = '\n'.join(
            f"{file}({line},1): warning CS1591: Missing XML comment for publicly visible type or member 'NewMember.Value'"
            for line in (5, 99))
        self.assertEqual(['src/Ofdrw.Net.Core/Models/NewMember.cs|NewMember.Value'], gate.parse(lines))
        self.assertEqual([], gate.unparsed_diagnostics(lines))
        self.assertEqual(['warning CS1591: unexpected localized output'],
                         gate.unparsed_diagnostics('warning CS1591: unexpected localized output'))

    def test_unreviewed_preview_record_blocks_release(self):
        gate = load('preview_record', 'verify-preview-record.py')
        with self.assertRaisesRegex(ValueError, 'not accepted'):
            gate.verify({'status': 'not-reviewed'}, '0.1.0-preview.8')

    def test_native_only_preview_record_cannot_approve_default_mode(self):
        gate = load('preview_record_default', 'verify-preview-record.py')
        sample = gate.ROOT / 'e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx'
        data = {'status': 'accepted', 'package_version': '0.1.0-preview.8',
                'source_fingerprint': 'fixture', 'sample_sha256': hashlib.sha256(sample.read_bytes()).hexdigest(),
                'native_ofd_sha256': '0' * 64, 'preview_pdf_sha256': '1' * 64,
                'conversion_mode': 'native', 'view_chain': 'DOCX -> native OFD -> PDF -> macOS Preview',
                'reviewer': 'fixture', 'reviewed_at': '2026-09-29', 'checked_pages': [1, 2], 'open_findings': []}
        with patch.object(gate, 'source_fingerprint', return_value='fixture'):
            with self.assertRaisesRegex(ValueError, 'default_ofd_sha256'):
                gate.verify(data, '0.1.0-preview.8')


class PreviewFingerprintTests(unittest.TestCase):
    def setUp(self):
        self.directory = tempfile.TemporaryDirectory()
        self.root = Path(self.directory.name)
        self.gate = load('preview_fingerprint', 'verify-preview-record.py')
        self.root_patch = patch.object(self.gate, 'ROOT', self.root)
        self.root_patch.start()
        self.git('init', '--quiet')
        (self.root / '.gitattributes').write_text('*.pdf filter=lfs -text\n')
        (self.root / 'README.md').write_text('original source\n')
        self.git('add', '.gitattributes', 'README.md')
        self.name = 'docs/验收 pages.pdf'
        self.path = self.root / self.name
        self.path.parent.mkdir()
        self.payload = b'fixture document contents\n'
        self.pointer = self.make_pointer(self.payload)
        self.stage(self.name, self.pointer)
        self.path.write_bytes(self.pointer)

    def tearDown(self):
        self.root_patch.stop()
        self.directory.cleanup()

    def git(self, *args, data=None):
        return subprocess.check_output(['git', *args], cwd=self.root, input=data, stderr=subprocess.PIPE)

    def stage(self, name, data):
        oid = self.git('hash-object', '-w', '--stdin', data=data).decode().strip()
        self.git('update-index', '--add', '--cacheinfo', '100644', oid, name)

    @staticmethod
    def make_pointer(payload):
        return (f'version https://git-lfs.github.com/spec/v1\noid sha256:{hashlib.sha256(payload).hexdigest()}\n'
                f'size {len(payload)}\n').encode()

    def test_pointer_and_verified_hydrated_file_have_same_fingerprint(self):
        expected = self.gate.source_fingerprint()
        self.path.write_bytes(self.payload)
        self.assertEqual(expected, self.gate.source_fingerprint())
        self.path.write_bytes(self.pointer.rstrip(b'\n'))
        self.assertEqual(expected, self.gate.source_fingerprint())

    def test_tampered_payload_size_or_hash_is_rejected(self):
        for content, message in [(self.payload + b'x', 'payload size'),
                                 (b'X' + self.payload[1:], 'payload hash')]:
            with self.subTest(message=message):
                self.path.write_bytes(content)
                with self.assertRaisesRegex(ValueError, message):
                    self.gate.source_fingerprint()

    def test_different_working_pointer_is_rejected(self):
        self.path.write_bytes(self.make_pointer(b'different document'))
        with self.assertRaisesRegex(ValueError, 'working pointer differs'):
            self.gate.source_fingerprint()

    def test_raw_staged_lfs_payload_is_rejected(self):
        self.stage(self.name, self.payload)
        self.path.write_bytes(self.payload)
        with self.assertRaisesRegex(ValueError, 'valid Git LFS pointer'):
            self.gate.source_fingerprint()

    def test_unstaged_ordinary_source_edit_changes_fingerprint(self):
        before = self.gate.source_fingerprint()
        (self.root / 'README.md').write_text('changed source\n')
        self.assertNotEqual(before, self.gate.source_fingerprint())

    def test_unmerged_index_path_is_rejected(self):
        oid = self.git('hash-object', '-w', '--stdin', data=self.pointer).decode().strip()
        self.git('update-index', '--force-remove', self.name)
        self.git('update-index', '--index-info',
                 data=f'100644 {oid} 1\t{self.name}\n100644 {oid} 2\t{self.name}\n'.encode())
        with self.assertRaisesRegex(ValueError, 'unmerged path'):
            self.gate.source_fingerprint()

    def test_staged_new_lfs_identity_changes_fingerprint_and_survives_commit(self):
        before = self.gate.source_fingerprint()
        payload = b'new reviewed document bytes\n'
        self.stage(self.name, self.make_pointer(payload))
        self.path.write_bytes(payload)
        after = self.gate.source_fingerprint()
        self.assertNotEqual(before, after)
        self.git('-c', 'user.name=Release gate fixture', '-c', 'user.email=fixture@example.invalid',
                 '-c', 'commit.gpgSign=false', '-c', 'core.hooksPath=/dev/null', 'commit', '-qm', 'fixture')
        self.assertEqual(after, self.gate.source_fingerprint())


if __name__ == '__main__':
    unittest.main()
