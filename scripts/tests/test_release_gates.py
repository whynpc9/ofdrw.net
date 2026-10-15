import importlib.util
import hashlib
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


if __name__ == '__main__':
    unittest.main()
