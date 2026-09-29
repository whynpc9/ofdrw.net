import importlib.util
from pathlib import Path
import unittest

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


if __name__ == '__main__':
    unittest.main()
