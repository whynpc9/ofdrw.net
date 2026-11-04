import contextlib
import importlib.util
import io
import json
import tempfile
import unittest
from pathlib import Path

spec = importlib.util.spec_from_file_location('advisories', Path(__file__).parents[1] / 'check-nuget-advisories.py')
advisories = importlib.util.module_from_spec(spec)
spec.loader.exec_module(advisories)


def report(*entries):
    packages = [{'id': package, 'resolvedVersion': version,
                 'vulnerabilities': [{'severity': 'High', 'advisoryurl': f'https://github.com/advisories/{identifier}'}]}
                for package, version, identifier in entries]
    return {'version': 1, 'projects': [{'path': 'a.csproj', 'frameworks': [
        {'framework': 'net10.0', 'topLevelPackages': packages[:1], 'transitivePackages': packages[1:]}]}]}


class NuGetAdvisoryTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.root = Path(self.temp.name)
        self.allowlist = self.root / 'allowlist.json'
        self.allowlist.write_text(json.dumps({'schema': 1, 'advisories': [
            {'package': 'Lib', 'version': '1.0.0', 'id': 'GHSA-aaaa', 'decision': 'D1', 'reference': 'doc.md'}]}))

    def tearDown(self):
        self.temp.cleanup()

    def run_gate(self, content):
        path = self.root / 'report.json'
        path.write_text(content if isinstance(content, str) else json.dumps(content))
        output = io.StringIO()
        with contextlib.redirect_stdout(output), contextlib.redirect_stderr(output):
            code = advisories.main(['--report', str(path), '--allowlist', str(self.allowlist)])
        return code, output.getvalue()

    def test_allowlisted_only_passes(self):
        self.assertEqual(0, self.run_gate(report(('Lib', '1.0.0', 'GHSA-aaaa')))[0])

    def test_unreviewed_advisory_fails_and_names_it(self):
        code, output = self.run_gate(report(('Lib', '1.0.0', 'GHSA-aaaa'), ('Other', '2.0.0', 'GHSA-bbbb')))
        self.assertEqual(1, code)
        self.assertIn('GHSA-bbbb', output)

    def test_same_advisory_on_another_version_is_not_allowlisted(self):
        self.assertEqual(1, self.run_gate(report(('Lib', '1.0.1', 'GHSA-aaaa')))[0])

    def test_stale_allowlist_entry_warns_without_failing(self):
        code, output = self.run_gate(report())
        self.assertEqual(0, code)
        self.assertIn('stale allowlist entry', output)

    def test_malformed_report_exits_2(self):
        self.assertEqual(2, self.run_gate('{not json')[0])
        self.assertEqual(2, self.run_gate({'version': 1})[0])


if __name__ == '__main__':
    unittest.main()
