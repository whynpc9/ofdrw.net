import json
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
import xml.etree.ElementTree as ET


CHECKER = Path(__file__).resolve().parents[1] / "validation/upstream-2026-11/check-test-results.py"
ACCOUNTING = "Ofdrw.Net.Packaging.Tests.ControlledReaderAccountingTests"
PROJECTS = ("Packaging", "Cli")


def new_test(project):
    return {
        "project": project,
        "item": "M02",
        "className": f"Ofdrw.Net.{project}.Tests.Upstream202611Fixture",
        "methodName": "OpensArchive",
        "displayName": 'OpensArchive(layout: "adjusted")',
    }


class UpstreamTestResultsTests(unittest.TestCase):
    def setUp(self):
        temp = tempfile.TemporaryDirectory()
        self.addCleanup(temp.cleanup)
        self.root = Path(temp.name)
        self.manifest = self.root / "manifest.json"
        self.manifest.write_text(json.dumps({"prs": {"A": {
            "fullProjects": list(PROJECTS), "newTests": [new_test(p) for p in PROJECTS],
        }}}))
        for project in PROJECTS:
            for suite in ("full", "new"):
                self.write_trx(project, suite)

    def write_trx(self, project, suite, *, outcome="Passed", display=None, assembly=None,
                  empty=False, total=None, orphan=False, accounting=True):
        root = ET.Element("TestRun", xmlns="http://microsoft.com/schemas/VisualStudio/TeamTest/2010")
        results = ET.SubElement(root, "Results")
        definitions = ET.SubElement(root, "TestDefinitions")
        rows = []
        spec = new_test(project)
        if not empty:
            rows.append((spec["className"], spec["methodName"], display or spec["displayName"], outcome))
        if project == "Packaging" and suite == "full" and accounting:
            rows.append((ACCOUNTING, "Check", "Check", "Passed"))
        for index, (cls, method, name, state) in enumerate(rows):
            test = ET.SubElement(definitions, "UnitTest", id=str(index),
                                 storage=f"C:\\bin\\Ofdrw.Net.{assembly or project}.Tests.dll")
            ET.SubElement(test, "TestMethod", className=cls, name=method)
            ET.SubElement(results, "UnitTestResult", testId="unknown" if orphan else str(index),
                          testName=name, outcome=state)
        summary = ET.SubElement(root, "ResultSummary")
        ET.SubElement(summary, "Counters", total=str(len(rows) if total is None else total),
                      passed=str(sum(row[3] == "Passed" for row in rows)),
                      failed=str(sum(row[3] == "Failed" for row in rows)),
                      notExecuted=str(sum(row[3] == "NotExecuted" for row in rows)))
        path = self.root / project / suite / f"{suite}.trx"
        path.parent.mkdir(parents=True, exist_ok=True)
        ET.ElementTree(root).write(path, encoding="utf-8")

    def check(self, *args):
        return subprocess.run(
            [sys.executable, str(CHECKER), "--manifest", str(self.manifest),
             "--results-dir", str(self.root), *args], text=True, capture_output=True)

    def test_passing_gate(self):
        result = self.check("--pr", "A")
        self.assertEqual(0, result.returncode, result.stderr)
        self.assertIn("required class", result.stdout)
        self.assertIn("matched", result.stdout)

    def test_invalid_results(self):
        for changes, reason in (
            ({"empty": True}, "Empty or inconsistent result count"),
            ({"outcome": "Failed"}, "Failed tests"),
            ({"outcome": "NotExecuted"}, "non-passed"),
            ({"outcome": "Aborted"}, "Unexpected outcomes"),
            ({"display": 'OpensArchive(layout: "plain")'}, "missing"),
            ({"assembly": "Packaging"}, "Wrong test assembly"),
            ({"total": 3}, "inconsistent result count"),
            ({"orphan": True}, "unknown"),
        ):
            with self.subTest(changes=changes):
                self.write_trx("Cli", "new", **changes)
                result = self.check("--pr", "A")
                self.assertNotEqual(0, result.returncode)
                self.assertIn(reason, result.stderr)

    def test_required_full_results(self):
        self.write_trx("Packaging", "full", accounting=False)
        self.assertIn("required class", self.check("--pr", "A").stderr)
        self.write_trx("Packaging", "full")
        self.write_trx("Cli", "full", display="Other row")
        self.assertIn("Cli full: missing", self.check("--pr", "A").stderr)

    def test_missing_result_file(self):
        (self.root / "Cli/new/new.trx").unlink()
        result = self.check("--pr", "A")
        self.assertNotEqual(0, result.returncode)
        self.assertIn(str(self.root / "Cli/new/new.trx"), result.stderr)

    def test_batch_and_manifest_selection(self):
        for args in ((), ("--pr", "D"), ("--pr", "B")):
            with self.subTest(args=args):
                self.assertNotEqual(0, self.check(*args).returncode)
        self.manifest.write_text('{"prs":{"A":{"fullProjects":[],"newTests":[]}}}')
        result = self.check("--pr", "A")
        self.assertNotEqual(0, result.returncode)
        self.assertIn("Missing required full projects", result.stderr)


if __name__ == "__main__":
    unittest.main()
