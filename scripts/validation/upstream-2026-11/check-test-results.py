#!/usr/bin/env python3
"""Check Plan 019's per-PR test identities and outcomes in fresh VSTest TRX files."""

import argparse
from collections import Counter
import json
from pathlib import Path
import sys
import xml.etree.ElementTree as ET

NS = {"trx": "http://microsoft.com/schemas/VisualStudio/TeamTest/2010"}
FULL_MINIMUM = {
    "A": {"Packaging", "Cli"},
    "B": {"Converter.Pdf"},
    "C": {"Converter.Pdf"},
}
REQUIRED_CLASSES = {
    "A": {"Packaging": {"Ofdrw.Net.Packaging.Tests.ControlledReaderAccountingTests"}},
    "B": {"Converter.Pdf": {"Ofdrw.Net.Converter.Pdf.Tests.SvgGeometryTests"}},
    "C": {},
}


def require(condition, message):
    if not condition:
        raise ValueError(message)


def read_trx(path, project):
    root = ET.parse(path).getroot()
    summary = root.find("trx:ResultSummary/trx:Counters", NS)
    require(summary is not None, "Missing ResultSummary/Counters")
    counters = {key: int(summary.attrib[key]) for key in ("total", "passed", "failed", "notExecuted")}
    definitions = {}
    assembly = f"Ofdrw.Net.{project}.Tests.dll".lower()
    for test in root.findall("trx:TestDefinitions/trx:UnitTest", NS):
        method = test.find("trx:TestMethod", NS)
        require(method is not None, "Missing TestMethod")
        storage = test.attrib["storage"].replace("\\", "/").rsplit("/", 1)[-1].lower()
        require(storage == assembly, f"Wrong test assembly {storage}, expected {assembly}")
        test_id = test.attrib["id"]
        require(test_id not in definitions, f"Duplicate TestDefinition {test_id}")
        definitions[test_id] = (method.attrib["className"], method.attrib["name"])

    outcomes = []
    for result in root.findall("trx:Results/trx:UnitTestResult", NS):
        class_name, method_name = definitions[result.attrib["testId"]]
        outcomes.append(((class_name, method_name, result.attrib["testName"]), result.attrib["outcome"]))
    counts = Counter(outcome for _, outcome in outcomes)
    require(counters["total"] > 0 and counters["total"] == len(outcomes), "Empty or inconsistent result count")
    for outcome, field in (("Passed", "passed"), ("Failed", "failed"), ("NotExecuted", "notExecuted")):
        require(counts[outcome] == counters[field], f"Mismatch for {field}")
    require(sum(counts[key] for key in ("Passed", "Failed", "NotExecuted")) == len(outcomes),
            f"Unexpected outcomes {dict(counts)}")
    require(counters["failed"] == 0, f"Failed tests {[(key, state) for key, state in outcomes if state == 'Failed']}")
    print(f"{path}: {counters}")
    return outcomes


def check(manifest, results_dir, batch):
    spec = manifest["prs"][batch]
    full_projects = spec["fullProjects"]
    require(len(full_projects) == len(set(full_projects)), "Duplicate full projects")
    require(FULL_MINIMUM[batch] <= set(full_projects), f"Missing required full projects for PR {batch}")
    require(all(p and all(part.isascii() and part.isalnum() for part in p.split(".")) for p in full_projects),
            "Invalid project name")
    by_project = {}
    for test in spec["newTests"]:
        project = test["project"]
        require(project in full_projects, f"Missing full project for new test in {project}")
        require(test["item"] in {"M01", "M02", "M04", "M05", "M06", "M10"}, f"Invalid Plan item {test}")
        key = tuple(test[k] for k in ("className", "methodName", "displayName"))
        require(all(isinstance(value, str) and value for value in key), f"Invalid test identity {test}")
        tests = by_project.setdefault(project, set())
        require(key not in tests, f"Duplicate expected test {key}")
        tests.add(key)
    require(FULL_MINIMUM[batch] <= by_project.keys(), f"Missing required new-test projects for PR {batch}")

    for project in full_projects:
        full = read_trx(results_dir / project / "full/full.trx", project)
        for class_name in REQUIRED_CLASSES[batch].get(project, set()):
            found = [state for key, state in full if key[0] == class_name]
            require(found and all(state == "Passed" for state in found), f"Missing or skipped required class {class_name}")
            print(f"  required class {class_name}: {len(found)} passed")
        for key, state in full:
            if state == "NotExecuted":
                print(f"  skipped (requires documented review): {key}")
        expected = by_project.get(project)
        if not expected:
            continue
        new = read_trx(results_dir / project / "new/new.trx", project)
        for label, results in (("full", full), ("new", new)):
            passed = Counter(key for key, state in results if state == "Passed")
            missing = sorted(key for key in expected if passed[key] != 1)
            require(not missing, f"{project} {label}: missing, non-passed or duplicated tests {missing}")
        require(len(new) == len(expected) and all(state == "Passed" for _, state in new),
                f"{project}: filtered run includes unexpected or non-passed results")
        for key in sorted(expected):
            print(f"  matched {key}")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--manifest", required=True, type=Path)
    parser.add_argument("--results-dir", required=True, type=Path)
    parser.add_argument("--pr", required=True, choices=sorted(FULL_MINIMUM))
    args = parser.parse_args()
    try:
        check(json.loads(args.manifest.read_text(encoding="utf-8")), args.results_dir, args.pr)
    except (OSError, ValueError, ET.ParseError, KeyError, TypeError, AttributeError) as error:
        print(f"PR {args.pr} TRX gate failed: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
