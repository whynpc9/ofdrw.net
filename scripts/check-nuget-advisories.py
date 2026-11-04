#!/usr/bin/env python3
"""Fail on NuGet advisories that are not in the reviewed allowlist.

Runs `dotnet list Ofdrw.Net.sln package --vulnerable --include-transitive --format json`
(after a restore), or reads a saved report with --report. Each allowlist entry must cite a
recorded decision; stale entries are reported as warnings so fixed advisories get removed.
Exit codes: 0 clean, 1 unreviewed advisory, 2 the report could not be produced or parsed.
"""
import argparse
import json
import subprocess
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
ALLOWLIST = ROOT / 'docs' / 'engineering' / 'baselines' / 'nuget-advisory-allowlist.json'


class ReportError(Exception):
    pass


def findings(report):
    """Return {(package, version, advisory id): severity} for every project and framework."""
    if not isinstance(report, dict) or not isinstance(report.get('projects'), list):
        raise ReportError('the vulnerability report has no projects list')
    result = {}
    for project in report['projects']:
        for framework in project.get('frameworks', []) or []:
            for kind in ('topLevelPackages', 'transitivePackages'):
                for package in framework.get(kind, []) or []:
                    for advisory in package.get('vulnerabilities', []) or []:
                        identifier = advisory['advisoryurl'].rstrip('/').rsplit('/', 1)[-1]
                        key = (package['id'], package['resolvedVersion'], identifier)
                        result[key] = advisory.get('severity', 'unknown')
    return result


def allowlisted(path=ALLOWLIST):
    data = json.loads(Path(path).read_text(encoding='utf-8'))
    if data.get('schema') != 1:
        raise ReportError(f'unsupported allowlist schema in {path}')
    return {(entry['package'], entry['version'], entry['id']): entry for entry in data['advisories']}


def evaluate(report, allowed):
    found = findings(report)
    unreviewed = sorted(key for key in found if key not in allowed)
    stale = sorted(key for key in allowed if key not in found)
    for package, version, identifier in unreviewed:
        print(f'Unreviewed advisory: {package} {version} {identifier} ({found[(package, version, identifier)]})')
    for package, version, identifier in stale:
        print(f'Warning: stale allowlist entry {package} {version} {identifier} no longer occurs.')
    print(f'NuGet advisory gate: {len(found)} advisories, {len(unreviewed)} unreviewed, {len(stale)} stale allowlist entries.')
    return 1 if unreviewed else 0


def load_report(path):
    if path is not None:
        text = Path(path).read_text(encoding='utf-8')
    else:
        completed = subprocess.run(
            ['dotnet', 'list', 'Ofdrw.Net.sln', 'package', '--vulnerable', '--include-transitive', '--format', 'json'],
            cwd=ROOT, capture_output=True, text=True)
        if completed.returncode != 0:
            raise ReportError('dotnet list failed: ' + (completed.stderr or completed.stdout)[-2000:])
        text = completed.stdout
    try:
        return json.loads(text)
    except json.JSONDecodeError as error:
        raise ReportError(f'the vulnerability report is not JSON: {error}') from error


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    parser.add_argument('--report', help='read a saved dotnet list --format json report')
    parser.add_argument('--allowlist', default=str(ALLOWLIST))
    args = parser.parse_args(argv)
    try:
        return evaluate(load_report(args.report), allowlisted(args.allowlist))
    except (ReportError, KeyError, OSError, ValueError) as error:
        print(f'NuGet advisory gate could not run: {error}', file=sys.stderr)
        return 2


if __name__ == '__main__':
    sys.exit(main())
