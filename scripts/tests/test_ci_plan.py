import contextlib
import importlib.util
import io
import json
import os
from pathlib import Path
import re
import subprocess
import tempfile
import unittest
from unittest.mock import patch

SCRIPTS = Path(__file__).parents[1]
spec = importlib.util.spec_from_file_location('ci_plan', SCRIPTS / 'ci-plan.py')
plan = importlib.util.module_from_spec(spec)
spec.loader.exec_module(plan)

ROOT = plan.ROOT
REPO = 'owner/repo'
HEAD_TREE = 'a' * 40
BASE_TREE = 'b' * 40
BEFORE = 'c' * 40


class FakeGit:
    def __init__(self, changed=(), ancestor=True, known=(BEFORE,)):
        self.changed = list(changed)
        self.ancestor = ancestor
        self.known = set(known)

    def tree(self, rev):
        return HEAD_TREE if rev == 'HEAD' else BASE_TREE

    def commit(self, rev):
        if rev == 'HEAD^1':
            return BEFORE
        if rev not in self.known:
            raise subprocess.CalledProcessError(1, ['git', 'rev-parse', rev])
        return rev

    def is_ancestor(self, ancestor, descendant):
        return self.ancestor

    def changed_files(self, base, head):
        return self.changed


def markers(*valid, error=None):
    calls = []

    def lookup(tree):
        calls.append(tree)
        if error:
            raise plan.MarkerLookupError(error)
        return plan.Marker(tree in valid, 'stub')
    lookup.calls = calls
    return lookup


class ClassifierTests(unittest.TestCase):
    def test_documentation_is_skippable(self):
        for path in ('plans/026-emergency-line-breaking.md', 'docs/content/guides/flow-layout.md',
                     'docs/site/package.json', 'docs/validation/d10-imagesharp-security-assessment.md',
                     'docs/content/releases/0.3.1.md', 'AGENTS.md', 'CHANGELOG.md'):
            self.assertTrue(plan.is_skippable(path), path)

    def test_consumed_and_code_paths_need_ci(self):
        for path in ('README.md', 'THIRD-PARTY-NOTICES.md', 'docs/licenses/MUSL-COPYRIGHT.txt',
                     'docs/content/releases/0.5.1.md', 'docs/engineering/baselines/public-api-cs1591.json',
                     'docs/validation/public-tables-evidence/acceptance.json', 'docs/content/reference/feature-parity.md',
                     'src/Ofdrw.Net.Core/Ofdrw.Net.Core.csproj', 'scripts/ci-plan.py', '.github/workflows/ci.yml',
                     'Directory.Build.props', '.gitattributes', 'global.json', 'LICENSE', 'src/README.md',
                     'docsx/a.md', 'README.MD'):
            self.assertFalse(plan.is_skippable(path), path)


class DecisionTests(unittest.TestCase):
    def decide(self, event='push', changed=(), valid=(), before=BEFORE, error=None, **kwargs):
        lookup = markers(*valid, error=error)
        git = FakeGit(changed, **kwargs)
        return plan.decide(event, git, lookup, before=before), lookup.calls

    def test_docs_only_inherits_validated_base(self):
        decision, calls = self.decide(changed=['docs/content/guides/a.md', 'plans/1.md'], valid=[BASE_TREE])
        self.assertEqual((False, 'docs-only', HEAD_TREE, BASE_TREE),
                         (decision.run_full, decision.mode, decision.tree, decision.base_tree))
        self.assertEqual([HEAD_TREE, BASE_TREE], calls)

    def test_docs_only_on_unvalidated_base_runs_full(self):
        decision, _ = self.decide(changed=['docs/content/guides/a.md'])
        self.assertEqual((True, 'full'), (decision.run_full, decision.mode))
        self.assertIn('not validated', decision.reason)

    def test_exact_tree_is_reused_even_for_code_changes(self):
        decision, calls = self.decide(changed=['src/a.cs'], valid=[HEAD_TREE])
        self.assertEqual((False, 'reused', HEAD_TREE), (decision.run_full, decision.mode, decision.tree))
        self.assertEqual([HEAD_TREE], calls)

    def test_mixed_changes_run_full(self):
        decision, calls = self.decide(changed=['docs/content/guides/a.md', 'README.md'], valid=[BASE_TREE])
        self.assertEqual((True, 'full'), (decision.run_full, decision.mode))
        self.assertIn('README.md', decision.reason)
        self.assertEqual([HEAD_TREE], calls)

    def test_empty_diff_is_not_docs_only(self):
        decision, _ = self.decide(changed=[], valid=[BASE_TREE])
        self.assertEqual((True, 'full'), (decision.run_full, decision.mode))

    def test_unusable_push_base_runs_full(self):
        for before, kwargs in (('0' * 40, {}), ('', {}), ('d' * 40, {}), (BEFORE, {'ancestor': False})):
            decision, _ = self.decide(changed=['docs/a.md'], valid=[BASE_TREE], before=before, **kwargs)
            self.assertEqual((True, 'full'), (decision.run_full, decision.mode), before)

    def test_pull_request_diffs_against_first_parent(self):
        decision, _ = self.decide('pull_request', changed=['docs/a.md'], valid=[BASE_TREE], before='')
        self.assertEqual('docs-only', decision.mode)

    def test_api_failure_runs_full(self):
        decision, _ = self.decide(changed=['docs/a.md'], error='HTTPError: 502')
        self.assertEqual((True, 'full'), (decision.run_full, decision.mode))
        self.assertIn('502', decision.reason)

    def test_base_lookup_failure_runs_full(self):
        calls = []

        def lookup(tree):
            calls.append(tree)
            if tree == BASE_TREE:
                raise plan.MarkerLookupError('timed out')
            return plan.Marker(False, 'miss')
        decision = plan.decide('push', FakeGit(['docs/a.md']), lookup, before=BEFORE)
        self.assertEqual((True, 'full'), (decision.run_full, decision.mode))

    def test_forced_dispatch_skips_lookup(self):
        lookup = markers(HEAD_TREE)
        decision = plan.decide('workflow_dispatch', FakeGit(), lookup, force_full=True)
        self.assertEqual((True, 'full', []), (decision.run_full, decision.mode, lookup.calls))
        decision = plan.decide('workflow_dispatch', FakeGit(), lookup, force_full=False)
        self.assertEqual('reused', decision.mode)
        self.assertEqual('full', plan.decide('workflow_dispatch', FakeGit(['docs/a.md']), markers(BASE_TREE)).mode)


def artifact(identifier=1, run_id=10, expired=False, head_repo=7, repo=7, name=plan.marker_name(HEAD_TREE)):
    return {'id': identifier, 'name': name, 'expired': expired,
            'workflow_run': {'id': run_id, 'repository_id': repo, 'head_repository_id': head_repo}}


def run(path='.github/workflows/ci.yml', status='completed', conclusion='success'):
    return {'path': path, 'status': status, 'conclusion': conclusion}


class MarkerLookupTests(unittest.TestCase):
    def find(self, artifacts, runs=None):
        requests = []

        def fetch(path):
            requests.append(path)
            if '/actions/artifacts?' in path:
                return artifacts if isinstance(artifacts, dict) else {'total_count': len(artifacts), 'artifacts': artifacts}
            return (runs or {})[int(path.rsplit('/', 1)[1])]
        return plan.find_valid_marker(HEAD_TREE, fetch, REPO), requests

    def test_valid_marker(self):
        marker, requests = self.find([artifact()], {10: run()})
        self.assertTrue(marker.found)
        self.assertEqual((1, 10), (marker.artifact_id, marker.run_id))
        self.assertIn('name=ci-validated-' + HEAD_TREE, requests[0])
        self.assertEqual('/repos/owner/repo/actions/runs/10', requests[1])

    def test_rejections(self):
        cases = {
            'fork': ([artifact(head_repo=8)], {}),
            'expired': ([artifact(expired=True)], {}),
            'failed run': ([artifact()], {10: run(conclusion='failure')}),
            'running': ([artifact()], {10: run(status='in_progress', conclusion=None)}),
            'other workflow': ([artifact()], {10: run(path='.github/workflows/docs-pages.yml')}),
            'other name': ([artifact(name=plan.marker_name(BASE_TREE))], {}),
            'none': ([], {}),
        }
        for label, (artifacts, runs) in cases.items():
            marker, _ = self.find(artifacts, runs)
            self.assertFalse(marker.found, label)

    def test_later_valid_artifact_is_used(self):
        marker, _ = self.find([artifact(1, 10, head_repo=8), artifact(2, 11), artifact(3, 12)],
                              {11: run(conclusion='cancelled'), 12: run()})
        self.assertEqual((True, 3), (marker.found, marker.artifact_id))

    def test_errors_and_unexpected_payloads_raise(self):
        def broken(path):
            raise TimeoutError('timed out')
        with self.assertRaises(plan.MarkerLookupError):
            plan.find_valid_marker(HEAD_TREE, broken, REPO)
        for payload in ({'message': 'Bad credentials'}, {'artifacts': {'id': 1}}, ['not an object'], {'artifacts': [artifact(expired=None)]},
                        {'artifacts': [{'name': plan.marker_name(HEAD_TREE), 'expired': False}]}):
            with self.assertRaises(plan.MarkerLookupError):
                self.find(payload, {10: run()})
        with self.assertRaises(plan.MarkerLookupError):
            self.find([artifact()], {10: ['not', 'a', 'run']})


class GateTests(unittest.TestCase):
    def needs(self, run_full, **results):
        needs = {'changes': {'result': 'success', 'outputs': {'run_full': run_full}}}
        needs.update({name: {'result': result, 'outputs': {}} for name, result in results.items()})
        return needs

    def test_full_requires_success_everywhere(self):
        self.assertTrue(plan.gate(self.needs('true', a='success', b='success'))[0])
        for bad in ('failure', 'cancelled', 'skipped'):
            self.assertFalse(plan.gate(self.needs('true', a='success', b=bad))[0], bad)

    def test_skip_requires_skips_everywhere(self):
        self.assertTrue(plan.gate(self.needs('false', a='skipped', b='skipped'))[0])
        for bad in ('success', 'failure', 'cancelled'):
            self.assertFalse(plan.gate(self.needs('false', a='skipped', b=bad))[0], bad)

    def test_changes_must_succeed_with_a_plan(self):
        needs = self.needs('true', a='success')
        needs['changes']['result'] = 'failure'
        self.assertFalse(plan.gate(needs)[0])
        self.assertFalse(plan.gate(self.needs('', a='skipped'))[0])
        self.assertFalse(plan.gate(self.needs('true'))[0])
        self.assertFalse(plan.gate({'a': {'result': 'success'}})[0])
        self.assertFalse(plan.gate(None)[0])


class MarkerFileTests(unittest.TestCase):
    def write(self, **env):
        with tempfile.TemporaryDirectory() as temp, patch.dict(os.environ, env, clear=True):
            output = Path(temp) / 'marker' / 'marker.json'
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                code = plan.main(['write-marker', '--output', str(output)])
            return code, json.loads(output.read_text()) if output.exists() else None

    def test_marker_contents(self):
        code, marker = self.write(TREE=HEAD_TREE, MODE='docs-only', BASE_TREE=BASE_TREE, GITHUB_SHA=BEFORE,
                                  GITHUB_RUN_ID='5', GITHUB_RUN_ATTEMPT='2', GITHUB_EVENT_NAME='push')
        self.assertEqual(0, code)
        self.assertEqual({'schema': 1, 'tree': HEAD_TREE, 'commit': BEFORE, 'run_id': '5', 'run_attempt': '2',
                          'event': 'push', 'mode': 'docs-only', 'base_tree': BASE_TREE}, marker)

    def test_refuses_reused_or_malformed(self):
        self.assertEqual(1, self.write(TREE=HEAD_TREE, MODE='reused')[0])
        self.assertEqual(1, self.write(TREE='', MODE='full')[0])
        self.assertEqual(1, self.write(TREE=HEAD_TREE, MODE='docs-only')[0])


# Every workflow is scanned except docs-pages.yml, which builds the documentation
# site and is not part of the CI gate.
UNSCANNED_WORKFLOWS = ('.github/workflows/docs-pages.yml',)
BUILD_EXTENSIONS = ('.csproj', '.props', '.targets', '.sln', '.slnf', '.cs', '.json', '.py', '.sh', '.ps1',
                    '.yml', '.yaml', '.swift', '.xml', '.config')
SCRIPT_REFERENCE = re.compile(r'[A-Za-z0-9_./-]+\.(?:py|sh|ps1|swift)\b')
DOCS_REFERENCE = re.compile(r'(?<![A-Za-z0-9_.-])((?:docs|plans)/[A-Za-z0-9_.\-/]*)')
ROOT_MARKDOWN_REFERENCE = re.compile(r'(?<![A-Za-z0-9_.\-/])([A-Za-z0-9_.-]+\.md)\b')
# Joins Python path segments such as ROOT / 'docs' / 'engineering' and
# Path.Combine(root, "docs", "x") into one docs/engineering string.
SEGMENT_JOIN = re.compile(r"""['"]\s*[/,]\s*['"]""")


def tracked_files():
    # Skip deletions not yet committed, so the scan reflects the working tree.
    output = subprocess.run(['git', 'ls-files', '-z'], cwd=ROOT, check=True, capture_output=True).stdout
    return {p.decode() for p in output.split(b'\0') if p and (ROOT / p.decode()).is_file()}


def workflow_seeds(tracked):
    return sorted(p for p in tracked if p.startswith('.github/workflows/') and p.endswith(('.yml', '.yaml'))
                  and p not in UNSCANNED_WORKFLOWS)


def ci_invoked_files(tracked):
    """Workflows, composite actions, script tests, and every script they reach."""
    queue = workflow_seeds(tracked)
    queue += [p for p in tracked if p.startswith('.github/actions/')
              or (p.startswith('scripts/tests/') and p.endswith('.py'))]
    seen = set()
    while queue:
        path = queue.pop()
        if path in seen or path not in tracked:
            continue
        seen.add(path)
        for token in SCRIPT_REFERENCE.findall((ROOT / path).read_text(errors='replace')):
            parts = token.split('/')
            for start in range(len(parts)):
                suffix = '/'.join(parts[start:])
                for base in ('', 'scripts/', str(Path(path).parent) + '/'):
                    candidate = os.path.normpath(base + suffix)
                    if candidate in tracked and candidate not in seen:
                        queue.append(candidate)
    return seen


def build_inputs(tracked):
    files = ci_invoked_files(tracked)
    files |= {p for p in tracked if p.endswith(BUILD_EXTENSIONS)
              and ('/' not in p or p.split('/', 1)[0] in ('src', 'tests', 'e2e', 'native'))}
    # The planner and this test list the exception paths themselves.
    return files - {'scripts/ci-plan.py', 'scripts/tests/test_ci_plan.py'}


class ConsumedDocumentationGuardTests(unittest.TestCase):
    """Fail when a build or CI input reads a path that the planner would skip."""

    @classmethod
    def setUpClass(cls):
        cls.tracked = tracked_files()
        cls.inputs = build_inputs(cls.tracked)

    def referenced_paths(self):
        references = {}
        for path in sorted(self.inputs):
            text = SEGMENT_JOIN.sub('/', (ROOT / path).read_text(errors='replace'))
            for match in DOCS_REFERENCE.findall(text):
                reference = match.rstrip('/.')
                if '/' in reference:
                    references.setdefault(reference, set()).add(path)
            for match in ROOT_MARKDOWN_REFERENCE.findall(text):
                if match in self.tracked:
                    references.setdefault(match, set()).add(path)
        return references

    def test_scan_covers_every_gating_workflow(self):
        seeds = workflow_seeds(self.tracked)
        for path in ('.github/workflows/ci.yml', '.github/workflows/publish-nuget.yml',
                     '.github/workflows/nuget-advisories.yml'):
            self.assertIn(path, seeds)
        self.assertNotIn('.github/workflows/docs-pages.yml', seeds)
        self.assertLessEqual(set(seeds), self.inputs)

    def test_scan_reaches_known_consumers(self):
        for path in ('scripts/run-ses-interop-e2e.sh', 'scripts/verify-ses-interop-evidence.py',
                     'scripts/verify-public-table-evidence.py', 'scripts/verify-preview-record.py',
                     'scripts/check-public-api-docs.py', 'scripts/install-ci-fonts.py',
                     'scripts/verify-graphics-evidence.py', 'scripts/validation/text-mixed/check-export-batch.py',
                     'Directory.Build.props', 'src/Ofdrw.Net.Crypto.Password/Ofdrw.Net.Crypto.Password.csproj',
                     'tests/Ofdrw.Net.Packaging.Tests/ArchivalDossierTests.cs'):
            self.assertIn(path, self.inputs)
        references = self.referenced_paths()
        for path in ('docs/content/reference/conversion-contracts.md', 'docs/engineering/baselines/public-api-cs1591.json',
                     'docs/validation/public-tables-evidence', 'README.md', 'THIRD-PARTY-NOTICES.md'):
            self.assertIn(path, references)

    def test_referenced_documentation_is_never_skippable(self):
        skippable = []
        for reference, consumers in sorted(self.referenced_paths().items()):
            covered = [reference] if reference in self.tracked else \
                [p for p in self.tracked if p.startswith(reference + '/')]
            for path in covered:
                if plan.is_skippable(path):
                    skippable.append(f'{path} (via {reference} in {", ".join(sorted(consumers))})')
        self.assertEqual([], skippable, 'Add these consumed paths to CONSUMED_DOCS in scripts/ci-plan.py')


if __name__ == '__main__':
    unittest.main()
