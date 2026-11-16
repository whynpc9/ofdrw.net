#!/usr/bin/env python3
"""Decide whether CI must run its heavy jobs for the checked-out tree.

A validation marker is an artifact named ``ci-validated-<git tree SHA>``. It
means: this exact tree passed the full CI suite, or differs only by skippable
(documentation) changes from a tree that did.

Decision order (``decide``):
  1. ``workflow_dispatch`` with ``force_full`` -> full.
  2. A valid marker for ``HEAD^{tree}`` -> skip heavy jobs (``reused``).
  3. Every changed file is skippable and the base tree has a valid marker ->
     skip heavy jobs (``docs-only``); the run uploads a marker for its tree.
  4. Otherwise -> full.

Every API, network, payload or git problem falls back to a full run.

Subcommands:
  decide        compute the plan; writes GITHUB_OUTPUT and GITHUB_STEP_SUMMARY
  gate          aggregate needs results for the "CI complete" job
  write-marker  write the marker JSON uploaded by "CI complete"
"""

import argparse
from dataclasses import dataclass, field
import json
import os
from pathlib import Path
import re
import subprocess
import sys
import urllib.error
import urllib.parse
import urllib.request

ROOT = Path(__file__).resolve().parents[1]
MARKER_PREFIX = 'ci-validated-'
CI_WORKFLOW_PATH = '.github/workflows/ci.yml'
ZERO_SHA = re.compile(r'^0+$')
OBJECT_ID = re.compile(r'^[0-9a-f]{40}([0-9a-f]{24})?$')
MAX_MARKER_CANDIDATES = 10

# Files under these prefixes (and root-level *.md) are documentation for CI.
SKIPPABLE_PREFIXES = ('docs/', 'plans/')

# ...except these, which the build, tests, packaging or CI scripts consume.
# Entries ending in '/' cover a whole directory. scripts/tests/test_ci_plan.py
# scans csproj/props/targets, source/test code and every script that CI or the
# release workflow invokes, and fails when a referenced docs path is skippable.
CONSUMED_DOCS = (
    # Packed into every package by Directory.Build.props.
    'docs/licenses/',
    'docs/content/reference/conversion-contracts.md',
    'docs/content/reference/feature-parity.md',
    'docs/content/releases/0.4.0.md',
    # Packed by individual projects.
    'docs/content/reference/password-crypto.md',
    'docs/content/guides/pdf-vector.md',
    'docs/content/guides/ses-interop.md',
    'docs/architecture/pdf-vector-design-contract.md',
    'docs/evidence/ses-interop/BouncyCastle-LICENSE.md',
    # Read by ArchivalDossierTests.
    'docs/content/reference/archival-standard-coverage.md',
    # API-documentation, NuGet-advisory and third-party dependency baselines.
    'docs/engineering/baselines/',
    # verify-public-table-evidence.py.
    'docs/validation/public-tables-evidence/',
    # verify-preview-record.py (release gate and its unit tests).
    'docs/validation/releases/',
)

# Root-level Markdown that packaging consumes (PackageReadmeFile and the
# notices that verify-package-artifacts.py compares byte-for-byte).
CONSUMED_ROOT_MARKDOWN = ('README.md', 'THIRD-PARTY-NOTICES.md')


def is_skippable(path):
    """True when a change to ``path`` cannot affect any CI job's result."""
    if '/' not in path:
        return path.endswith('.md') and path not in CONSUMED_ROOT_MARKDOWN
    if not path.startswith(SKIPPABLE_PREFIXES):
        return False
    for consumed in CONSUMED_DOCS:
        if path == consumed or (consumed.endswith('/') and path.startswith(consumed)):
            return False
    return True


def marker_name(tree):
    return MARKER_PREFIX + tree


class MarkerLookupError(Exception):
    """The marker store could not be queried reliably; never skip on this."""


@dataclass
class Marker:
    found: bool
    detail: str
    artifact_id: int = 0
    run_id: int = 0


@dataclass
class Decision:
    run_full: bool
    mode: str
    tree: str
    reason: str
    base_tree: str = ''
    changed: list = field(default_factory=list)


def require(condition, message):
    if not condition:
        raise MarkerLookupError(message)


def find_valid_marker(tree, fetch_json, repository):
    """Return the first valid marker for ``tree``.

    ``fetch_json(path)`` performs a GET against the REST API and returns the
    decoded JSON; any exception it raises is reported as a lookup error.
    """
    name = marker_name(tree)
    try:
        listing = fetch_json(f'/repos/{repository}/actions/artifacts?'
                             + urllib.parse.urlencode({'name': name, 'per_page': 100}))
        require(isinstance(listing, dict) and isinstance(listing.get('artifacts'), list),
                'artifact listing has no artifacts array')
        rejected = []
        checked = 0
        for artifact in listing['artifacts']:
            require(isinstance(artifact, dict), 'artifact entry is not an object')
            if artifact.get('name') != name:
                continue
            identifier = artifact.get('id')
            expired = artifact.get('expired')
            require(isinstance(expired, bool), f'artifact {identifier} has no boolean expired flag')
            if expired:
                rejected.append(f'{identifier}: expired')
                continue
            run_ref = artifact.get('workflow_run')
            require(isinstance(run_ref, dict), f'artifact {identifier} has no workflow_run')
            run_id = run_ref.get('id')
            repo_id = run_ref.get('repository_id')
            head_repo_id = run_ref.get('head_repository_id')
            require(isinstance(run_id, int) and isinstance(repo_id, int),
                    f'artifact {identifier} has no workflow run or repository id')
            if head_repo_id != repo_id:
                rejected.append(f'{identifier}: produced by a fork run ({head_repo_id} != {repo_id})')
                continue
            if checked >= MAX_MARKER_CANDIDATES:
                break
            checked += 1
            run = fetch_json(f'/repos/{repository}/actions/runs/{run_id}')
            require(isinstance(run, dict), f'workflow run {run_id} is not an object')
            path = run.get('path')
            require(isinstance(path, str), f'workflow run {run_id} has no path')
            if path.split('@', 1)[0] != CI_WORKFLOW_PATH:
                rejected.append(f'{identifier}: produced by {path}')
                continue
            if run.get('status') != 'completed' or run.get('conclusion') != 'success':
                rejected.append(f"{identifier}: run {run_id} is {run.get('status')}/{run.get('conclusion')}")
                continue
            return Marker(True, f'artifact {identifier} from successful CI run {run_id}', identifier, run_id)
        return Marker(False, f'no valid {name}' + (f" (rejected: {'; '.join(rejected)})" if rejected else ''))
    except MarkerLookupError:
        raise
    except Exception as error:  # Network, HTTP, JSON or anything unexpected.
        raise MarkerLookupError(f'{type(error).__name__}: {error}') from error


class GitHubApi:
    def __init__(self, api_url, token, timeout=20):
        self.api_url = api_url.rstrip('/')
        self.token = token
        self.timeout = timeout

    def __call__(self, path):
        request = urllib.request.Request(self.api_url + path, headers={
            'Accept': 'application/vnd.github+json',
            'Authorization': f'Bearer {self.token}',
            'X-GitHub-Api-Version': '2022-11-28',
            'User-Agent': 'ci-plan',
        })
        with urllib.request.urlopen(request, timeout=self.timeout) as response:
            return json.loads(response.read().decode('utf-8'))


class Git:
    def __init__(self, cwd=ROOT):
        self.cwd = cwd

    def run(self, *args):
        return subprocess.run(['git', *args], cwd=self.cwd, check=True, capture_output=True).stdout

    def tree(self, rev):
        return self.run('rev-parse', '--verify', f'{rev}^{{tree}}').decode().strip()

    def commit(self, rev):
        return self.run('rev-parse', '--verify', '--quiet', f'{rev}^{{commit}}').decode().strip()

    def is_ancestor(self, ancestor, descendant):
        return subprocess.run(['git', 'merge-base', '--is-ancestor', ancestor, descendant],
                              cwd=self.cwd, capture_output=True).returncode == 0

    def changed_files(self, base, head):
        raw = self.run('diff', '--name-only', '--no-renames', '-z', base, head)
        return [p.decode('utf-8', 'surrogateescape') for p in raw.split(b'\0') if p]


def lookup(lookup_marker, tree):
    """(Marker or None, error text)."""
    try:
        return lookup_marker(tree), ''
    except MarkerLookupError as error:
        return None, str(error)


def resolve_base(event, git, before, head='HEAD'):
    """Return (base commit, '') or ('', why the diff base is unusable)."""
    try:
        if event == 'pull_request':
            return git.commit(f'{head}^1'), ''
        if event == 'push':
            if not before or ZERO_SHA.match(before):
                return '', 'push has no previous commit (new branch or missing before)'
            if not OBJECT_ID.match(before):
                return '', f'push before {before!r} is not an object id'
            base = git.commit(before)
            if not git.is_ancestor(base, head):
                return '', f'push before {before[:12]} is not an ancestor of HEAD (force push?)'
            return base, ''
        return '', f'event {event} has no diff base'
    except (subprocess.CalledProcessError, OSError) as error:
        return '', f'cannot resolve diff base: {error}'


def decide(event, git, lookup_marker, before='', force_full=False, head='HEAD'):
    tree = git.tree(head)
    if event == 'workflow_dispatch' and force_full:
        return Decision(True, 'full', tree, 'manual dispatch with force_full')

    marker, error = lookup(lookup_marker, tree)
    if error:
        return Decision(True, 'full', tree, f'marker lookup for tree {tree[:12]} failed, running full: {error}')
    if marker.found:
        return Decision(False, 'reused', tree, f'tree {tree[:12]} already validated: {marker.detail}')

    base, why = resolve_base(event, git, before, head)
    if not base:
        return Decision(True, 'full', tree, f'{why}; {marker.detail}')
    try:
        changed = git.changed_files(base, head)
        base_tree = git.tree(base)
    except (subprocess.CalledProcessError, OSError) as error:
        return Decision(True, 'full', tree, f'cannot diff against {base[:12]}: {error}')
    if not changed:
        return Decision(True, 'full', tree, f'no changed files against {base[:12]} and {marker.detail}',
                        base_tree, changed)
    code = [p for p in changed if not is_skippable(p)]
    if code:
        shown = ', '.join(code[:5]) + (f' and {len(code) - 5} more' if len(code) > 5 else '')
        return Decision(True, 'full', tree, f'{len(code)} of {len(changed)} changed files need CI: {shown}',
                        base_tree, changed)
    base_marker, error = lookup(lookup_marker, base_tree)
    if error:
        return Decision(True, 'full', tree, f'docs-only change but base marker lookup failed: {error}',
                        base_tree, changed)
    if not base_marker.found:
        return Decision(True, 'full', tree, f'docs-only change but base tree {base_tree[:12]} is not validated: '
                        f'{base_marker.detail}', base_tree, changed)
    return Decision(False, 'docs-only', tree, f'{len(changed)} docs-only files on validated base tree '
                    f'{base_tree[:12]}: {base_marker.detail}', base_tree, changed)


def write_outputs(decision, output_path, summary_path):
    reason = ' '.join(decision.reason.split())
    if output_path:
        with open(output_path, 'a', encoding='utf-8') as output:
            output.write(f"run_full={'true' if decision.run_full else 'false'}\n"
                         f'mode={decision.mode}\ntree={decision.tree}\n'
                         f'base_tree={decision.base_tree}\nreason={reason}\n')
    if summary_path:
        lines = ['## CI plan', '', f'- Mode: `{decision.mode}`',
                 f"- Heavy jobs: {'run' if decision.run_full else 'skipped'}",
                 f'- Tree: `{decision.tree}`']
        if decision.base_tree:
            lines.append(f'- Base tree: `{decision.base_tree}`')
        lines += [f'- Reason: {reason}', f'- Changed files: {len(decision.changed)}']
        if decision.changed:
            lines += ['', '<details><summary>Changed files</summary>', '']
            lines += [f"- `{p}`{'' if is_skippable(p) else ' (needs CI)'}" for p in decision.changed[:200]]
            lines += ['', '</details>']
        with open(summary_path, 'a', encoding='utf-8') as summary:
            summary.write('\n'.join(lines) + '\n')


def event_before(event_path):
    if not event_path:
        return ''
    try:
        return str(json.loads(Path(event_path).read_text(encoding='utf-8')).get('before') or '')
    except (OSError, ValueError, AttributeError):
        return ''


def command_decide(args):
    event = args.event or os.environ.get('GITHUB_EVENT_NAME', '')
    before = args.before if args.before is not None else event_before(os.environ.get('GITHUB_EVENT_PATH'))
    force_full = (args.force_full or os.environ.get('FORCE_FULL', '')).strip().lower() == 'true'
    if args.offline:
        valid = set(args.valid_tree)

        def lookup_marker(tree):
            return Marker(tree in valid, f"offline stub {'hit' if tree in valid else 'miss'}")
    else:
        repository = os.environ.get('GITHUB_REPOSITORY', '')
        token = os.environ.get('GITHUB_TOKEN', '')
        api = GitHubApi(os.environ.get('GITHUB_API_URL', 'https://api.github.com'), token)

        def lookup_marker(tree):
            if not repository or not token:
                raise MarkerLookupError('GITHUB_REPOSITORY or GITHUB_TOKEN is not set')
            return find_valid_marker(tree, api, repository)

    decision = decide(event, Git(), lookup_marker, before, force_full, args.head)
    write_outputs(decision, os.environ.get('GITHUB_OUTPUT'), os.environ.get('GITHUB_STEP_SUMMARY'))
    print(f'mode={decision.mode} run_full={str(decision.run_full).lower()} tree={decision.tree}')
    print(f'reason: {decision.reason}')
    for path in decision.changed:
        print(f"  {'docs' if is_skippable(path) else 'code'}  {path}")
    return 0


def gate(needs):
    """Return (ok, message) for the aggregate "CI complete" check."""
    if not isinstance(needs, dict) or not isinstance(needs.get('changes'), dict):
        return False, 'needs has no changes job'
    changes = needs['changes']
    if changes.get('result') != 'success':
        return False, f"changes job result is {changes.get('result')}"
    run_full = (changes.get('outputs') or {}).get('run_full')
    if run_full not in ('true', 'false'):
        return False, f'changes job produced run_full={run_full!r}'
    others = {name: job.get('result') if isinstance(job, dict) else None
              for name, job in needs.items() if name != 'changes'}
    if not others:
        return False, 'no CI jobs besides changes'
    expected = 'success' if run_full == 'true' else 'skipped'
    wrong = {name: result for name, result in others.items() if result != expected}
    if wrong:
        return False, f'run_full={run_full} requires every CI job to be {expected}: {wrong}'
    if run_full == 'true':
        return True, f'All {len(others)} source, platform, package and crypto jobs passed.'
    return True, f'Heavy jobs skipped by plan; all {len(others)} CI jobs were skipped as expected.'


def command_gate(args):
    try:
        needs = json.loads(os.environ.get('NEEDS', ''))
    except ValueError as error:
        needs = None
        print(f'NEEDS is not JSON: {error}', file=sys.stderr)
    ok, message = gate(needs)
    print(message, file=sys.stdout if ok else sys.stderr)
    return 0 if ok else 1


def command_write_marker(args):
    tree, mode = os.environ.get('TREE', ''), os.environ.get('MODE', '')
    if not OBJECT_ID.match(tree) or mode not in ('full', 'docs-only'):
        print(f'Refusing to write a marker for tree={tree!r} mode={mode!r}', file=sys.stderr)
        return 1
    marker = {'schema': 1, 'tree': tree, 'commit': os.environ.get('GITHUB_SHA', ''),
              'run_id': os.environ.get('GITHUB_RUN_ID', ''), 'run_attempt': os.environ.get('GITHUB_RUN_ATTEMPT', ''),
              'event': os.environ.get('GITHUB_EVENT_NAME', ''), 'mode': mode}
    if mode == 'docs-only':
        base_tree = os.environ.get('BASE_TREE', '')
        if not OBJECT_ID.match(base_tree):
            print(f'docs-only marker needs a base tree, got {base_tree!r}', file=sys.stderr)
            return 1
        marker['base_tree'] = base_tree
    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    Path(args.output).write_text(json.dumps(marker, indent=2) + '\n', encoding='utf-8')
    print(json.dumps(marker))
    return 0


def main(argv=None):
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest='command', required=True)
    decide_parser = commands.add_parser('decide')
    decide_parser.add_argument('--event', help='defaults to GITHUB_EVENT_NAME')
    decide_parser.add_argument('--before', help='push base; defaults to the event payload')
    decide_parser.add_argument('--force-full', help="'true' forces a full run; defaults to FORCE_FULL")
    decide_parser.add_argument('--head', default='HEAD', help='commit to plan for (local dry runs)')
    decide_parser.add_argument('--offline', action='store_true',
                               help='local dry run: no API calls; only --valid-tree trees have markers')
    decide_parser.add_argument('--valid-tree', action='append', default=[], metavar='TREE')
    decide_parser.set_defaults(handler=command_decide)
    commands.add_parser('gate').set_defaults(handler=command_gate)
    marker_parser = commands.add_parser('write-marker')
    marker_parser.add_argument('--output', required=True)
    marker_parser.set_defaults(handler=command_write_marker)
    args = parser.parse_args(argv)
    return args.handler(args)


if __name__ == '__main__':
    sys.exit(main())
