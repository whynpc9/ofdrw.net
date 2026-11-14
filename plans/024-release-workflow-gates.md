# Plan 024: A release tag cannot publish unless CI and SES passed on that commit

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- .github/workflows/publish-nuget.yml docs/engineering/ci.md docs/engineering/release-gates.md`
> If any in-scope file changed since this plan was written, compare the
> "Current state" excerpts against the live code before proceeding; on a
> mismatch, treat it as a STOP condition.

## Status

- **Priority**: P1
- **Effort**: S–M
- **Risk**: MED (a workflow that only runs at release time; a mistake blocks the next release)
- **Depends on**: none
- **Category**: security (release pipeline)
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

`.github/workflows/publish-nuget.yml` publishes twelve packages to nuget.org when
a `v*` tag is pushed. Publishing cannot be undone. Today the workflow:

1. reruns only Linux checks. It never confirms that the three-platform `CI`
   workflow or the `SES dev interop` workflow succeeded for the tagged commit,
   although `docs/engineering/ci.md` calls SES "required for task acceptance";
2. skips the NuGet advisory gate (`scripts/check-nuget-advisories.py`) and the
   script unit tests that ordinary CI runs;
3. grants `id-token: write` (the OIDC permission that lets a step obtain a
   nuget.org publishing key) to one big job that also restores packages, runs
   `pip install`, runs LibreOffice and the whole test suite. Any of those steps
   could request the publishing token;
4. has no `concurrency` group, so a tag push and a manual recovery run for the
   same tag can race.

After this plan: the release job fails early unless the tagged commit has
successful `CI complete` and `ses-interop` check runs; it runs the advisory gate
and script tests; and the OIDC token exists only in a small second job that
downloads the already-validated packages and pushes them.

## Current state

File: `.github/workflows/publish-nuget.yml` (182 lines, one job `publish`).

Header and permissions (lines 19–29):

```yaml
jobs:
  publish:
    runs-on: ubuntu-latest
    env:
      DOTNET_CLI_HOME: ${{ github.workspace }}/artifacts/dotnet-home
      NUGET_PACKAGES: ${{ github.workspace }}/artifacts/nuget-cache
      DOTNET_SKIP_FIRST_TIME_EXPERIENCE: "1"
      DOTNET_CLI_TELEMETRY_OPTOUT: "1"
    permissions:
      contents: read
      id-token: write
```

Step order (step names, with line numbers): `Resolve existing release tag` (32,
`id: release`; sets outputs `tag` and `commit`, and env `PACKAGE_VERSION`),
`Checkout` (73), `Record release run provenance` (79),
`Fetch test fixtures from LFS` (95), `Setup .NET` (99), nuget-cache (103),
`Install document conversion dependencies` (105), setup-python (110),
`Install deterministic CJK font` (114), `Restore` (117),
`Verify release license and resolved dependency inventory` (120),
`Build and reject new undocumented public API` (123),
`Run mandatory original-font geometry and replacement tests without skips` (126),
`Test` (129), `Pack SDK packages` (132, writes `artifacts/nuget`),
`Consume the exact release artifacts` (147, writes
`artifacts/nuget/package-manifest.json`),
`Upload package manifest and validation pages` (150), `List artifacts` (160),
`Require candidate-bound macOS Preview acceptance` (163),
`NuGet login (OIDC)` (166), `Publish packages and symbols` (172).

The last two steps (lines 166–182):

```yaml
      - name: NuGet login (OIDC)
        id: nuget-login
        uses: NuGet/login@8d196754b4036150537f80ac539e15c2f1028841 # v1
        with:
          user: whynpc

      - name: Publish packages and symbols
        shell: bash
        run: |
          python3 scripts/verify-package-artifacts.py artifacts/nuget "${{ env.PACKAGE_VERSION }}" --verify-manifest
          shopt -s nullglob
          for package in artifacts/nuget/*.nupkg; do
            if [[ "$package" == *.snupkg ]]; then
              continue
            fi
            dotnet nuget push "$package" --source "https://api.nuget.org/v3/index.json" --api-key "${{ steps.nuget-login.outputs.NUGET_API_KEY }}" --skip-duplicate
          done
```

Facts you need:

- The check runs on a released commit are named exactly `CI complete` (job
  `ci-complete` in `.github/workflows/ci.yml:570-571`) and `ses-interop` (job id
  in `.github/workflows/ses-interop.yml:17`). Verified on the `v0.3.1` commit
  `f16ba8e`: both exist with conclusion `success`.
- `CI` and `SES dev interop` run on pushes to `main`. Release tags point at
  commits on `main` (for example `v0.3.1` → the merge commit `f16ba8e`), so
  those check runs exist for a correctly prepared release.
- `scripts/verify-package-artifacts.py ... --verify-manifest` needs a restored
  solution (it re-validates the dependency license inventory), so it must stay
  in the job that ran `dotnet restore`.
- `dotnet nuget push x.nupkg` also pushes the matching `x.snupkg` beside it.
- CI uploads packages with `compression-level: 0` (`ci.yml:133-139`); follow that.
- Trusted Publishing on nuget.org is bound to the repository and the workflow
  file name. Splitting the file into two jobs keeps the same workflow file. Do
  NOT add an `environment:` to any job in this plan: a GitHub environment must
  be created by the maintainer and may need a matching nuget.org policy change.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| YAML parses | `python3 -c "import yaml,sys; yaml.safe_load(open('.github/workflows/publish-nuget.yml')); print('ok')"` | `ok` |
| Lint (if installed) | `actionlint .github/workflows/publish-nuget.yml` | no output. If `actionlint` is not installed, say so in the report. |
| Script tests | `python3 -m unittest discover -s scripts/tests -v` | `OK` (run after `dotnet restore Ofdrw.Net.sln`) |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `.github/workflows/publish-nuget.yml`
- `docs/engineering/ci.md` (add a short section)
- `docs/engineering/release-gates.md` (one paragraph in the release-order list)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- `.github/workflows/ci.yml`, `ses-interop.yml`, `docs-pages.yml` — plan 025.
- The existing validation steps and their order, the tag-resolution script, the
  Preview acceptance check, the `NuGet/login` pin and `user:` value.
- Creating GitHub environments or changing anything on nuget.org.
- Any historical acceptance JSON under `docs/validation/`.

## Git workflow

- Branch: `codex/release-workflow-gates`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push, tag, or trigger any workflow. Never create a `v*` tag.

## Steps

### Step 1: Add a concurrency group

At workflow top level (after the `env:` block) add:

```yaml
concurrency:
  group: publish-nuget-${{ github.event_name == 'workflow_dispatch' && inputs.release_tag || github.ref_name }}
  cancel-in-progress: false
```

**Verify**: YAML parse command → `ok`.

### Step 2: Rename the job to `validate` and remove its OIDC permission

Rename job id `publish` to `validate`. Change its permissions to:

```yaml
    permissions:
      contents: read
      checks: read
```

Add job outputs so the next job can read the version and tag:

```yaml
    outputs:
      package_version: ${{ steps.release.outputs.package_version }}
      tag: ${{ steps.release.outputs.tag }}
```

In the `Resolve existing release tag` script add, next to the existing
`echo "tag=..."` line: `echo "package_version=$PACKAGE_VERSION" >> "$GITHUB_OUTPUT"`.

**Verify**: `grep -n "id-token" .github/workflows/publish-nuget.yml` → no match
yet inside the `validate` job (it reappears in Step 5 only).

### Step 3: Require CI and SES success for the release commit

Insert a step directly after `Record release run provenance`:

```yaml
      - name: Require CI and SES success on the release commit
        shell: bash
        env:
          GH_TOKEN: ${{ github.token }}
          RELEASE_COMMIT: ${{ steps.release.outputs.commit }}
        run: |
          set -euo pipefail
          RUNS="$(gh api "repos/$GITHUB_REPOSITORY/commits/$RELEASE_COMMIT/check-runs?per_page=100&filter=latest" --jq '.check_runs[] | [.name, .conclusion] | @tsv')"
          for required in "CI complete" "ses-interop"; do
            if ! grep -Fxq "$required"$'\t'"success" <<< "$RUNS"; then
              echo "Release commit $RELEASE_COMMIT has no successful '$required' check run."
              exit 1
            fi
          done
```

**Verify**: YAML parse → `ok`. Then run the same query read-only for the last
release to confirm the names still match:
`gh api "repos/{owner}/{repo}/commits/$(git rev-list -n1 v0.3.1)/check-runs?per_page=100&filter=latest" --jq '.check_runs[] | [.name, .conclusion] | @tsv' | grep -c -E "^(CI complete|ses-interop)	success$"`
→ `2`. If `gh` is unavailable or unauthenticated, record that this check was not
run.

### Step 4: Run the advisory gate and script tests

Insert directly after the `Restore` step:

```yaml
      - name: Reject NuGet advisories outside the allowlist
        run: python3 scripts/check-nuget-advisories.py

      - name: Test repository scripts
        run: python3 -m unittest discover -s scripts/tests -v
```

(These are the commands at `ci.yml:77` and `ci.yml:86`.)

**Verify**: YAML parse → `ok`.

### Step 5: Move login and push into a small `publish` job

1. In `validate`, keep `Require candidate-bound macOS Preview acceptance`. Move
   the line
   `python3 scripts/verify-package-artifacts.py artifacts/nuget "${{ env.PACKAGE_VERSION }}" --verify-manifest`
   into a new final `validate` step named `Verify package manifest`, followed by
   a step that writes checksums and uploads the packages:

```yaml
      - name: Record package checksums
        shell: bash
        run: cd artifacts/nuget && sha256sum *.nupkg *.snupkg > SHA256SUMS

      - name: Upload validated packages
        uses: actions/upload-artifact@v4
        with:
          name: release-packages
          compression-level: 0 # nupkg/snupkg payloads are already compressed.
          path: |
            artifacts/nuget/*.nupkg
            artifacts/nuget/*.snupkg
            artifacts/nuget/SHA256SUMS
          if-no-files-found: error
```

2. Delete `NuGet login (OIDC)` and `Publish packages and symbols` from
   `validate` and add a second job:

```yaml
  publish:
    needs: validate
    runs-on: ubuntu-latest
    permissions:
      contents: read
      id-token: write
    env:
      DOTNET_SKIP_FIRST_TIME_EXPERIENCE: "1"
      DOTNET_CLI_TELEMETRY_OPTOUT: "1"
    steps:
      - name: Setup .NET
        uses: actions/setup-dotnet@v4
        with:
          dotnet-version: "10.0.x"

      - name: Download validated packages
        uses: actions/download-artifact@v4
        with:
          name: release-packages
          path: packages

      - name: Verify package checksums
        shell: bash
        run: cd packages && sha256sum --check --strict SHA256SUMS

      - name: NuGet login (OIDC)
        id: nuget-login
        uses: NuGet/login@8d196754b4036150537f80ac539e15c2f1028841 # v1
        with:
          user: whynpc

      - name: Publish packages and symbols
        shell: bash
        run: |
          shopt -s nullglob
          for package in packages/*.nupkg; do
            dotnet nuget push "$package" --source "https://api.nuget.org/v3/index.json" --api-key "${{ steps.nuget-login.outputs.NUGET_API_KEY }}" --skip-duplicate
          done
```

   The `publish` job does not check out the repository and runs no repository
   code. Keep it that way.

**Verify** (all must hold):

- YAML parse → `ok`
- `python3 - <<'PY'` check:

```python
import yaml
w = yaml.safe_load(open('.github/workflows/publish-nuget.yml'))
j = w['jobs']
assert set(j) == {'validate', 'publish'}, set(j)
assert 'id-token' not in j['validate']['permissions']
assert j['validate']['permissions'].get('checks') == 'read'
assert j['publish']['permissions'] == {'contents': 'read', 'id-token': 'write'}
assert j['publish']['needs'] == 'validate'
assert not any('checkout' in str(s.get('uses', '')) for s in j['publish']['steps'])
assert 'environment' not in j['validate'] and 'environment' not in j['publish']
names = [s.get('name', '') for s in j['validate']['steps']]
for n in ['Require CI and SES success on the release commit', 'Reject NuGet advisories outside the allowlist',
          'Test repository scripts', 'Require candidate-bound macOS Preview acceptance', 'Upload validated packages']:
    assert n in names, n
assert names.index('Require candidate-bound macOS Preview acceptance') < names.index('Upload validated packages')
print('structure ok')
PY
```

  → prints `structure ok`.
- `actionlint` (if installed) → no output.

### Step 6: Document the gates

- `docs/engineering/ci.md`: add a section `## Release workflow gates` (English,
  like the rest of the file) stating: the release workflow requires successful
  `CI complete` and `ses-interop` check runs on the tagged commit; it runs the
  advisory gate and script tests; validation and publishing are separate jobs
  and only the `publish` job holds `id-token: write`; a protected GitHub
  environment with required reviewers is a recommended manual follow-up that
  must be coordinated with the nuget.org Trusted Publishing policy.
- `docs/engineering/release-gates.md`: in the `## 发布候选顺序` list, item 4,
  add one Chinese sentence saying the tag workflow first requires the tagged
  commit to have successful `CI complete` and `ses-interop` check runs, and that
  packages are pushed by a separate job that only downloads the validated
  artifacts. Do not edit any other paragraph of that file.

**Verify**: `grep -n "Release workflow gates" docs/engineering/ci.md` → 1 match;
`grep -n "ses-interop" docs/engineering/release-gates.md` → at least 1 match.

## Test plan

A release workflow cannot be executed by the executor. Verification is the
structural check in Step 5, `actionlint` when available, and the read-only check
run query in Step 3. End-to-end proof is an operator action after merge (see
Maintenance notes); report it as "not performed".

## Done criteria

- [ ] YAML parse → `ok`; structural script prints `structure ok`
- [ ] `grep -c "id-token: write" .github/workflows/publish-nuget.yml` → `1`
- [ ] Docs greps of Step 6 match
- [ ] `git status --short` lists only in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated; the report lists "end-to-end release rehearsal: not performed (operator action)"

## STOP conditions

Stop and report back (do not improvise) if:

- The step list in "Current state" does not match the live workflow.
- The check run names on the latest release commit are not `CI complete` and
  `ses-interop`.
- You believe a validation step must move into the `publish` job, or the
  `publish` job needs to check out the repository.
- Anything suggests the nuget.org Trusted Publishing policy is bound to the job
  name `publish` in a way this split would break (for example a comment in the
  repository docs). Report it; do not guess.

## Maintenance notes

- **Operator rehearsal after merge:** run the workflow manually from `main` with
  `release_tag` set to the latest already-published tag (for example `v0.3.1`).
  Expected: `validate` passes, `publish` logs in and every push reports the
  package already exists (`--skip-duplicate`). This exercises the two-job
  hand-off without publishing anything new. If nuget.org rejects the login,
  the Trusted Publishing policy needs updating before the next real release.
- A release prepared on a commit whose `CI` or `SES dev interop` run failed or
  never ran will now fail at the new step; rerun those workflows for that commit
  first.
- Follow-up for the maintainer (not automated here): create a `nuget-release`
  environment with required reviewers and a `v*` tag rule, add
  `environment: nuget-release` to the `publish` job, and set the same environment
  in the nuget.org policy.
- Plan 025 adds `timeout-minutes` to workflow jobs, including these two.
