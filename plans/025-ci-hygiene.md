# Plan 025: Weekly advisory check, job timeouts and explicit workflow permissions

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- .github/workflows docs/engineering/ci.md`
> `publish-nuget.yml` is expected to differ if plan 024 has landed (see Step 3).
> For every other in-scope file that changed, compare the "Current state"
> excerpts against the live code before proceeding; on a mismatch, treat it as a
> STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none (works before or after plan 024)
- **Category**: dx (CI)
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

Three small gaps in the GitHub workflows:

1. The NuGet advisory gate (`scripts/check-nuget-advisories.py`) only runs when
   someone pushes. An advisory published during a quiet week goes unnoticed.
   `docs/engineering/ci.md` already says "a separate small scheduled workflow
   would also catch advisories published between pushes"; none exists.
2. No job has `timeout-minutes`, so each uses GitHub's 360-minute default. The
   jobs run LibreOffice and an owned-process supervisor; a hang burns up to six
   hours per job (up to nine jobs in parallel, including macOS runners).
3. `ci.yml` and `ses-interop.yml` have no `permissions:` block and rely on the
   repository default staying read-only. The Pages workflow cancels an in-flight
   deployment when a new push arrives.

## Current state

Workflows under `.github/workflows/` and their jobs at commit `427433a`:

| File | Jobs (job id) | `permissions:` | `timeout-minutes` |
|---|---|---|---|
| `ci.yml` | `build-test`, `builtin-docx` (matrix of 3 OS), `package-e2e`, `windows-package-e2e`, `macos-package-e2e`, `password-crypto`, `ci-complete` | none | none |
| `ses-interop.yml` | `ses-interop` | none | none |
| `docs-pages.yml` | `build`, `deploy` | top level `contents: read`; `deploy` adds `pages: write`, `id-token: write` | none |
| `publish-nuget.yml` | `publish` (or `validate` + `publish` once plan 024 has landed) | per job | none |

No workflow has a `schedule:` trigger (`grep -rn "schedule:" .github/workflows`
prints nothing).

`ci.yml` top (lines 1–11):

```yaml
name: CI

on:
  push:
    branches: ["main", "master", "codex/document-tools"]
  pull_request:
    branches: ["main", "master", "codex/**"]

concurrency:
  group: ${{ github.workflow }}-${{ github.event_name == 'pull_request' && github.ref || github.run_id }}
  cancel-in-progress: ${{ github.event_name == 'pull_request' }}
```

`ses-interop.yml` top (lines 1–9):

```yaml
name: SES dev interop
on:
  pull_request:
    branches: ["main", "master", "codex/document-tools"]
  push:
    branches: ["main", "master"]
concurrency:
  group: ${{ github.workflow }}-${{ github.event_name == 'pull_request' && github.ref || github.run_id }}
  cancel-in-progress: ${{ github.event_name == 'pull_request' }}
```

`docs-pages.yml` (lines 8–13 and the `deploy` job, lines 38–50):

```yaml
permissions:
  contents: read
concurrency:
  group: docs-pages-${{ github.ref }}
  cancel-in-progress: true
...
  deploy:
    if: github.event_name != 'pull_request'
    needs: build
    runs-on: ubuntu-latest
    permissions:
      pages: write
      id-token: write
    environment:
      name: github-pages
      url: ${{ steps.deployment.outputs.page_url }}
```

How CI runs the advisory gate today (`ci.yml:52-77`, shortened):

```yaml
      - name: Setup .NET
        uses: actions/setup-dotnet@v4
        with:
          dotnet-version: "10.0.x"
      ...
      - name: Restore
        run: dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

      - name: Reject NuGet advisories outside the reviewed allowlist
        run: python3 scripts/check-nuget-advisories.py
```

The script needs only a restored solution and Python 3. It exits 0 when clean,
1 on an unreviewed advisory, 2 when the report could not be produced. It does
not need Git LFS payloads, fonts or LibreOffice.

All jobs in `ci.yml` only read the repository and upload artifacts; none needs
a write permission. `ci.md` records a historical median wall time of 768
seconds for the whole CI workflow, so the timeouts below are several times the
normal duration.

## Commands you will need

| Purpose | Command | Expected on success |
|---|---|---|
| YAML parses | `for f in .github/workflows/*.yml; do python3 -c "import yaml,sys; yaml.safe_load(open(sys.argv[1]))" "$f" && echo "ok $f"; done` | one `ok` line per file |
| Lint (if installed) | `actionlint` | no output. If not installed, say so in the report. |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `.github/workflows/nuget-advisories.yml` (create)
- `.github/workflows/ci.yml` (top-level `permissions:` and per-job `timeout-minutes` only)
- `.github/workflows/ses-interop.yml` (same)
- `.github/workflows/docs-pages.yml` (timeouts and deploy concurrency only)
- `.github/workflows/publish-nuget.yml` (`timeout-minutes` only)
- `docs/engineering/ci.md` (the "NuGet advisory gate" section)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- Any step, trigger, matrix, cache key, artifact name or job id. The
  `ci-complete` job and its `needs:` list must stay exactly as they are.
- `.github/actions/**`.
- Pinning GitHub-owned actions to commit SHAs (not selected).

## Git workflow

- Branch: `codex/ci-hygiene`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or trigger workflows unless the operator instructed it.

## Steps

### Step 1: Add the scheduled advisory workflow

Create `.github/workflows/nuget-advisories.yml`:

```yaml
name: NuGet advisories

on:
  schedule:
    - cron: "17 3 * * 1" # Mondays 03:17 UTC
  workflow_dispatch:

permissions:
  contents: read

# No fixtures are needed; skip LFS payloads.
env:
  GIT_LFS_SKIP_SMUDGE: "1"
  DOTNET_CLI_HOME: ${{ github.workspace }}/artifacts/dotnet-home
  DOTNET_SKIP_FIRST_TIME_EXPERIENCE: "1"
  DOTNET_CLI_TELEMETRY_OPTOUT: "1"

jobs:
  advisories:
    runs-on: ubuntu-latest
    timeout-minutes: 20
    steps:
      - name: Checkout
        uses: actions/checkout@v4
        with:
          lfs: false

      - name: Setup .NET
        uses: actions/setup-dotnet@v4
        with:
          dotnet-version: "10.0.x"

      - name: Restore
        run: dotnet restore Ofdrw.Net.sln --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false

      - name: Reject NuGet advisories outside the reviewed allowlist
        run: python3 scripts/check-nuget-advisories.py
```

**Verify**: YAML parse command prints `ok .github/workflows/nuget-advisories.yml`.
Locally, after `dotnet restore Ofdrw.Net.sln ...` (with the `DOTNET_CLI_HOME`
variables from `AGENTS.md` set), `python3 scripts/check-nuget-advisories.py`
exits 0. If the restore needs network access you do not have, record that the
local run was not performed.

### Step 2: Explicit read-only permissions

Add this block to `ci.yml` and to `ses-interop.yml`, directly before their
`concurrency:` line:

```yaml
permissions:
  contents: read
```

**Verify**: `grep -c "^permissions:" .github/workflows/ci.yml .github/workflows/ses-interop.yml`
→ `1` for each file.

### Step 3: Timeouts on every job

Add `timeout-minutes:` directly under each job's `runs-on:` line:

| File | Job id | `timeout-minutes` |
|---|---|---:|
| `ci.yml` | `build-test` | 60 |
| `ci.yml` | `builtin-docx` | 90 |
| `ci.yml` | `package-e2e` | 60 |
| `ci.yml` | `windows-package-e2e` | 90 |
| `ci.yml` | `macos-package-e2e` | 90 |
| `ci.yml` | `password-crypto` | 45 |
| `ci.yml` | `ci-complete` | 10 |
| `ses-interop.yml` | `ses-interop` | 60 |
| `docs-pages.yml` | `build` | 20 |
| `docs-pages.yml` | `deploy` | 15 |
| `publish-nuget.yml` | `publish` if it is the only job | 90 |
| `publish-nuget.yml` | `validate` (present after plan 024) | 90 |
| `publish-nuget.yml` | `publish` when a `validate` job exists | 15 |

**Verify**:

```sh
python3 - <<'PY'
import glob, yaml
missing = []
for path in sorted(glob.glob('.github/workflows/*.yml')):
    for job_id, job in yaml.safe_load(open(path))['jobs'].items():
        if not isinstance(job.get('timeout-minutes'), int):
            missing.append(f'{path}:{job_id}')
print('missing:', missing)
assert not missing
PY
```

→ prints `missing: []`.

### Step 4: Do not cancel an in-flight Pages deployment

In `docs-pages.yml`:

1. Change the top-level `cancel-in-progress: true` to
   `cancel-in-progress: ${{ github.event_name == 'pull_request' }}`.
2. Give the `deploy` job its own group so deployments queue instead of
   overlapping:

```yaml
    concurrency:
      group: docs-pages-deploy
      cancel-in-progress: false
```

**Verify**: YAML parse → `ok`;
`python3 -c "import yaml; w=yaml.safe_load(open('.github/workflows/docs-pages.yml')); assert w['jobs']['deploy']['concurrency']['cancel-in-progress'] is False; print('ok')"`
→ `ok`.

### Step 5: Prove nothing else changed in `ci.yml`

```sh
git diff 427433a -- .github/workflows/ci.yml | grep '^[-+]' | grep -v '^+++\|^---' | grep -v 'timeout-minutes:\|^+permissions:\|^+  contents: read\|^+$'
```

**Verify**: prints nothing (every changed line is a timeout, the permissions
block, or a blank line). Run the same check for `ses-interop.yml`.

### Step 6: Update the engineering note

In `docs/engineering/ci.md`, section `## NuGet advisory gate`, replace the last
sentence ("The gate runs only when CI runs; a separate small scheduled workflow
would also catch advisories published between pushes.") with a statement that
`nuget-advisories.yml` runs the same gate every Monday and on manual dispatch,
and that a failed scheduled run means a new advisory needs review (fix the
dependency, or add an allowlist entry with a recorded decision). Add one
sentence that every job has a `timeout-minutes` limit and that the limits are
several times the measured medians.

**Verify**: `grep -n "nuget-advisories.yml" docs/engineering/ci.md` → at least 1 match.

## Test plan

Workflows cannot be executed locally. Verification is YAML parsing, the two
structural scripts above, the "only expected lines changed" diff check, and
`actionlint` when available. Report "workflow runs on GitHub: not performed".

## Done criteria

- [ ] Every workflow file parses; `actionlint` clean or reported as unavailable
- [ ] Timeout script prints `missing: []`
- [ ] `grep -rn "schedule:" .github/workflows` → exactly one match, in `nuget-advisories.yml`
- [ ] Step 5 diff checks print nothing for `ci.yml` and `ses-interop.yml`
- [ ] `git status --short` lists only in-scope files
- [ ] `git add -A && python3 scripts/check-repository-storage.py` exits 0
- [ ] `plans/plan-index.md` status row updated

## STOP conditions

Stop and report back (do not improvise) if:

- A job in `ci.yml` or `ses-interop.yml` turns out to need a permission other
  than `contents: read` (search the file for `gh `, `github.token`, `GH_TOKEN`,
  `id-token`, `packages:`; at `427433a` there are none).
- The job lists in "Current state" do not match the live files, apart from the
  `publish-nuget.yml` split that plan 024 introduces.
- You are tempted to change a step, cache key or the `ci-complete` gate.

## Maintenance notes

- A scheduled workflow only runs from the default branch and GitHub disables
  schedules in repositories with no activity for 60 days; the maintainer may
  need to re-enable it.
- A scheduled failure does not block anything by itself. Decide who watches it
  (GitHub notifies the user who last edited the cron line).
- If a legitimate job starts hitting its timeout, raise that number with the
  measured duration in the commit message rather than removing the limit.
