# Plan 010: Compile every e2e harness in CI so orphaned acceptance projects cannot rot

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/README.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 09a2e7e..HEAD -- .github/workflows/ci.yml e2e scripts`
> If the e2e project list or the `build-test`/`package-e2e` jobs changed since this plan was written, re-derive the orphan list (Step 1) before proceeding.

## Status

- **Priority**: P2
- **Effort**: M
- **Risk**: LOW (CI-only; no product code)
- **Depends on**: none
- **Category**: tests / dx
- **Planned at**: commit `09a2e7e`, 2026-11-03

## Why this matters

Several acceptance harnesses under `e2e/` are referenced by no solution, workflow, composite action or CI-invoked script. They include signature appearance/scope, HTML preview, archival preflight, typed objects, R09 graphics and benchmarks. These harnesses back recorded acceptance evidence. Because nothing builds them, an SDK API change can break them silently, and re-running the evidence later fails or tests a stale package. A build-only gate is cheap and catches API drift on every PR.

## Current state

- `Ofdrw.Net.sln` and `Ofdrw.Net.PlatformTests.slnf` include some e2e projects; CI runs others through scripts (`scripts/run-converter-package-e2e.sh`, `run-document-tools-e2e.sh`, `run-graphics-e2e.sh`, `run-skia-package-e2e.sh`, `run-pdf-vector-package-e2e.sh`), all called from `.github/workflows/ci.yml` (lines ~115-361).
- At `09a2e7e` the following have zero references, or are referenced only by scripts that are themselves never invoked:

  | Project | Reference style |
  |---|---|
  | `e2e/Ofdrw.Net.Graphics.R09.E2E` | ProjectReference only |
  | `e2e/Ofdrw.Net.Signatures.Appearance.Visual.E2E` | ProjectReference only |
  | `e2e/Ofdrw.Net.TypedObjects.E2E` | ProjectReference when `OfdrwPackageVersion` is empty, else PackageReference |
  | `e2e/Ofdrw.Net.Signatures.Scope.E2E` | ProjectReference when `OfdrwScopePackageVersion` is empty |
  | `e2e/Ofdrw.Net.Benchmarks` | PackageReference `Ofdrw.Net.Converter` only (`OfdrwPackageVersion`, default `0.1.0-preview.8`) |
  | `e2e/Ofdrw.Net.Packaging.Preflight.E2E` | PackageReference `Ofdrw.Net.Packaging` (default `0.1.0-preview.8`) |
  | `e2e/Ofdrw.Net.Converter.Svg.Html.E2E` | PackageReference `Converter.Svg`, `Packaging` (default `0.1.0-preview.8`) |
  | `e2e/Ofdrw.Net.Signatures.Appearance.E2E` | PackageReference `Packaging`, `Signatures`, `Signatures.SesInterop` |
  | `e2e/Ofdrw.Net.AttachmentCleanup.E2E` | PackageReference `Core`, `Packaging`, `Reader` |

  Unreferenced scripts: `scripts/run-html-preview-package-e2e.sh`, `run-signature-appearance-package-e2e.sh`, `run-signature-scope-package-e2e.sh`, `run-attachment-cleanup-package-e2e.sh`.
- CI jobs:
  - `build-test` builds the solution in its step "Build and reject new undocumented public API" (`python3 scripts/check-public-api-docs.py --package-version ...`) and later packs candidates.
  - `package-e2e` (line 298+) downloads the `candidate-packages` artifact into `artifacts/candidate-packages`, then runs the consumer scripts.
  - Env flags used everywhere: `--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false`.
  - CI version env: `OFDRW_CI_PACKAGE_VERSION: "0.1.0-preview.10"`.
- NuGet.Config pattern used by the package scripts (`scripts/run-signature-scope-package-e2e.sh:41-49`):
  ```python
  root=ET.Element('configuration'); sources=ET.SubElement(root,'packageSources'); ET.SubElement(sources,'clear')
  mapping=ET.SubElement(root,'packageSourceMapping')
  for key,path,pattern in [('r13-local',sys.argv[2],'Ofdrw.Net.*'),('cache',sys.argv[3],'*'),('nuget','https://api.nuget.org/v3/index.json','*')]:
      ET.SubElement(sources,'add',key=key,value=path)
      ET.SubElement(ET.SubElement(mapping,'packageSource',key=key),'package',pattern=pattern)
  ```
  The scripts also set `NUGET_PACKAGES` to a fresh temp dir, so stale cached `Ofdrw.Net.*` packages cannot satisfy the restore.
- Repository rule (AGENTS.md): do not rewrite historical evidence. The e2e csproj default versions may be pinned by evidence manifests, so **do not edit e2e csproj files**; pass versions on the command line instead.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$PWD/artifacts/dotnet-home" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
F="--disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false"
```

| Purpose | Command | Expected |
|---|---|---|
| Orphan census | the Step 1 loop | the list above, or a re-derived one |
| Project-ref build | `dotnet build e2e/<P>/<P>.csproj -c Release $F` | exit 0 |
| Local candidate | `./scripts/run-converter-package-e2e.sh 0.1.0-preview.10 --pack-only` | packages under `artifacts/package-e2e/0.1.0-preview.10/packages` |
| New script | `scripts/build-e2e-harnesses.sh 0.1.0-preview.10 artifacts/package-e2e/0.1.0-preview.10/packages` | exit 0 |
| Workflow lint | `python3 -c "import yaml,sys; yaml.safe_load(open('.github/workflows/ci.yml'))"` | no exception (if PyYAML is missing, `pip install pyyaml` into a venv under `artifacts/`) |
| Script tests | `python3 -m unittest discover -s scripts/tests -v` | pass |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope**: `scripts/build-e2e-harnesses.sh` (create); `.github/workflows/ci.yml` (one new step in `build-test`, one in `package-e2e`); `docs/engineering/ci.md` (one paragraph).

**Out of scope**: any file under `e2e/`, including the stale default versions (do not edit them); running the harnesses (build only); the four orphan `run-*.sh` scripts; `Ofdrw.Net.sln`, since adding projects changes `dotnet test Ofdrw.Net.sln` behavior.

## Git workflow

Branch `advisor/010-e2e-harness-build-gate`; commit e.g. `ci: compile every e2e harness against source and candidate packages`. No push or PR unless instructed.

## Steps

### Step 1: Re-derive the orphan list

```sh
for p in e2e/*/; do n=$(basename $p); c=$(grep -l "$n" Ofdrw.Net.sln Ofdrw.Net.PlatformTests.slnf .github/workflows/*.yml .github/actions/*/*.yml scripts/*.sh scripts/*.py 2>/dev/null | tr '\n' ' '); echo "$n :: $c"; done
```
A project is orphaned if it is listed only in files that are themselves never invoked by CI. Check with `grep -n "<script name>" .github/workflows/ci.yml`. Also skip `e2e/visual-corpus` (not a project) and `e2e/Ofdrw.Net.Graphics.Sample` (referenced by Graphics E2E).

**Verify**: a written list matching the table above, or a documented difference.

### Step 2: Project-reference build gate

Locally build each ProjectReference-capable orphan (`Graphics.R09.E2E`, `Signatures.Appearance.Visual.E2E`, `TypedObjects.E2E`, `Signatures.Scope.E2E`) with `dotnet build ... -c Release $F`.
- If one fails to compile on current `main`, that is the rot this plan exists to catch. **STOP** and report the compiler errors; do not fix e2e code in this plan.
- If all succeed, add a `build-test` step right after "Build and reject new undocumented public API":
  ```yaml
  - name: Compile source-referenced e2e harnesses
    run: |
      for p in Graphics.R09.E2E Signatures.Appearance.Visual.E2E TypedObjects.E2E Signatures.Scope.E2E; do
        dotnet build "e2e/Ofdrw.Net.$p/Ofdrw.Net.$p.csproj" -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
      done
  ```

**Verify**: the local loop exits 0; the YAML lint passes.

### Step 3: Candidate-package build script

Create `scripts/build-e2e-harnesses.sh VERSION PACKAGES_DIR` (bash, `set -euo pipefail`, same `FLAGS` array and env exports as `run-signature-scope-package-e2e.sh:15-19`). It should:
- Create a temp dir via `mktemp -d` with `trap` cleanup.
- Write a NuGet.Config with the pattern above (local `PACKAGES_DIR` mapped to `Ofdrw.Net.*`; `${NUGET_PACKAGES:-$HOME/.nuget/packages}` cache and nuget.org for `*`), and set `NUGET_PACKAGES` to a fresh dir under the temp dir.
- For each package-only orphan (`Benchmarks`, `Packaging.Preflight.E2E`, `Converter.Svg.Html.E2E`, `Signatures.Appearance.E2E`, `AttachmentCleanup.E2E`):
  - Read its `PackageReference Include="Ofdrw.Net.*"` names (`grep -o`).
  - If any required `NAME.VERSION.nupkg` is missing from `PACKAGES_DIR`, print `SKIP <project>: candidate lacks <name>` and continue. Do not fail; `Signatures.SesInterop` is likely absent.
  - Otherwise `dotnet restore <csproj> --configfile <cfg> -p:OfdrwPackageVersion=$VERSION "${FLAGS[@]}"` and `dotnet build <csproj> -c Release --no-restore -p:OfdrwPackageVersion=$VERSION "${FLAGS[@]}"`.
- Finally print a summary of built and skipped projects. Exit non-zero only if a restore or build fails.

Run it locally against a locally packed candidate (the "Local candidate" command).

**Verify**: exit 0. Built projects compile; skips are only for missing packages.

### Step 4: Wire into `package-e2e`

After "Consume the exact candidate packages on Linux", add:
```yaml
- name: Compile package-referenced e2e harnesses against the candidate
  run: ./scripts/build-e2e-harnesses.sh ${{ env.OFDRW_CI_PACKAGE_VERSION }} artifacts/candidate-packages
```
`chmod +x` the script and commit the mode (`git update-index --chmod=+x`).

**Verify**: the YAML lint passes; `git ls-files -s scripts/build-e2e-harnesses.sh` shows mode `100755`.

### Step 5: Document

In `docs/engineering/ci.md`, add a short paragraph listing the two new steps, explaining that they are build-only, and saying how to add a future harness to the gate.

## Done criteria

- [ ] All four ProjectReference orphans build locally; the CI step is added
- [ ] `scripts/build-e2e-harnesses.sh` builds the package-only orphans whose packages exist in the candidate and skips the others with a message
- [ ] `git diff --stat -- e2e` is empty
- [ ] YAML lint and script unit tests pass; storage gate exits 0
- [ ] `plans/README.md` updated. In the report, include the list of skipped projects and why.

## STOP conditions

- Any orphan fails to compile on current source or packages. Report the errors; fixing them is a separate task.
- A harness requires network access at build time beyond NuGet (e.g. `go` tooling for Signatures.Scope's Go verifier is not needed for the C# build; if the csproj invokes it during build, report it).
- The CI job time grows by more than ~5 minutes (measure locally and report).

## Maintenance notes

- Any new e2e project must be added to one of the two loops, or to the sln/slnf. Reviewers should check `e2e/` additions against `scripts/build-e2e-harnesses.sh`.
- Running (not just compiling) the signature, HTML and preflight harnesses in CI would need their visual tooling and is a larger decision.
