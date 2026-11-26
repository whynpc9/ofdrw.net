# CI performance and evidence

CI keeps three-platform full-solution coverage, Linux no-Office test filters, package-consumer isolation (including the optional SES consumer), rendering comparisons, and the release Preview gate. Test selection is all-or-nothing: a run either executes every heavy job or, when the tree is already validated, skips all of them. There is no per-test or per-project path selection.

## Change planning and validation markers

The `changes` job runs `scripts/ci-plan.py decide` and every heavy job (`build-test`, `builtin-docx`, `password-crypto`, and through `build-test` the three package jobs) runs only when it reports `run_full == 'true'`. Decision order:

1. A manual dispatch with `force_full` (the default) or `parallel_tests` runs everything.
2. A valid marker for the checked-out tree (`HEAD^{tree}`) skips the heavy jobs (mode `reused`). This is the common main push: the merge result is the tree its PR already passed.
3. If every changed file is skippable documentation and the diff base tree has a valid marker, the heavy jobs are skipped (mode `docs-only`). The base is the merge commit's first parent for pull requests and the push event's `before` commit for pushes; a new branch, a force push or any other event has no base.
4. Otherwise everything runs (mode `full`).

Every API, network, payload or git error falls back to a full run; the planner never skips because a lookup failed. A docs-only run still runs `check-repository-storage.py` in the `changes` job, because `build-test` normally owns that check.

Skippable paths are `docs/**`, `plans/**` and root-level `*.md`, except the files that the build, packages, tests or CI scripts consume: `README.md`, `THIRD-PARTY-NOTICES.md`, and the `CONSUMED_DOCS` list in `scripts/ci-plan.py` (packed licenses and guides, the release note packed into every package, test-read references, engineering baselines, public-table evidence and release Preview records). The guard test in `scripts/tests/test_ci_plan.py` scans project files, source/test code, every workflow except `docs-pages.yml`, the composite actions, the script tests and every script they reach, and fails when one of them references a docs path that the planner would skip. Add a newly consumed document to `CONSUMED_DOCS` when that test fails.

A validation marker is an artifact named `ci-validated-<git tree SHA>`, retained for 30 days. `CI complete` uploads it after its gate passes in `full` or `docs-only` mode (never for `reused` or a `parallel_tests` trial). The JSON records the tree, commit, run, event and mode, and for `docs-only` the validated base tree. It means: this exact tree passed the full suite, or differs only by skippable files from a tree that did. A marker is accepted only when it is unexpired, was produced by a `.github/workflows/ci.yml` run that completed with `success`, and that run's head repository is this repository (fork runs are rejected); at most ten candidates are checked. Trust assumption: any branch of this repository whose CI run succeeds can upload a marker, so marker trust equals push access to the repository. A marker carries no other proof.

A green `CI complete` therefore means one of two things. In `full` mode, `changes` succeeded and every heavy job, including the Linux SES consumer, succeeded in this run. In `reused`/`docs-only` mode, every heavy job was skipped as planned because an earlier successful CI run validated this tree or its documentation-only base. Any other mixture of results (a failure, cancellation, or a heavy job skipped in a full run) fails the gate. The job summary of `changes` records the mode, reason and changed files.

## Build and caches

- The public API documentation checker owns the full Release rebuild. A preceding ordinary build is unnecessary. Its CS1591 baseline and failure handling remain unchanged.
- Linux `build-test` runs the full Release solution and mandatory trusted-font suite. The `builtin-docx` matrix builds the full solution and runs those same suites on Windows and macOS, retaining native-path results and skip reasons in `full-platform-tests-<os>`.
- Linux no-Office steps run inside `build-test`, right after the API-documentation build and before `libreoffice-writer` and `poppler-utils` are installed: a guard fails if LibreOffice is visible at any location `LibreOfficeExecutableResolver` searches, then the `BuiltIn|Auto|Native|DualLayer` DOCX filter (TRX in `artifacts/no-office-tests`) and a CLI `docx-to-pdf --docx-engine auto` run exercise engine selection and fallback without Word or LibreOffice. Only `zstd` is installed earlier, for the evidence-archive script tests and the public-table verifier. The DOCX CJK viewer-local tests also run once with the pinned trusted fixture after the mandatory font gate. The separate ubuntu matrix leg and `Ofdrw.Net.PlatformTests.slnf` were removed; its other filtered reruns could not depend on tool presence and already ran in the same job's full `Test` step against the same build (see the 2026-11-16 section).
- Document-tools E2E and its CLI use Release consistently.
- Only third-party build dependencies are cached. Consumers retain a fresh NuGet directory and source mapping that binds `Ofdrw.Net.*` to the exact candidate feed.
- CI fonts are cached after static instancing. Both the generated font and license must match source-controlled SHA-256 values on every hit. Invalid cache entries regenerate from the pinned source; source or generated-byte mismatches fail.
- Git LFS fixtures (`e2e/**`, `tests/**`) come from `.github/actions/lfs-fixtures`: one cache entry keyed by the sorted object IDs, restorable on all three OSes, re-hashed against those IDs before use, with missing objects fetched and the working tree materialized by `git lfs pull`. Caches saved by a pull-request run are visible only to that PR, while caches saved on `main` are visible to every branch. Because most `main` pushes now skip the heavy jobs, the `changes` job warms the cache (`warm-only`) on the default branch and before any full run; an exact hit is a lookup only, and a warming failure does not fail the plan. `build-test` then calls the action a second time with `include: docs/validation/public-tables-evidence/**`, which keeps a separate cache entry. This call prunes the fixture objects from `.git/lfs/objects`, which is safe because by then the fixtures are materialized, their cache entry is saved, and no later step in that job runs `git`. Every LFS byte that Actions downloads counts against the account's Git LFS bandwidth quota, so new LFS consumers in CI should use this action rather than a plain `git lfs pull`. The release workflow is the only exception: it keeps a plain `git lfs pull` because local actions resolve from the checked-out release commit, and older tags lack the action.
- Linux caches downloaded apt packages, but still refreshes repository metadata and installs dependencies normally. Windows caches micromamba package downloads, not a relocated installation prefix. These caches reduce downloads; they do not eliminate all solver, repository, or network latency.
- Windows download cache keys include architecture and pinned bootstrap/tool versions. Bump the matching key when changing those versions; unrelated edits to `ci.yml` no longer invalidate the downloads. Candidate nupkg/snupkg files use artifact compression level zero because their payloads are already compressed.
- Full Windows/macOS solution tests install LibreOffice and Poppler before running; the PDF import and DOCX conversion tests require those external tools. The pinned Windows Poppler bootstrap is hash-checked. The package visual jobs install their own renderer tools on separate runners.
- The package visual corpus embeds the same verified Noto CJK font bytes on all three platforms. Its four macOS golden pages were aligned with the already reviewed Linux/Windows bytes after the CI runner and a separate local macOS capture produced identical page hashes; earlier host-font pages remain recoverable from Git history.
- New PR commits cancel older runs of the same workflow and PR. Main and release runs are not cancelled by this policy.

## One candidate, parallel consumers

`ci.yml` declares one `OFDRW_CI_PACKAGE_VERSION` (currently `0.4.0`). Immediately after checkout, `build-test` verifies that it matches the default `VersionPrefix` and `VersionSuffix` in `Directory.Build.props`; a source-version change without a matching CI update fails before the build. The same variable stamps the API-documentation rebuild and drives pack-only, candidate artifact paths, all three platform consumers, optional Skia/vector consumers, and the Linux mixed-text reader evidence paths.

`build-test` tests that Release build and invokes `run-converter-package-e2e.sh` with the candidate version and `--pack-only --no-build`. Pack-only requires an empty feed. The thirteen default packages (`PACKAGES` in `verify-package-artifacts.py`, including the opt-in `Ofdrw.Net.Archival` added on main after 0.4.0) are validated and uploaded once, together with their SHA-256 manifest. The validator rejects any package whose transitive `Ofdrw.Net.*` dependencies reach an opt-in package, and the Converter consumer's resolved `deps.json` must not contain `Ofdrw.Net.Archival`. The optional `Ofdrw.Net.Signatures.SesInterop` package is packed from the same build into a separate `ses-interop-package` artifact; it never joins the default candidate feed.

Linux, native Windows, and native macOS package jobs depend directly on `build-test` and download that one artifact in parallel. Linux restores source metadata for license validation but does not rebuild the full source solution or repack the candidate. Each consumer verifies the existing manifest before and after use and retains a fresh dependency cache with package source mapping. The Windows and macOS consumers also compare each selected runtime DLL against its nupkg payload, isolated cache, and copied output. Optional Skia/vector gates still consume the same defaults and separately build their optional adapters. The source-based document and graphics probes retain their own required builds.

The Linux package job is also the SES consumer. After the default consumer it runs `run-ses-interop-e2e.sh OUTPUT_DIR VERSION --packages-dir artifacts/candidate-packages --optional-packages-dir artifacts/ses-interop-candidate`: it verifies the default feed against its manifest before and after use, checks that the optional feed holds only that SES package with exact `[VERSION]` dependencies on candidate packages, and builds nothing from `src/`. The SES-specific consumer, CLI reference-only checks, rendering and `verify-ses-interop-evidence.py` are unchanged; `candidate-source.json` binds the evidence to the source commit and both package hashes. The evidence is uploaded as `ses-interop-evidence`. The 44 `Ofdrw.Net.Signatures.SesInterop.Tests` cases are part of `Ofdrw.Net.sln`, so they run in the full-solution tests on all three platforms and their TRX is in `test-results` and `full-platform-tests-<os>`; the former `SES dev interop` workflow (`ses-interop.yml`), which rebuilt and re-consumed the default SDK, was removed.

The platform matrix runs independently. `CI complete` aggregates the planner and every heavy job as described above.

The shell scripts' default build-and-consume behavior remains available for standalone runs: `run-ses-interop-e2e.sh [OUTPUT_DIR] [VERSION]` still builds, packs and consumes locally. `--consume-only` supports existing release feeds, while `--pack-only --no-build` requires a caller-provided Release build stamped with the requested version. The package validator checks package metadata, dependency versions, payload requirements, and manifest hashes.

## NuGet advisory gate

Right after restore, `build-test` runs `scripts/check-nuget-advisories.py`, which lists every top-level and transitive package with known advisories (`dotnet list Ofdrw.Net.sln package --vulnerable --include-transitive --format json`) and fails on any (package, version, advisory) not in `docs/engineering/baselines/nuget-advisory-allowlist.json`. Allowlist entries that no longer occur are printed as stale warnings so fixed advisories get removed. An entry may only be added with a recorded decision and reference; the five ImageSharp 2.1.13 entries cite [D10](../validation/d10-imagesharp-security-assessment.md).

`nuget-advisories.yml` runs the same gate every Monday at 03:17 UTC and on manual dispatch. If a scheduled run fails, check the logs; for a new advisory, fix the dependency or record the decision before adding an allowlist entry. Every workflow job sets `timeout-minutes`, with limits several times the measured medians where timings are available.

## Release workflow gates

`publish-nuget.yml` requires a successful `CI complete` check from the `main` push run of `ci.yml` on the tagged commit. When the checked-out release commit still contains `.github/workflows/ses-interop.yml` (tags made before SES moved into `ci.yml`, for example in a manual recovery run), it additionally requires that workflow's `ses-interop` check, as before. That `CI complete` may come from a `reused` or `docs-only` plan; the `validate` job rebuilds and retests the release commit regardless. `scripts/verify-preview-record.py` fingerprints every tracked file, so documentation-only commits after a Preview record still invalidate that record at release time even though CI skips them. The `validate` job runs the NuGet advisory gate and repository script tests, consumes the exact packages, verifies the candidate-bound macOS Preview record and package manifest, and uploads the packages with SHA-256 checksums. The separate `publish` job downloads that artifact by ID and verifies its package inventory and hashes before signing in to NuGet. Only `publish` has `id-token: write`. Tag pushes and manual recovery runs for the same tag share a concurrency group.

As a manual follow-up, configure a protected GitHub environment with required reviewers and coordinate its name with the NuGet Trusted Publishing policy before attaching it to `publish`.

## Build-only e2e harness gate

Harnesses that back recorded acceptance evidence but are not built by any other CI step are compiled on every run; they are not executed. `build-test` builds the source-referenced harnesses (`Graphics.R09.E2E`, `Signatures.Appearance.Visual.E2E`, `TypedObjects.E2E`, `Signatures.Scope.E2E`) right after the API-documentation build. The Linux package job runs `scripts/build-e2e-harnesses.sh VERSION PACKAGES_DIR` after the candidate consumer: it restores the package-referenced harnesses (`Benchmarks`, `Packaging.Preflight.E2E`, `Converter.Svg.Html.E2E`, `Signatures.Appearance.E2E`) against the exact candidate with an isolated package cache and source mapping, building copies under the ignored `artifacts/` directory so `e2e/` and its evidence-pinned default versions stay untouched. A harness whose `Ofdrw.Net.*` packages are absent from the candidate is reported as `SKIP` (currently `Signatures.Appearance.E2E`, which needs `Signatures.SesInterop`). To gate a new harness, add it to the source loop in `ci.yml` if it uses ProjectReference, or to `HARNESSES` in the script if it only consumes packages.

## 2026-11-01 duplicate-work reduction

The Linux no-Office job described here has since been folded into `build-test`; see the 2026-11-16 section. Baseline: `b02162bef61abc17761349d891237fc2518d3989`. The latest 100 workflow records span 2026-10-30 01:11 to 2026-11-01 20:33 Asia/Shanghai: CI has 39 successes, 8 failures and 3 cancellations; SES has 49 successes and 1 cancellation. Step timings were collected for the latest 15 CI runs (2026-11-01 09:54–20:33), with medians below computed from the 14 successful runs. Failed/cancelled runs are not counted as performance wins.

The [latest successful run](https://github.com/whynpc9/ofdrw.net/actions/runs/41307050234) confirms that the full solution already runs all 107 password tests on Linux, Windows and macOS, with zero password-test skips. Each platform also passes the separate 74-case mandatory trusted-font gate. Windows/macOS then redundantly reran ten filtered test commands against the same build and environment; those commands now run only in the separate Linux no-Office job. The password job now runs only Linux fixture generation and both isolated package consumers. Its unit tests remain in all three full-solution runs.

| Historical work removed | Median elapsed seconds |
| --- | ---: |
| Dedicated Windows password test job | 117 |
| Dedicated macOS password test job | 87.5 |
| Linux password unit-test/build step | 66.5 |
| Windows filtered reruns and redundant native-path upload | 91.5 |
| macOS filtered reruns and redundant native-path upload | 60.5 |

The per-run sum of the removed jobs, steps and redundant native-path uploads has a median of **414.5 seconds (6.9 runner minutes)**, range 368–474 seconds, approximately **15.8%** of summed job time. CI expands to **9 jobs instead of 11**. These are historical gross-work measurements, not measured post-change savings: the retained crypto E2E build may now do work previously warmed by the deleted unit-test step, and runner/cache/network variance remains. The historical median workflow wall time is 768 seconds; the `build-test` → Linux `package-e2e` chain remains the main critical path, so wall-time improvement will be smaller than runner-time reduction.

The sampled [failed run](https://github.com/whynpc9/ofdrw.net/actions/runs/41278534258) failed the full-solution `Test` step and the aggregate gate; both gates remain. No test cases, rendering pages, visual comparisons, candidate checks, SES acceptance, release Preview requirements, workflow triggers or artifact retention windows were removed. Windows/macOS native-path records now live in `full-platform-tests-<os>` instead of a second filtered artifact. Job IDs used by `CI complete` are unchanged; that gate still rejects any failed, cancelled or skipped dependency.

Local validation passed: `actionlint` 1.7.12 over all workflows, `git diff --check`, the repository storage checker, and all 15 existing evidence/font-cache/package-artifact Python tests (zero skips, after solution restore). A parsed before/after comparison also verified effective environments, Linux no-Office commands, full-solution gates, consumer commands, triggers and the aggregate gate. Validation records and the raw API timing snapshots live under the ignored `artifacts/ci-slimming-20261101/`. GitHub-hosted post-change timing still needs a new CI run. This configuration-only change does not regenerate documents or claim new PNG/Preview visual acceptance.

## Parallel test trial

All full-solution `dotnet test Ofdrw.Net.sln` steps (`build-test` and the Windows/macOS `builtin-docx` legs) pass `-m:1`, so the eleven test assemblies run one after another. A manual dispatch with `parallel_tests` replaces only that flag with `-m` through the workflow-level `OFDRW_SOLUTION_TEST_NODES` variable, which is substituted before any shell runs (bash or pwsh). Pushes, pull requests and default dispatches produce the unchanged `-m:1` command line. A trial always runs the full plan and never uploads a validation marker. Parallel tests are not the default: their safety is unproven. `Converter.Docx.Tests` starts LibreOffice, `Cli.Tests` can reach the Auto engine through `docx-to-pdf`, the PDF import tests can start `pdftoppm`, and several process tests are timing-sensitive (see "Startup cancellation measurements").

Input for the trial: every LibreOffice run in `src/` goes through `DocxToPdfConverter`. For ordinary installations (apt `/usr/bin/soffice`, the macOS cask, the Windows MSI) it isolates the user profile by default: each conversion passes `-env:UserInstallation` pointing at a `profile` directory inside a fresh owner-only temporary work directory, and `--outdir` is that same work directory. The process working directory is the shared LibreOffice program directory, which is only read. Profile isolation is off only for LibreOffice Portable or `SecureUserConfig` builds or when `IsolateUserProfile = false`; those conversions are serialized by a static in-process semaphore, which does not cover two test processes. CI does not use such builds. Concurrent LibreOffice processes with separate profiles have not been measured on hosted runners.

A trial must confirm that each platform still produces all eleven assembly TRX files with the same counts and skip reasons, that the LibreOffice-backed and process-timing tests pass across several runs, and the wall-time change. Only then should the default change.

## 2026-11-16 workload reduction

Historical measurements that motivated this change, collected from GitHub Actions records for 2026-11-01 to 2026-11-16 before it (not post-change results):

- 69 `main` push CI runs: 37 re-ran a tree already passed in its PR, and 12 more changed only documentation.
- 12 of 92 PR runs changed only documentation.
- A full CI run took about 2,702 job-seconds (median), plus 274 s for the separate SES workflow, of which about 208 s rebuilt and re-consumed the default SDK.
- The ubuntu no-Office leg of `builtin-docx` took about 140 s: a fresh runner, a 52 s build of `Ofdrw.Net.PlatformTests.slnf`, filtered `--no-build` test steps and a CLI run.

Changes: the `changes` planner and validation markers; SES as a candidate consumer inside the Linux package job, with the SES package packed from the tested build; the cross-OS LFS fixture cache warmed by `changes`; the ubuntu no-Office leg folded into `build-test`; and the `parallel_tests` trial switch (off by default).

Former ubuntu-leg steps not kept as separate steps, because none of the code they exercise depends on whether Word, LibreOffice or Poppler is present, and the same tests run in `build-test`'s full `Test` step against the same build (TRX in `test-results`):

| Former step | Why it is a duplicate |
| --- | --- |
| HarfBuzz native loading and prelayout (`TextPrelayoutTests`) | In-process shaping with pinned HarfBuzz; no external tool. |
| Image API (`ImageIoTests`, `ImageDocumentToolsIntegrationTests`, `ImageMainIntegrationTests`) | Image/OFD/PDF via ImageSharp, Docnet and the OFD→PDF writer; no `PdfToOfdConverter`, so no `pdftoppm` fallback. |
| Image CLI (`ImageCommandTests`) | Only `image-to-ofd` and `ofd-to-image`, same code paths as above. |
| CLI path safety (`Audit004RegressionTests`, `TextReplacementCommitNotificationTests`) | `ofd-to-pdf`, `reorder`, `replace-text` path and publication checks; no Office engine. Its former `native-path-tests-ubuntu-latest` artifact is no longer produced. |
| External process ownership (`ExternalProcessTests`, `Audit024*`) | Spawns shell-script and `dotnet` probe processes it creates itself; tool presence is irrelevant. |
| Owned file intent (`OfdOwnedFileIntentTests`) and replacement sessions (`TextReplacementFileSessionTests`, `TextReplacementOwnedSessionTests`) | File-system semantics only. |
| Font identity (`DocxFontIdentityTests` with `OFDRW_REQUIRE_TRUSTED_FONTS=1`) | Native DOCX→OFD→PDF; already included, with zero skips required, in `run-trusted-font-tests.py`. |
| Embedded bold and font identity (`EmbeddedBoldCompatibilityTests`, `FontIdentityRegressionTests`) | Native mode only, same environment as the full `Test` step. |

Deliberately not implemented:

- Caching the LibreOffice installation on Windows/macOS. Windows spends about 66 s in the MSI install and only about 13 s downloading, so a download cache buys little, and caching an installed prefix contradicts the cache policy above (downloads only, never a relocated installation).
- Parallel test assemblies by default; this needs the trial above.

None of this has run on GitHub-hosted runners yet; local validation covered `actionlint`, YAML parsing, `shellcheck`, the repository script tests and an end-to-end local run of the SES candidate mode. The first real runs must confirm:

- the planner finds markers with the job token: a PR run uploads `ci-validated-<tree>`, the following `main` push reports `reused`, all heavy jobs are skipped and `CI complete` passes; a documentation-only PR or push reports `docs-only` and uploads its own marker; a fork PR runs in full;
- the LFS cache: one save on the first full run, hits on Linux, Windows and macOS (cross-OS restore), warming on `main`, and no job failure when warming fails;
- `build-test`: no LibreOffice on a fresh `ubuntu-latest` image before the install step, both apt invocations, the no-Office results, and the changed job duration on the critical path;
- the SES package artifact, the SES consumer in `package-e2e`, the `ses-interop-evidence` archive, and that `build-e2e-harnesses.sh` still reports `Signatures.Appearance.E2E` as `SKIP`;
- the release gate on the next tag: `CI complete` alone for new commits, and the extra `ses-interop` check for an older tag in a manual recovery run;
- after the trial, whether parallel assemblies are stable and worthwhile.

## Evidence retention

Ordinary PR artifacts retain 14 days, main CI artifacts 30 days. The release workflow keeps its existing retention and candidate-bound Preview acceptance requirement. Before ordinary artifacts expire, promote any evidence required for long-term acceptance according to `docs/engineering/repository-storage.md`.

Large Linux evidence directories are uploaded as `.tar.zst` plus a JSON inventory. Every file, including all OFD/PDF/SVG documents and page images, is retained byte-for-byte. Long-window compression reuses repeated payloads across files. No visual assertions or pages are removed.

Download the artifact, verify the archive SHA-256 in the JSON inventory, then extract it with:

```sh
zstd -d evidence.tar.zst -o evidence.tar
tar -xf evidence.tar
```

The inventory lists relative paths, byte counts and SHA-256 for the original files. It also records original and archived byte counts for comparisons. The archive root is the original evidence directory name. Package candidates are uploaded separately for native Windows consumption, so consumers need not download Linux page evidence.

Tests that depend on restored NuGet metadata must run after solution restore, as in CI. Local conversion validation must still follow the repository's PNG and macOS Preview requirements; archive integrity alone is not visual acceptance.

The mixed-text consumer keeps seventeen bounded one-line samples, including
repeated-letter/GID occurrences and palindromic words. After
package consumption, Linux/Windows/macOS independently decode scalar-to-CID/GID
mappings and original outlines, compare every visible PDF contour to the actual
OFD positions, and associate independently parsed physical word boxes. Native
macOS also checks PDFKit whole selection and logical/reversed Find. These
headless checks do not replace final Preview inspection/copy/Find. Reader
Python dependencies use a separate pinned virtual environment, leaving the CI
font generator's pinned fontTools version and byte checks unchanged. Each
platform uploads its raw compatibility strings, tool/reader identities,
geometry receipt and input hashes alongside the existing package evidence.

## Startup cancellation measurements

The Unix socket protocol test measures receiver exit from the actual cancellation
request, keeping its one-second response limit and checking the exact caller token.
A dedicated sender waits for the test's invocation announcement, then requests
cancellation; this announcement is not an observation inside the production poll.
The TRX output retains invocation, request and exit ticks plus stopwatch frequency.
The separate 350 ms startup-deadline test still measures the startup wall clock.
Production polling, test parallelism and ThreadPool settings are unchanged.

The first main run at `b73e3ab9` failed the former total-time assertion, after
throwing the expected cancellation exception. That record did not separate timer
trigger time from receiver response, so its original delay remains unattributed.
An isolated child scheduling control demonstrated that a late timer can fail that
assertion while the unchanged receiver reacts promptly; it does not establish the
original runner's scheduling history. An ignore-cancellation mutation must still
fail with the wrong exception type/protocol deadline, rather than pass this test.
