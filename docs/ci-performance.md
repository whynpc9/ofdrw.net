# CI performance and evidence

CI keeps three-platform full-solution coverage, Linux no-Office test filters, package-consumer isolation, rendering comparisons, and the release Preview gate. There is no path-based test selection.

## Build and caches

- The public API documentation checker owns the full Release rebuild. A preceding ordinary build is unnecessary. Its CS1591 baseline and failure handling remain unchanged.
- Linux `build-test` runs the full Release solution and mandatory trusted-font suite. The platform matrix builds the full solution and runs those same suites on Windows and macOS; Linux retains the smaller `Ofdrw.Net.PlatformTests.slnf` build for its additional no-Office filters. Only Linux reruns the filtered tests with `--no-build --no-restore`; Windows/macOS retain their results and native-path skip reasons in the full-suite TRX artifacts.
- Document-tools E2E and its CLI use Release consistently.
- Only third-party build dependencies are cached. Consumers retain a fresh NuGet directory and source mapping that binds `Ofdrw.Net.*` to the exact candidate feed.
- CI fonts are cached after static instancing. Both the generated font and license must match source-controlled SHA-256 values on every hit. Invalid cache entries regenerate from the pinned source; source or generated-byte mismatches fail.
- Linux caches downloaded apt packages, but still refreshes repository metadata and installs dependencies normally. Windows caches micromamba package downloads, not a relocated installation prefix. These caches reduce downloads; they do not eliminate all solver, repository, or network latency.
- Windows download cache keys include architecture and pinned bootstrap/tool versions. Bump the matching key when changing those versions; unrelated edits to `ci.yml` no longer invalidate the downloads. Candidate nupkg/snupkg files use artifact compression level zero because their payloads are already compressed.
- Full Windows/macOS solution tests install LibreOffice and Poppler before running; the PDF import and DOCX conversion tests require those external tools. The pinned Windows Poppler bootstrap is hash-checked. The package visual jobs install their own renderer tools on separate runners.
- The package visual corpus embeds the same verified Noto CJK font bytes on all three platforms. Its four macOS golden pages were aligned with the already reviewed Linux/Windows bytes after the CI runner and a separate local macOS capture produced identical page hashes; earlier host-font pages remain recoverable from Git history.
- New PR commits cancel older runs of the same workflow and PR. Main and release runs are not cancelled by this policy.

## One candidate, parallel consumers

`build-test` stamps its API-documentation rebuild with `0.1.0-preview.8`, tests it, and invokes `run-converter-package-e2e.sh 0.1.0-preview.8 --pack-only --no-build`. Pack-only requires an empty feed. The twelve default packages are validated and uploaded once, together with their SHA-256 manifest.

Linux, native Windows, and native macOS package jobs depend directly on `build-test` and download that one artifact in parallel. Linux restores source metadata for license validation but does not rebuild the full source solution or repack the candidate. Each consumer verifies the existing manifest before and after use and retains a fresh dependency cache with package source mapping. The Windows and macOS consumers also compare each selected runtime DLL against its nupkg payload, isolated cache, and copied output. Optional Skia/vector gates still consume the same defaults and separately build their optional adapters. The source-based document and graphics probes retain their own required builds.

The platform matrix runs independently. `CI complete` succeeds only when every source, platform, Linux/Windows/macOS consumer, and password-crypto job succeeds; failures, cancellation, and skipped dependencies fail this aggregate gate. The separate SES interoperability workflow remains required for task acceptance.

The shell script's default build-and-consume behavior remains available for standalone and SES runs. `--consume-only` supports existing release feeds, while `--pack-only --no-build` requires a caller-provided Release build stamped with the requested version. The package validator checks package metadata, dependency versions, payload requirements, and manifest hashes.

## 2026-11-01 duplicate-work reduction

Baseline: `b02162bef61abc17761349d891237fc2518d3989`. The latest 100 workflow records span 2026-10-30 01:11 to 2026-11-01 20:33 Asia/Shanghai: CI has 39 successes, 8 failures and 3 cancellations; SES has 49 successes and 1 cancellation. Step timings were collected for the latest 15 CI runs (2026-11-01 09:54–20:33), with medians below computed from the 14 successful runs. Failed/cancelled runs are not counted as performance wins.

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

## Evidence retention

Ordinary PR artifacts retain 14 days, main CI artifacts 30 days. The release workflow keeps its existing retention and candidate-bound Preview acceptance requirement. Before ordinary artifacts expire, promote any evidence required for long-term acceptance according to `docs/repository-storage.md`.

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
