# CI performance and evidence

CI keeps the existing test filters, operating-system matrix, package-consumer isolation, rendering comparisons, and release Preview gate. There is no path-based test selection.

## Build and caches

- The public API documentation checker owns the full Release rebuild. A preceding ordinary build is unnecessary. Its CS1591 baseline and failure handling remain unchanged.
- Linux `build-test` runs the full Release solution and mandatory trusted-font suite. The platform matrix builds the full solution and runs those same suites on Windows and macOS; Linux retains the smaller `Ofdrw.Net.PlatformTests.slnf` build for its additional platform filters. All existing platform test filters subsequently use `--no-build --no-restore`.
- Document-tools E2E and its CLI use Release consistently.
- Only third-party build dependencies are cached. Consumers retain a fresh NuGet directory and source mapping that binds `Ofdrw.Net.*` to the exact candidate feed.
- CI fonts are cached after static instancing. Both the generated font and license must match source-controlled SHA-256 values on every hit. Invalid cache entries regenerate from the pinned source; source or generated-byte mismatches fail.
- Linux caches downloaded apt packages, but still refreshes repository metadata and installs dependencies normally. Windows caches micromamba package downloads, not a relocated installation prefix. These caches reduce downloads; they do not eliminate all solver, repository, or network latency.
- Full Windows/macOS solution tests install LibreOffice and Poppler before running; the PDF import and DOCX conversion tests require those external tools. The pinned Windows Poppler bootstrap is hash-checked. The package visual jobs install their own renderer tools on separate runners.
- The package visual corpus embeds the same verified Noto CJK font bytes on all three platforms. Its four macOS golden pages were aligned with the already reviewed Linux/Windows bytes after the CI runner and a separate local macOS capture produced identical page hashes; earlier host-font pages remain recoverable from Git history.
- New PR commits cancel older runs of the same workflow and PR. Main and release runs are not cancelled by this policy.

## One candidate, parallel consumers

`build-test` stamps its API-documentation rebuild with `0.1.0-preview.8`, tests it, and invokes `run-converter-package-e2e.sh 0.1.0-preview.8 --pack-only --no-build`. Pack-only requires an empty feed. The twelve default packages are validated and uploaded once, together with their SHA-256 manifest.

Linux, native Windows, and native macOS package jobs depend directly on `build-test` and download that one artifact in parallel. Linux restores source metadata for license validation but does not rebuild the full source solution or repack the candidate. Each consumer verifies the existing manifest before and after use and retains a fresh dependency cache with package source mapping. The Windows and macOS consumers also compare each selected runtime DLL against its nupkg payload, isolated cache, and copied output. Optional Skia/vector gates still consume the same defaults and separately build their optional adapters. The source-based document and graphics probes retain their own required builds.

The platform matrix runs independently. `CI complete` succeeds only when every source, platform, Linux/Windows/macOS consumer, and password-crypto job succeeds; failures, cancellation, and skipped dependencies fail this aggregate gate. The separate SES interoperability workflow remains required for task acceptance.

The shell script's default build-and-consume behavior remains available for standalone and SES runs. `--consume-only` supports existing release feeds, while `--pack-only --no-build` requires a caller-provided Release build stamped with the requested version. The package validator checks package metadata, dependency versions, payload requirements, and manifest hashes.

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
