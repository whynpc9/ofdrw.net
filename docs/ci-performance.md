# CI performance and evidence

CI keeps the existing test filters, operating-system matrix, package-consumer isolation, rendering comparisons, and release Preview gate. There is no path-based test selection.

## Build and caches

- The public API documentation checker owns the full Release rebuild. A preceding ordinary build is unnecessary. Its CS1591 baseline and failure handling remain unchanged.
- `Ofdrw.Net.PlatformTests.slnf` builds the six platform test projects and their dependencies once. All existing platform test filters subsequently use `--no-build --no-restore`.
- Document-tools E2E and its CLI use Release consistently.
- Only third-party build dependencies are cached. Consumers retain a fresh NuGet directory and source mapping that binds `Ofdrw.Net.*` to the exact candidate feed.
- CI fonts are cached after static instancing. Both the generated font and license must match source-controlled SHA-256 values on every hit. Invalid cache entries regenerate from the pinned source; source or generated-byte mismatches fail.
- Linux caches downloaded apt packages, but still refreshes repository metadata and installs dependencies normally. Windows caches micromamba package downloads, not a relocated installation prefix. These caches reduce downloads; they do not eliminate all solver, repository, or network latency.
- New PR commits cancel older runs of the same workflow and PR. Main and release runs are not cancelled by this policy.

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
