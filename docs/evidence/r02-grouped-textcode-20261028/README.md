# R02 grouped TextCode evidence

This optional evidence bundle supports [the bounded validation report](../../validation/r02-grouped-textcode-validation.md). It records actual inputs and outputs; it is not a general unknown-document interoperability corpus.

Product source: `e45379a9d04fc0e9b82ad67bc5dd2b6aaa479081`, incorporating PR #69/main `0d3c3b5c`. Candidate packages: `0.1.0-r02.20261028.18`. Later report/tooling commits leave the tested product source unchanged; final PR CI is separately bound to its own head.

The bundle includes:

- Directed source-writer and public pinned Go-writer OFD originals, including the refused Go source-count-delta specimen.
- Actual isolated-package PDF/SVG/HTML outputs from macOS, WSL and native Windows, plus source combining/legacy characterizations and required multilingual native/default DOCX regression pages.
- Independent-reader PDF/SVG output, geometry and runtime identity receipts, the candidate manifest and tool/environment information.
- The actual PDF/SVG page images and native Preview inspection record. SVG images come from the SVG outlet, not PDF rasterization.

It excludes caches, full logs, build directories, package DLL copies, the candidate package feed and third-party source copies. Fonts embedded in the OFD/PDF are necessary input/output payloads; the original font fixture remains in the existing shared test directory. Its OFL license accompanies the evidence.

`inventory.json` lists 178 retained files with byte counts and SHA-256; `archive.json` identifies the 2,334,324-byte archive (about 2.23 MiB), SHA-256 `9a54aa461c411fe87e9323feb1c9a78c253b873b957b93ff808cfebd87fffc6e`. Byte-identical visual inputs/pages share one retained payload; the portable `visual-inputs.json` maps all 32 inputs and 34 inspected pages to those payloads. Original machine paths and the original manifest are retained as provenance, not required extraction destinations.

From the repository root, fetch only this optional directory, verify the archive checksum, then extract to a new ignored directory:

```sh
git lfs pull --include='docs/evidence/r02-grouped-textcode-20261028/**' --exclude=''
shasum -a 256 docs/evidence/r02-grouped-textcode-20261028/evidence.tar.zst
mkdir -p artifacts/r02-grouped-evidence
zstd -d docs/evidence/r02-grouped-textcode-20261028/evidence.tar.zst -o artifacts/r02-grouped-evidence/evidence.tar
tar -xf artifacts/r02-grouped-evidence/evidence.tar -C artifacts/r02-grouped-evidence
```

The archived `visual-review/report.md` and `receipt.json` separate 34 PNG inspections, 19 actual Preview pages, clipboard code points and uncertain/failed Find behavior. No general reader-search acceptance is claimed. The source validation recipes live in [`scripts/validation/r02-grouped`](../../../scripts/validation/r02-grouped/README.md).
