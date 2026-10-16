# PR23 retained validation evidence

The original 2026-10-16 evidence from `artifacts/followup-integration/`,
`artifacts/plan-003-followup/`, and `artifacts/pr23-review-fixes/` is retained
byte for byte in one optional LFS archive. This replaces individual tracked
runtime files, source snapshots, test logs and duplicate viewing products in
the default checkout. The existing acceptance reports in `plans/` remain the
entry points; their historical `artifacts/` paths are paths after extraction.
This storage change does not repeat or extend any functional or visual acceptance.

Fetch and restore the original paths from the repository root:

```sh
git lfs pull --include="docs/evidence/pr23-integration/**" --exclude=""
zstd -d -c --long=27 docs/evidence/pr23-integration/evidence.tar.zst | tar -xf -
```

Extract into a clean temporary directory if those `artifacts/` paths already
contain another run. `manifest.json` records every original relative path,
size and SHA-256, plus the archive size and SHA-256. Nothing was regenerated.
The [small coordinator summary](root-evidence-review.json) remains directly
readable. Original reports are at:

- `artifacts/followup-integration/acceptance.md`
- `artifacts/followup-integration/root-evidence-review.json`
- `artifacts/followup-integration/preview/report.md`
- `artifacts/followup-integration/browser/report.md`
- `artifacts/plan-003-followup/acceptance.md`
- `artifacts/pr23-review-fixes/preview/report.md`
- `artifacts/pr23-review-fixes/browser/report.md`

Future runs stay in ignored `artifacts/`. Retain a small report of the source
commit, commands, results and limitations. Keep only the inputs, outputs and
visual records needed to reproduce or assess the change, in a named LFS bundle
with a manifest. Do not commit build output, package caches, frozen source trees,
repeated full test logs or repeated font payloads; source commits and the shared
test fixtures already preserve those inputs. Failed-run records are retained
only when needed to demonstrate a regression or explain an acceptance limitation.
