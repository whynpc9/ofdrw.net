# Pinned original-font query fixtures

The content-addressed ZIP contains exactly two unchanged OFL font payloads, two full licenses, `manifest.json` and `NOTICE.md`. `bundle.json` pins archive size/SHA and every member. Fonts remain OFL-1.1; see the readable license and notice copies here. The CJK font is the historical fixed Regular fixture instantiated at wght400 from the pinned Noto variable source, with original metadata and source/tool/recipe pins retained. It was not regenerated this round and is not described as an identical upstream-distributed static TTF.

Prepare offline with Python's standard library:

```text
python3 scripts/prepare-trusted-fonts.py --output-dir artifacts/trusted-fonts
```

The emitted JSON supplies explicit fixture paths and `OFDRW_REQUIRE_TRUSTED_FONTS=1`. Preparation validates the pinned spec, archive and central-directory count before ZIP member allocation, then exact paths, duplicates, regular-file kind, sizes, expanded total, compression ratio and streamed SHA. It stages beside the target, verifies, and atomically renames. Existing verified content is reused; corrupt content fails without overwriting. It never downloads, installs OS fonts, subsets or regenerates a TTF. An independent directory needs only the copied script, `bundle.json` and the ZIP passed via `--manifest`/`--bundle`; no sibling checkout path is used.

The mandatory original-font gate prepares and verifies the fonts, sets task-local .NET state, and runs all 65 current geometry, locator and replacement cases: `TrustedFontGeometryTests` (33), `KeywordLocatorRealFontTests` (26) and `TextReplacementRealFontTests` (6). It requires actual results with zero skips:

```text
python3 scripts/run-trusted-font-tests.py --results-directory artifacts/font-gate/run-1
```

Use a fresh results directory. Tests also resolve the prepared repository-local directory when explicit paths are absent. Explicit paths remain supported and are size/hash checked. `OFDRW_REQUIRE_TRUSTED_FONTS=1` keeps real tests enabled so missing or corrupt assets fail; an optional ungated run without prepared assets may skip them and is not font acceptance.

This repository-local portable font asset/tool gate does not itself configure CI workflows, publication lists, uploads or the default eleven-package-consumer gate. Production geometry still consumes original font bytes from the input document under its explicit trusted profile; these test assets never replace document resources.
