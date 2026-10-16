# 003 follow-up acceptance: BuiltIn DOCX PDF italic closed

The existing direct BuiltIn subtitle italic defect is repaired through the public graphics drawing interface. Functional, automatic rendering, PNG visual review and native macOS Preview gates all passed for the scoped samples below. This does not assert arbitrary Word fidelity or fix the unrelated source landscape page-setup behavior documented below.

## Identity and scope

- Frozen base: `3247c76f95c9a3312cc26144e24c4984767af26f`.
- Validated/generated source HEAD: `432077e938e8c9d2fdf7930d07087c7b82e4eb22`; later evidence commits change documentation/artifacts only.
- Branch: `codex/audit-followup-003`.
- Managed worktree: `/Users/wanghongyi/.codex/worktrees/audit-followup-003/ofdrw.net`.
- Main source: existing deterministic `e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx`; its bytes and fonts were not changed to make the fix pass. Supplemental `visual/italic-boundaries.docx` supplies real-italic, host-regular fallback and repeated mixed styles.
- PdfSharpCore and MigraDoc stay at 1.3.67. No vendored/modified package, global NuGet-cache patch, private reflection, false italic metadata, text-string matching or repeated text painting.
- Original transaction production file and ten tests have zero diff from the frozen base.

## Repair and small experiments

The frozen archive source is verified against git show in `baseline-source-provenance.json`; source attribution does not rely on build metadata inherited from its parent directory. The baseline public DOCX regression failed at the expected style assertion (`glyph3 skew=0`), with baseline/adjacent regular controls running. The same actual PDF regression passed through `IXGraphicsRenderer.DrawString`, before PDF glyph encoding. The interface provides the exact occurrence, XFont, brush and baseline; therefore no CID/GID occurrence association is guessed after encoding.

For each selected italic face, the conversion snapshots the physical sfnt through the existing content-addressed registry, inspects actual style metadata, and binds that physical XFont family to its real italic flag. A requested italic drawn with an upright face receives one baseline-preserving shear inside Save/Restore; a true italic face passes through once. Unknown physical identities, unsupported anchors/containers/clipping are rejected before candidate commit. Ordinary drawing operations are forwarded; documents without italic retain original RenderPages. Bold simulation is outside this follow-up.

The pinned upstream [XGlyphTypeface construction](https://github.com/ststeiger/PdfSharpCore/blob/d6ac8b092129a4f797365bbbf3eea2723d6c3ebd/PdfSharpCore/Drawing/XGlyphTypeface.cs#L76-L110) discards resolver simulation. A flags-only fix is insufficient. A dependency fork would require new package provenance/distribution and affect all PDF consumers; the small exact-drawing experiment proved a narrower route with the existing public API. A post-encoding font-name/text-string replacement was rejected because it cannot prove styled occurrences after font-table coalescing.

## Functional gates

| Gate | Actual result | Evidence |
| --- | --- | --- |
| Initial baseline defect |1 expected style failure, no setup failure|baseline-tests/baseline-italic.trx|
| Style/drawing plus original transactions|17/17; zero skip|breadth-tests/breadth.trx; final full suite also executes them|
| Final full solution|2819/2819;11 projects; zero failed/skipped|test-summary-final.json; full-tests-final/*.trx|
| DOCX/PDF projects|285/285 and514/514|same final TRX set|
| Original transactions|10/10, unchanged|final DOCX TRX; frozen-base diff empty|
| Trusted font geometry/replacement|65/65; zero skipped|trusted-font-tests-final/trusted-fonts.trx|
| Python tooling|20/20|python-tests-final.log|
| API docs/dual netstandard targets|passed;272 historical diagnostics, no new debt|api-docs-final.log|
| Local11-package consumption|passed, isolated consumer cache; version 0.0.0-audit003final|package-output-final/package-manifest.json; package-e2e-final.log|
| Packed exact DLL style/transaction checks|17/17; test runtime DLL byte-equal to package entry|packed-style-runtime.json; packed-style-tests/packed-style.trx|
| Final model artifact export|2/2 controls|final-model-tests/final-model-artifacts.trx|

Final command arrays and exit codes are in `validation-commands-final.json`; runtime/generator binary hashes in `generator-binary-identity.json`; environment in `environment.json` and `dotnet-info.txt`. All dotnet commands used writable CLI home and required single-node/build-server flags. Full/docs/package gates used the shared validation lease and released it before GUI work.

Seven new tests cover exact skew ratio (guards against a future double simulation), adjacent upright text, repeated identical words with different style, mixed wrapping, unchanged glyph/metric/style tables, host false flags, genuine authored italic outlines without another shear, rejection of unknown identity/unsupported anchor, forwarded image bounds, table paths/fill, headers and PAGE/NUMPAGES, actual portrait/landscape page dimensions, rotation and unchanged glyph baselines.

The independent pypdf proof (`pdf-style-proof.json`) binds the actual boundary PDF font resources to physical SHA names: true Arial Italic has added CM/TM skew0; regular fallback has CM skew0.3420201; adjacent regular has skew0. The genuine source face SHA is `ce1d2f1ab89db45f9796100eee960f5702a40e84c225c2b48c3ec3e81d153f98`; fallback is original Arial Unicode SHA `876af2cd4854644e7f3e7feb2f688997fdb3343c6df6693611209c9dfb47ccec`. Original glyph/metric/style tables are unchanged; name rewriting is the registry's existing identity isolation. Supplemental Arial font files are task-local links to the installed system fonts, not redistributed source font files.

## Text and rendering

- Explicit Native and default OFD contain exactly the original202 characters, including source whitespace, once in TextCode order. Native selected page 2 has exactly its source text.
- All primary full PDFs contain all189 non-whitespace source characters exactly once in raw content order; the supplemental boundary source and selected page text also match. `pdftotext -raw` avoids table-column reading-order changes; this is not a claim of byte-identical PDF whitespace extraction.
- Native/default page 1 and page 2 PNGs are byte-identical. Direct BuiltIn page 2 is byte-identical to the same-environment frozen-base rendering. Its page 1 changes the previously upright subtitle to italic.
-18 page PNGs were generated at144DPI including baseline/no-italic controls. All 12 target pages below were visually reviewed. Final regeneration at the validated source produced PNG bytes identical to the initially reviewed renderings (`final-png-equivalence.json`).
- Main source has no image/header/footer; the mixed graphics control supplies an actual blue PNG, table shading/borders, header and visible page numbering. Prepared orientation control uses authored block glyphs from narrow.ttf; these are intentional symbols, not missing reading-font glyphs.
- Direct PDF size changed from97038 to160909 bytes (+63871, about 65.8%) because the immutable styled physical face has a separate PDF font subset. Text is still present once. Native/default outputs are unaffected.

## Native Preview acceptance

Actual chain: `DOCX → explicit native/default OFD → PDF → macOS Preview`, plus the direct BuiltIn and focused graphics PDFs. Fresh final PDF paths were opened after final generation; no old candidate/cache was substituted. Date 2026-10-16 Asia/Shanghai, macOS 26.6.2 arm64, Preview 11.0(1113.5.3).

All 12 target pages were checked for scoped italic/upright boundaries, text completeness/legibility, line wrapping, baselines/alignment, table fill/borders, page content/selection, clipping/overlap/ghosting/blank pages, and applicable images/headers/footers/page numbers. Native/default versus direct retains the engine's existing layout/style differences; this report does not assert equivalence between those engines.

| PDF | Pages actually viewed | Result |
| --- | --- | --- |
| `visual/native2/native2.pdf` |1,2|passed scoped checks|
| `visual/default2/default2.pdf` |1,2|passed scoped checks|
| `visual/directBuiltIn2/directBuiltIn2.pdf` |1,2|passed scoped checks|
| `visual/selected1/selected1.pdf` |1|passed scoped checks|
| `visual/italic-boundaries/italic-boundaries.pdf` |1|passed scoped checks|
| `visual/mixed-sections/mixed-sections.pdf` |1,2|passed scoped checks|
| `visual/mixed-sections/mixed-orientations.pdf` |1,2|passed scoped checks|

`preview/pages.json` and individual `.ax.json` files (losslessly encoded raw AX text) record the actual complete file URL and page indicator for every page. The 12 `.jpg` files are original JPEG capture bytes, with SHA256 and byte count; no recompression/resizing was used. `preview/pdf-before.json` and `pdf-after.json` prove all seven final PDF SHA256 values stayed unchanged. Only this task's PDF tabs were closed; `remaining-windows.ax.json` confirms none of their titles remains and other documents remain open. GUI lease was released.

## Existing unrelated source page-setup finding

Original BuiltIn ApplyPageSetup supplies source landscape width/height and Landscape together; MigraDoc expects orientation-neutral width/height and swaps them. A420x300 source model therefore produces300x420 through both original RenderPages and this adapter. PDFs, pdfinfo and hashes from the failed exploratory control are retained in `legacy-orientation-evidence/`. This source page-setup behavior predates and is unchanged by the italic repair; it is not claimed passed or fixed. The independent prepared MigraDoc control proves the actual new drawing seam preserves300x420 portrait and420x300 landscape PageInfo and anchors. A separate source-page-setup follow-up is needed; no expansion was made here.

## Handoff

Scoped003 italic issue and its required functional/visual gates are closed. No outstanding scoped gate remains. All changes are local; no push, PR, target/main merge, tag or publication was performed. Coordinator owns final multi-task integration and its combined gates.
