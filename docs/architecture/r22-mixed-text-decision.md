# R22 mixed text: paragraph and PDF architecture decision

Status: **reference comparison complete; two-word product encoding unresolved**
on `4d9f003a43d0bfc8a1be0ccce298b40080329d75` (2026-10-30).
This record does not widen the accepted writer profile or select a replacement
PDF backend.

## 2026-10-30 reader acceptance decision

R22 不再要求 PdfPig 0.1.15 默认 `GetWords()` 和 `Page.Text` 充当 RTL 逻辑阅读顺序的必过判据；保留其原始结果作为版本化兼容性记录，保留独立的字符映射及几何检查。

双词输出仍须通过原文、字体、原 GID、实际位置、逻辑词搜索、独立几何验证以及真实 Preview 原始复制和 Find 门禁。替代读者及检查方式先验证、后冻结；在产品链路取得新证据之前，writer 不扩大支持范围。

PDF→OFD 对默认词提取的生产依赖作为独立缺陷处理，不将导出契约调整解释为反向导入已获支持。旧包、参考文件和诊断文件的验收范围保持不变。

This decision separates output correctness, versioned reader compatibility,
and the reverse-import defect. It changes no historical matrix result and does
not make the fpdf2 reference a product candidate. The proposed replacement
checks must be validated on the frozen positive and negative reference files
before their exact versions, APIs, inputs, and failure rules are frozen. The
[candidate validation protocol](../validation/r22-reader-replacement-candidate.md)
specifies the first bounded controls; it is not yet an accepted reader contract.

## Problem and fixed evidence

The public `TextMixedParagraphEngine` already returns logical UTF-16 source
intervals, line decisions, visual run positions, shaped original GIDs and
clusters for a bounded Latin/Hebrew LTR query. The current writer makes one
logical OFD TextObject per word or space and rejects adjacent/repeated Hebrew
words, edge-position Hebrew, wrapping, other fonts and sizes. It does so
because the PDF path's source-order semantic text and visibly positioned RTL
word outlines do not consistently copy and search in target readers.

The fixed counterexample corpus is in
[`r22-mixed-paragraph-validation.md`](../validation/r22-mixed-paragraph-validation.md).
The original ten-case matrix used source commit
`27abdf5779eb8768d3d6d96f9dac7c9e66bb5a79` plus dirty patch SHA-256
`db567480bc1e1962d028720dfad9f9ba5abec931fd0193e1a37a490e7622335c`.
Its face was the derived DejaVu Sans 2.37 fixture, SHA-256
`d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc`;
the [font manifest](../../e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/upstream-manifest.json)
records upstream and derivation identity. The two-word source is
`abc שלום מקום def`. The immutable product PDF used by the A–D diagnostics
has SHA-256 `73ea85db5bc36100b2e3098d70034d8a600266ca3cdcaad315385292df051864`.
The A–D, character E–H, word ActualText, logical tags and single-`TJ`
diagnostics all left at least one required target API or geometry criterion
failing. Product Preview also reproduced adjacent/repeated-word copy and Find
failures. These are counterexamples to those encodings, not proof that no PDF
encoding can work.

The merged `main` validation of `4d9f003a` accepts the current bounded
single-word profile only. The 12-package manifest SHA-256 is
`2bbe2b2d9638be438ae495947855b38a756bfa8827883b08ef4f25370f223c7d`.
Those immutable packages are a baseline, not the candidate for an expanded
two-word contract.

## Reference implementations and comparison boundary

[Pango's pipeline](https://docs.gtk.org/Pango/pango_rendering.html) separates
itemization by direction/font, shaping, line breaking, justification and
rendering. [PangoLayout](https://docs.gtk.org/Pango/class.Layout.html) owns
paragraph line decisions and maps logical indices to positions. Pango uses
HarfBuzz for shaping and FriBidi for bidi in its free-software stack. This is a
reference for paragraph geometry and run boundaries; a Pango/Cairo PDF cannot
by itself establish the OFD/PDF reader contract.

[Cairo's low-level glyph API](https://www.cairographics.org/manual/cairo-text.html)
accepts positioned glyphs and a UTF-8-to-glyph cluster map through
`cairo_show_text_glyphs`; it does not calculate the paragraph layout. Its
backward cluster flag can associate logical text with reverse glyph order.
The target PDF readers still determine their own selection and Find behavior,
so the generated PDF must be measured through the same independent matrix.

[fpdf2's text shaping](https://py-pdf.github.io/fpdf2/TextShaping.html)
combines optional `uharfbuzz` shaping with a Unicode bidi pass when enabled;
its flow APIs make it an independent PDF-generator comparison. It is a
separate Python implementation, not a source of OFD mapping semantics. Compare
the fixed-position baseline and paragraph-flow output separately. Lock the
actual fpdf2/uharfbuzz versions and use the exact same font bytes.
The [fpdf2 method table](https://py-pdf.github.io/fpdf2/Text.html) says
`.text()` does **not** support shaping, while `.cell()` and `.multi_cell()`
do. The former is retained as a negative control only.

| Experiment | Same inputs | What it can answer |
| --- | --- | --- |
| Fixed-position PDF | Exact face, `abc שלום מקום def`, repeated word and isolated words; same 4 mm em | Whether another PDF encoder's glyph/cluster representation improves text, Find and rectangles. A separate fpdf2 run uses the product's recorded per-run X positions and logical draw order. |
| Paragraph layout | Same source/face and declared widths; fixed LTR base direction | Whether Pango and fpdf2 choose similar run order, line breaks and glyph positions. Different layout is recorded, not silently scored as a PDF encoding failure. |

Record source text/codepoints, font and input hashes, exact versions, generated
PDF hashes, embedded fonts, content and ToUnicode/ActualText/tag structure,
raw extraction for every fixed API, all word rectangles, PDFKit whole
selection and both search directions, and actual Preview copy/Find/geometry.
Visual comparison must use each reference's declared layout expectation.
For a product candidate also record OFD write/readback and actual SVG/PNG/PDF
pages. No diagnostic rewrite or text normalization counts as product success.

### 2026-10-30 reference result

The local receipt is under ignored `artifacts/r22-next-phase/reference-20261029/`:
`pango-manifest.json`, `fpdf2-manifest.json`, `analysis.json`,
`pdfpig.json`, `pdfkit.json`, `offline-visual-receipt.json` and each
hashed PDF/PNG. The exact source text and font hash are in both manifests.
Pango 1.58.2, Cairo 1.18.4, HarfBuzz 14.4.0 and Fontconfig 2.18.3 used
`PANGOCAIRO_BACKEND=fc` plus an isolated Fontconfig file; `pdffonts`
confirmed embedded subsets named `R22DejaVuSans-Regular`. An initial
CoreText run selected Helvetica/Lucida Grande despite `fc-match`; it was
discarded and the six Pango cases were regenerated with the exact face.
fpdf2 2.8.9 used uharfbuzz 0.57.0. Poppler was 26.09.0, pypdf 6.21.2,
PdfPig 0.1.15 and macOS PDFKit 1451.5.3. PDFKit probes were API-only; no
Preview or clipboard access occurred.

| Reference (one page each) | PDFKit whole selection / logical Find | PdfPig 0.1.15 default and ActualText mode | Page image |
| --- | --- | --- | --- |
| Pango/Cairo two-word, PDF SHA-256 `c20d0d39456727ea7f967b679120cb6ffc0f52bb89cfc7816b746f48d7a99513` | Selection `םולש abc\ndef םוקמ`; neither logical Hebrew word found, reversed queries found | Page text `abc שלום מקוםdef`; word tokens are reversed Hebrew | RTL phrase and glyphs visually correct |
| fpdf2 shaped `.cell()` two-word, PDF SHA-256 `7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8` | Exact source; each Hebrew word found once at distinct positive rectangles | Page text `abc םוקמ םולש def`; logical words absent from `GetWords()` | RTL phrase and glyphs visually correct |
| fpdf2 shaped `.multi_cell()`, zero cell margin | Both logical words found; selection still changes cross-line order | Logical words absent from `GetWords()` | Declared between-word/after-phrase line breaks visually correct |
| fpdf2 shaped cells at product run X positions in logical source order | Both words found at the recorded positions; whole selection becomes `שלום abc\ndef מקום` | Both Hebrew word tokens reversed; raw text is not the source | Positions and glyphs visually match the reference |
| fpdf2 unshaped `.text()` control | Logical words not found | Some raw strings happen to equal source | Hebrew glyph order is wrong; invalid as a candidate |
| Three PDF-only operator-order diagnostics from the shaped fpdf2 PDF | None has exact selection and both logical Finds | Logical character order recovers page text in one variant, but its word tokens remain reversed | All three rasters are byte-identical to the visually reviewed shaped fpdf2 page |
| PDF-only Type 3 visible word-glyph diagnostic, PDF SHA-256 `65e1b4aba679d163a4e44473023167fa5885fa040f3c91902bf43204f1df2df1` | Both logical Finds once at positive distinct rectangles; whole selection `שלום abc\ndef מקום` is not source | Exact source in `Page.Text` for default/ActualText modes, each logical word once in default `GetWords()` | Same visible arrangement at page scale; 144 dpi raster is not pixel-identical (645 changed pixels confined to Hebrew outline region). One PDF character per word loses individual character geometry. |

All 23 PDF hashes, exact extracted strings, Find rectangles, PDF structures,
embedded subset hashes and 23 PNG hashes
remain in the receipts. The Pango and fpdf2 shaped outputs have ToUnicode;
none has a tag tree, and only one Pango wrap contains an ActualText token.
The visual reviewer inspected all 20 independent reference pages with matching
hashes; the three diagnostic rasters match the shaped two-word reference
byte-for-byte. These results show that another mature generator can move the
tradeoff between readers; neither tested output passes the combined
text/word/geometry criteria.
No conclusion transfers from PDFKit API to actual Preview copy or Find.
The pinned PdfPig source trace and a separate, falsifiable visible-word-glyph
encoding hypothesis are recorded in
[`r22-pdfpig-reader-root-cause.md`](../validation/r22-pdfpig-reader-root-cause.md).
That Type 3 PDF-only diagnostic has now been run and falsified by PDFKit whole
selection; its structure, one-page PNG, reader results and source hashes are
in the [curated LFS evidence archive](../evidence/r22-next-phase/README.md)
and ignored `type3-word-diagnostic-receipt.json`. The repeated-word variant,
candidate Preview and product writer path were not run after the first
decisive API failure.
The nine-file exact-hash Preview manifest is in ignored
`artifacts/r22-next-phase/reference-20261029/preview-input-manifest.json`.
Terra High inspected all nine one-page PDFs under three exclusive GUI grants.
The immutable receipts are `preview-partial-receipt-20261030.json`,
`preview-continuation-partial-receipt-20261030.json` and
`preview-remaining-six-receipt-20261030.json`. The first two grants stopped
after native state reads returned `noWindowsAvailable` immediately after the
last task window was closed; the final grant kept an owned PDF open while
checking the other six, then closed it and observed the normal Open chooser.

| Actual Preview finding | Scope |
| --- | --- |
| Shaped fpdf2 two-word and repeated-word | Correct visible phrase; logical `שלום`/`מקום` Find highlighted the expected words. Repeated `שלום` had two distinct highlights. Reversed queries had no results. |
| Shaped fpdf2 between-word wrap | The `מקום` query twice showed “Found on 1 page” but the result row said “Not found,” navigation was disabled and no target highlight appeared. PDFKit had reported one positive match. |
| Shaped fpdf2 after-phrase wrap | Both logical Hebrew words had one visible highlight; reversed queries had none. |
| Pango two-word, repeated and both wraps | Logical Hebrew queries returned no results. Reversed-word queries sometimes found only a partial glyph highlight; repeated `םולש` showed one result rather than PDFKit's two. Reversed `םוקמ` in the two-word and between-wrap cases showed a found-page banner without an actionable result or highlight. |
| Unshaped fpdf2 `.text()` control | Hebrew appearance was reversed; logical queries failed and reversed queries found the visible words. |

Preview accessibility page text and screenshot-space highlight estimates are
recorded separately from PDFKit API results. The permitted GUI API did not
expose raw clipboard Unicode, and the exact PNGs were not independently
compared inside Preview. Thus actual copy-codepoint and PNG-comparison gates
remain unverified. These reference results neither admit the product writer
nor demonstrate full R22 visual acceptance.

## Decision

1. Keep the original logical paragraph and source spans as the authority for
   OFD and extraction. Retain visual placement, script, direction, font
   identity, original GIDs, cluster source ranges, advances and offsets as
   distinct fields. A page's X coordinates or PDF stream order cannot
   reconstruct logical order reliably.
2. Put bidi analysis, itemization, legal break opportunities, line assignment
   and final reshaping above the run shaper. Preserve the current bounded
   query's behavior until a new profile has evidence. HarfBuzz remains a run
   shaper; it does not decide whole-paragraph line breaks or PDF semantics.
3. Keep the OFD writer's fail-closed source/GID checks. A two-word output may
   be admitted only by a general rule over run structure, source intervals and
   validated clusters. Do not add a specific font hash, font size, word string
   or extra two-word special case to make a test pass.
4. Treat positioned visible glyphs and semantic Unicode as one PDF contract
   with two independently checked representations. Compare cluster-aware text
   operators, ToUnicode, word-scoped ActualText and structured reading order
   through normal production encoding paths. The existing A–D, E–H, tags and
   single-`TJ` failures remain negative controls. The replacement reader and
   independent character/word geometry checks must be demonstrated and frozen
   before they can judge a product PDF. Preserve raw PdfPig 0.1.15 `Page.Text`
   and default `GetWords()` as compatibility observations, not RTL logical-order
   acceptance gates. A product PDF still needs exact source text, logical word
   search, distinct positioned rectangles, original font/GID and visual proof,
   including raw Preview clipboard and Find. Do not admit a two-word writer
   output or select Pango/Cairo/fpdf2 as the product backend from reference
   evidence alone.
5. Keep font subsetting as a separate bounded decision. A PDF encoding
   prototype may use the complete fixture font for diagnosis, but production
   OFD/package acceptance still requires original-GID preservation, embedded
   subset identity and the existing font tests.
6. Keep native loading explicit and measurable. Any new native paragraph
   engine would need per-platform packaged assets, selected runtime DLL/native
   byte identities, actual load and isolation tests, plus cross-platform
   package consumers. This architecture record does not adopt Pango as a
   runtime dependency or replace the current .NET/PDF path.

## Stages and gates

1. **Complete:** Freeze the counterexample manifest and run both independent
   reference comparisons under the exclusive generator/build grant. Preserve
   failures, versions and hashes. The operator-order diagnostic remains
   diagnostic-only and supplies no passing product hypothesis.
2. Validate and freeze replacement reader and independent character/word
   geometry checks against the frozen positive and negative reference files.
   Then find a PDF encoding that passes the fixed two-word source, both logical
   word Finds, distinct word rectangles and unchanged visual page under that
   contract. Implement one generalizable two-word path from public API through
   OFD write/read to PDF. Reject other unverified structures before writing.
   Measure source continuity, GID/cluster identity, OFD readback, PDF
   structure, exact unmodified text, word rectangles and visual pages.
3. Run focused tests, then shared full regression and exact isolated package
   consumer gates when the candidate touches shared layout/font/converter
   code. Inspect generated SVG and PNG pages; obtain a separate GUI grant
   for Terra High to inspect the exact final PDFs in Preview. Keep raw
   clipboard evidence distinct from PDFKit and AX Find-field observations.
4. For each final PR head, check current `origin/main`, PR48 ancestry,
   storage/LFS, complete CI, independent SES and valid review closure before
   the authorized normal merge. Re-freeze and repeat affected gates after a
   repair merge.

Unknown imported mappings (R02), general bidi/Arabic/Indic and R20 remain
outside this two-word increment. No release or public package publication is
part of this decision.
