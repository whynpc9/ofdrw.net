# R22 mixed text: paragraph and PDF architecture decision

Status: **reference comparison complete; two-word product encoding unresolved**
on `4d9f003a43d0bfc8a1be0ccce298b40080329d75` (2026-10-30).
This record does not widen the accepted writer profile or select a replacement
PDF backend.

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
   single-`TJ` failures remain negative controls. A product PDF must satisfy
   the declared logical text, independent word search and distinct positioned
   rectangles in PdfPig, PDFKit and Preview, alongside appearance. The present
   reference comparison does not identify such an encoding. Do not admit a
   two-word writer output or select Pango/Cairo/fpdf2 as the product backend
   from this evidence.
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
2. Find a PDF encoding that passes the fixed two-word source, both logical
   word Finds, distinct word rectangles and unchanged visual page in the
   pinned readers. Then implement one generalizable two-word path from public
   API through OFD write/read to PDF. Reject other unverified structures before writing.
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
