# R22 bounded mixed paragraph validation

This checkpoint covers the existing regular-font Latin/Hebrew LTR profile in
PR #71. It does not accept general bidi or expand the public input contract.
The local layout baseline is `27abdf5779eb8768d3d6d96f9dac7c9e66bb5a79`;
record the actual HEAD and dirty patch SHA-256 for every subsequent run.

## Separate contracts and evidence

The query accepts 3–128 BMP characters, ASCII Latin words, Hebrew letters and
single ASCII word spaces, with both scripts present. The typed writer is
narrower: each run has 1–8 characters, Hebrew words have 2–8 letters with one
distinct nominal GID per letter, and Latin has one-to-one glyphs or a complete
2/3-character ligature. Repeating a valid Hebrew word in two runs is different
from repeating a GID inside one word. Query success is not output acceptance.

| Check | Required observation |
|---|---|
| Layout | Logical source intervals survive; expected per-line visual run coordinates, shaped glyph order, trailing spaces and fit are correct. |
| OFD | Readback retains the exact logical source, original GIDs and font association. |
| PDF structure | Record content-order ToUnicode decoding, scoped ActualText bytes/nesting, and any tag structure. Do not call stream order the sole reading-order oracle. |
| Fixed extractors | Save exact unmodified Unicode and versions; run PdfPig 0.1.15 with UseActualText false and true. Save word rectangles separately. |
| PDFKit API | Save whole-selection text, normal/reversed word searches, counts, pages and every rectangle. |
| Preview UI | On the exact candidate, inspect appearance, copy raw Unicode and check each Find result/highlight. API results and screenshots cannot replace clipboard codepoints. |
| Visuals | Compare each diagnostic with its own correctly laid-out product baseline; inspect PDF/SVG renders and font/GID evidence separately. |

Only soft line breaks recorded by the producer may replace their original
separator space in the selection expectation. The matrix retains both the
exact source and that predeclared line-separated spelling. No sorting,
reversal, bidi-control removal or general whitespace normalization is used.

## Fixed corpus

`e2e/Ofdrw.Net.Text.HarfBuzz.E2E` generates these from public product APIs:

| Case | Purpose |
|---|---|
| `abc שלום def` wide and wrapped | Existing single-word baseline |
| `abc שלום מקום def` wide and exact fit | Consecutive RTL words with a shared first visual GID |
| `abc שלום שלום def` | Two equal words must remain two occurrences with distinct positions |
| `abc שלום def מקום ghi` | Latin separates two RTL sequences |
| `שלום abc` and `abc שלום` | Fixed LTR paragraph boundaries |
| Double-word sample wrapped between Hebrew words | Trailing source space belongs to the finished line and is not moved into the next RTL sequence |
| Double-word sample wrapped after the Hebrew phrase | Reorder the phrase before the line-end source space |

The existing two-page isolated-word package consumer remains a shared converter
regression probe. It calls `OfdPackageWriter` and `OfdToPdfConverter` directly;
the mixed-query requirement for both scripts does not classify that input as
unsupported. The existing `HebrewVisualWordPdfUsesOneLogicalSourceCid` and
`DifferentHebrewWordsSharingFirstVisualGidKeepSeparatePdfSemantics` tests admit
this generic explicit-glyph path, including isolated words on separate pages.
Keep its observed Preview failure as an open supported-path converter issue.

## Controlled semantic matrix

| Variant | PDF group order on each line | Semantic layer |
|---|---|---|
| A | Logical | Current per-word semantic CID |
| B | Logical | A plus per-word ActualText |
| C | Visual | Current per-word semantic CID |
| D | Visual | C plus per-word ActualText |

`scripts/validation/r22-mixed/make-matrix.py` takes the product output directory
and a separate diagnostic output directory. It requires pypdf **6.21.2** and
Poppler. It verifies input hashes, decodes each source group's ToUnicode,
moves complete balanced graphics groups, and applies only word-scoped
replacements. Every diagnostic must rasterize identically to its own source.
This is a diagnostic transformation, not a product implementation or acceptance.

The independent `consumer/Consumer.csproj` pins PdfPig **0.1.15** and accepts
`matrix.json output.json`. It records all results, including failures; it does
not normalize text or use the rewriter's assertion as the extraction oracle.
`pdfkit-probe.swift matrix.json output.json` reads the same hash-checked files
through native PDFKit without launching an application or using the clipboard.
`summarize.py matrix.json pdfpig.json pdfkit.json summary.json` joins the exact
file identities and separately compares word counts, distinct rectangles and
expected horizontal extent (0.1 PDF point tolerance for metric rounding).
Run .NET through the shared resource launcher with the required single-node
flags and a writable DOTNET_CLI_HOME.

Keep the phrase-wide ActualText probe as a negative control: at the previous
checkpoint PdfPig recovered the exact full sentence but merged the two words
into one rectangle covering only the left word. A single semantic CID for an
entire phrase is also not accepted. Per-word ActualText must be measured; it
does not by itself prove cross-word reading order.

If all four variants fail the combined criteria, the next isolated experiment
keeps visible outlines and original font metrics unchanged while testing a
character-granular semantic layer. Explicit tag-tree order is a separate option
if required. Neither option is selected without evidence.

## First matrix checkpoint (2026-10-28)

The expanded E2E produced ten product cases and passed all seven new boundary
assertions. Six focused converter tests matching `Hebrew` passed with no skips.
**Version correction completed:** the first A/B/C/D round inherited PDF 1.4
without a catalog version override. Marked-content ActualText requires PDF 1.5
([ISO 32000-2 clause 14.9.4 errata](https://pdf-issues.pdfa.org/32000-2-2020/clause14.html)).
The raw observations below remain valid for those exact files. The generator
now declares PDF 1.7 for all four variants. A controlled follow-up used the
same ten immutable product inputs and regenerated 40 probes in
`artifacts/r22-harfbuzz/actualtext-version-20261028/`. Each old/new pair has
the same page content stream, embedded font bytes and raster; none has a lower
catalog `/Version` override. Pinned PdfPig 0.1.15 and native PDFKit returned
identical raw text, word rectangles and Find results across both versions.
Version correction therefore did not resolve these readers' observed order
conflict. There is still no A/B/C/D variant that passes all ten cases; the
14 of 40 individually passing automated cells are unchanged. Actual Preview
remains a separate blocked check on any proposed product candidate.
The matrix produced 40 diagnostic PDFs with unchanged baseline pixels and
verified that every producer output retained its recorded hash. All 40 had
the expected separate word rectangles in PdfPig and PDFKit, including two
distinct matches for the repeated word. Text selection remained incompatible:

| Corpus subset | A/B | C/D |
|---|---|---|
| Single-word wide/wrapped, separated RTL phrases | Both APIs recover the declared text. | Same. |
| Two distinct adjacent words, wide/exact fit | PdfPig exact; PDFKit selection is `שלום abc\ndef מקום`. | PDFKit exact; PdfPig exposes `abc מקום שלום def`. |
| Repeated adjacent word | PDFKit selection is split/reordered. | Both APIs recover the declared text with distinct positions. |
| Leading/trailing RTL | PDFKit swaps the Latin and Hebrew words. | Same. |
| Wrap between/after RTL words | PDFKit differs from the predeclared line text. | PDFKit still differs; after-phrase wrapping also retains visual PdfPig word order. |

Per-word ActualText made no difference to the measured text, word rectangles,
or PDFKit searches between A/B or C/D in any of these cases. There is no variant
that passes the complete corpus. Full-page PDF rasters and enlarged PDF/SVG
content for all ten cases were inspected; word placement, line breaks and
glyph appearance matched the bounded expectations. This is offline PNG/SVG
review, not Preview acceptance. Preview remains BLOCKED by the existing GUI
connection issue; known earlier product copy/Find failures remain FAIL.

Receipts live under ignored `artifacts/r22-harfbuzz/mixed-matrix-20261028/`:
`source-v2-identity.json`, `product-v2/`, `matrix-v2/matrix.json`, `pdfpig.json`,
`pdfkit.json`, `summary.json`, focused TRX and `visual/`. The first diagnostic
round was superseded because its baseline renderer overwrote producer PNG
filenames. Its disposition is retained. The corrected script writes baseline
rasters only beneath the diagnostic directory and checks all producer hashes
before and after; the second round uses fresh separate directories.

## Character-granularity checkpoint (2026-10-28)

`make-character-matrix.py` reads the same ten product cases plus both pages of
the existing isolated-word OFD/PDF. It replaces only the single hidden semantic
CID with explicitly positioned character CIDs. E/F retain logical word-group
order, G/H use visual word-group order; E/G show characters in logical order,
F/H in visual order. The existing clip/opacity and visible paths remain.

All 48 diagnostics preserved original embedded font bytes, used-GID outlines,
`hmtx`, source PDF/OFD bytes and baseline pixels. The original ten product
rasters and independent renders of both isolated source pages also match.
No variant passed the combined API criteria. E keeps exact PdfPig page text,
but PDFKit fragments/reorders Hebrew characters and fails whole-word Find.
F/H improve PDFKit on selected cases and both isolated pages, while PdfPig
page text is in visual character order. Leading/trailing RTL and mixed-line
wrap cases still fail PDFKit selection. These files use no ActualText, so the
first matrix's version caveat does not explain these results.

| Character variant (12 cases) | pypdf exact raw source | PdfPig exact page text | PDFKit declared selection | PDFKit word geometry |
|---|---:|---:|---:|---:|
| E | 0 | 12 | 0 | 0 |
| F | 6 | 0 | 5 | 12 |
| G | 0 | 9 | 0 | 0 |
| H | 7 | 0 | 8 | 12 |

The pypdf column is raw exact equality, including its line breaks; it is not a
normalised reading-order score. These are consumer-specific measurements, not
a rule that every extractor's raw stream API is the sole semantic oracle.

PdfPig `GetWords()` returned visual-character tokens such as `םולש` even when
`page.Text` returned logical `שלום`; the expected-token geometry check therefore
fails. This is a measured consumer limitation, not proof that every rectangle
is physically misplaced. Keep raw text, token strings and rectangles separate.
Likewise the earlier phrase ActualText probe's merged PdfPig token remains a
failed consumer observation; it must not be confused with proving that every
target reader merges its Find rectangles. Actual Preview remains unverified.

Artifacts: `artifacts/r22-harfbuzz/character-matrix-20261028/` holds source
identity, 48 immutable outputs, PdfPig/PDFKit receipts and summary. No production
semantic change is selected. The first matrix's version declaration was
corrected and rerun above. A tagged-reading-order investigation
must validate the logical tree, MCIDs and ParentTree with an independent
tag-aware consumer; the [PDF Association RTL discussion](https://github.com/pdf-association/pdf-issues/issues/718)
distinguishes structure order, glyph order and reader heuristics. It does not
establish Apple PDFKit/Preview compatibility or waive their direct checks.

## Logical tag-tree checkpoint (2026-10-28)

The version-correct first matrix was followed by a seven-case T-L/T-V
diagnostic. Both variants preserve each product PDF's positioned word/space
groups and PDF 1.7; T-L keeps logical page stream order, while T-V writes the
groups in visual order. Both attach a logical Document/P/Span tree with each
word and space's own MCID, matching page `StructParents` and ParentTree links,
`en-US` LTR paragraph attributes and `he-IL` Hebrew spans. The generator
verified all links, font bytes and pixel identity. It does not claim PDF/UA.

Independent Poppler `pdfinfo -struct-text` recovered the exact logical source
in all 14 PDFs, including T-V's visual page stream order. Yet PDFKit's raw
document text, full selection and every Find count/rectangle were byte-for-byte
unchanged from the corresponding untagged A/C control. PdfPig 0.1.15 with
`UseActualText=false` and `true` likewise returned the same text, words and
rectangles as those controls. Three individual API cells met the combined
diagnostic criteria; no tag variant covers the seven-case set. The structure
tree establishes an independent logical reading order, but these target APIs
did not use it to repair copy/search. Preview remains unmeasured for the new
diagnostics because the existing GUI connection and window-ownership block
has not cleared.

The preserved source and receipts are in ignored
`artifacts/r22-harfbuzz/tagged-reading-order-20261028/`: original Poppler
`*.structure.txt`, parsed `poppler-structure-receipt.json`, MCID/ParentTree
checks in `matrix.json`, `tag-effect.json`, pinned PdfPig and native PDFKit
outputs. The observation narrows this tagged approach. It does not rule out
other standards-compliant PDF encodings or imply a product fix.

## Delivery decision after the controlled experiments

No tested PDF encoding meets the full mixed-paragraph corpus across pinned
PdfPig and PDFKit. Direct macOS Preview checks of the two-word A–D diagnostic
PDFs from `7c67057e` confirmed the tradeoff: A/B copied
`שלום abc\ndef מקום` and could not Find `מקום`; C/D copied the exact logical
sentence and independently highlighted both words, while their pinned PdfPig
text remained in visual word order. The exact paths, hashes, clipboard code
points and screenshots are in
`artifacts/r22-harfbuzz/preview-abcd-retry-20261029/receipt.md`.

After absorbing PR #72, `9f81fccf` regenerated ten product mixed cases.
Focused E2E and 125 PDF tests passed, all ten OFD/PNG and SVG pages were
visually checked, and source-order content-stream decoding and PdfPig text
were exact. PDFKit whole-document selection was exact in only two of ten cases.
Direct Preview of four hash-verified product PDFs reproduced the actual defect:
the single-word wide case copied and searched correctly, but adjacent distinct
words, repeated words and a line break between RTL words copied in the wrong
order; the distinct second word was not found, and the repeated word yielded
only one Find result. See
`artifacts/r22-harfbuzz/pr71-after-pr72-9f81fccf/preview-recovery/receipt.json`.

A separate PDF-only, single-`TJ` diagnostic preserved the source raster and
let pypdf extract the logical sentence, but PDFKit omitted `מקום` from both
selection and Find; PdfPig also attached a source space to `שלום` in its word
geometry. The pinned reader outputs are in
`artifacts/r22-harfbuzz/pr71-after-pr72-9f81fccf/one-tj-result.json`. This
variant is rejected as a product repair. PR #71 stays Draft with its writer/PDF
promise.

The bounded query and line-layout API was split into PR #72 and merged as
`187aa78ffca9ab4e81bdb0abb7d95b52d1e10337` on 2026-10-29. Its final
`325f1bce` head passed separate source, package, CI, review and four-page
Preview gates for the generated DOCX regression sample. Those checks establish
the query profile and the inherited shared converter regression only; they do
not validate the mixed writer or Hebrew PDF copy/Find behavior in this Draft PR.
The shared PDF bold repair PR #73 was also merged before PR #72 and does not
resolve the Hebrew reading-order tradeoff.

This branch has absorbed that main baseline. Its writer, explicit-glyph PDF
semantics and diagnostics remain here until one product-generated candidate
passes the reader and visual gates on the same final head. The previous
`210f696f` CI and package results and the unpushed diagnostic candidates do
not transfer to the new merged head.

## Completion gates

Each result is PASS, FAIL or BLOCKED with file hashes and tool versions. Current
Preview copy/Find failures remain FAIL; a computer-use timeout is BLOCKED.
Keep PR #71 draft until one product-generated candidate passes the required
logical-text, word-position and Preview checks. Then regenerate full regression,
isolated package, independent OFD-reader, font/GID and visual evidence from the
same final commit; verify latest-main/PR48 ancestry, storage, CI and valid review
findings before the existing push/merge gates. Diagnostic files do not satisfy
the final package gate.
