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

## Completion gates

Each result is PASS, FAIL or BLOCKED with file hashes and tool versions. Current
Preview copy/Find failures remain FAIL; a computer-use timeout is BLOCKED.
Keep PR #71 draft until one product-generated candidate passes the required
logical-text, word-position and Preview checks. Then regenerate full regression,
isolated package, independent OFD-reader, font/GID and visual evidence from the
same final commit; verify latest-main/PR48 ancestry, storage, CI and valid review
findings before the existing push/merge gates. Diagnostic files do not satisfy
the final package gate.
