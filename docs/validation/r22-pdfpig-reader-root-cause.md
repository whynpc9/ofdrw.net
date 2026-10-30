# R22 shaped-reference reader trace

Status: read-only diagnosis on `codex/r22-next-phase-architecture` at
`e643a6d1e7213d27bb93fb1d61385bb0c2335867`. This does not admit the
two-word writer profile. The immutable local receipts are in ignored
`artifacts/r22-next-phase/reference-20261029/`.

## Fixed input and observed disagreement

The shaped fpdf2 `.cell()` PDF for `abc שלום מקום def` is one page, SHA-256
`7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8`.
It uses the same DejaVu Sans fixture (input SHA-256
`d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc`)
as the product probe. The PDF has an embedded subset and `/ToUnicode`, with
no `/ActualText` token. Its per-character mappings decode to the expected
Hebrew scalars; the positioned Hebrew character operators are in **visual**
left-to-right order. For example, the left word's positioned codes decode
`םוקמ` in operator order, while the visible word is `מקום`.

The pinned PdfPig 0.1.15 `Page.Text` returns `abc םוקמ םולש def` and
`GetWords()` returns the two reversed Hebrew tokens. `UseActualText=true`
has the same result. PDFKit 1451.5.3 returns the exact source from whole-page
selection and finds each logical Hebrew word once with distinct positive
rectangles; it finds neither reversed word. Those are observations from
`pdfpig.json`, `pdfkit.json` and `analysis.json`, not a claim about PDFKit's
private extraction algorithm or actual Preview behavior.

The Pango/Cairo two-word control, SHA-256
`c20d0d39456727ea7f967b679120cb6ffc0f52bb89cfc7816b746f48d7a99513`,
reaches a different split: PdfPig `Page.Text` includes the logical Hebrew
words but loses a space before `def`; its default `GetWords()` still reverses
both word tokens. PDFKit selection and logical Hebrew Find fail. Thus neither
mature generator's tested PDF satisfies the combined gate.

## Trace through the exact PdfPig assembly source

The NuGet 0.1.15 package identifies repository commit
[`f131f642976936e06ee91cb19d3ed728f9dd18b6`](https://github.com/UglyToad/PdfPig/tree/f131f642976936e06ee91cb19d3ed728f9dd18b6).
Its [`BaseStreamProcessor`](https://github.com/UglyToad/PdfPig/blob/f131f642976936e06ee91cb19d3ed728f9dd18b6/src/UglyToad.PdfPig/Graphics/BaseStreamProcessor.cs)
runs content operations in sequence; the text-show path decodes each character
code with the font's Unicode map. Its
[`ContentStreamProcessor`](https://github.com/UglyToad/PdfPig/blob/f131f642976936e06ee91cb19d3ed728f9dd18b6/src/UglyToad.PdfPig/Graphics/ContentStreamProcessor.cs)
creates a `Letter` for each rendered glyph and appends it to the letter list.
The optional ActualText branch needs a marked-content `/ActualText` value.

[`Page.Text`](https://github.com/UglyToad/PdfPig/blob/f131f642976936e06ee91cb19d3ed728f9dd18b6/src/UglyToad.PdfPig/Content/Page.cs)
joins those letter values in content order. The same page's `GetWords()`
passes the letters to
[`DefaultWordExtractor`](https://github.com/UglyToad/PdfPig/blob/f131f642976936e06ee91cb19d3ed728f9dd18b6/src/UglyToad.PdfPig/Util/DefaultWordExtractor.cs),
which sorts by descending Y then ascending X before grouping. The
[`Word`](https://github.com/UglyToad/PdfPig/blob/f131f642976936e06ee91cb19d3ed728f9dd18b6/src/UglyToad.PdfPig/Content/Word.cs)
concatenates values in that incoming order. It does not apply a bidi reversal
to the grouped Hebrew letters. The pinned
[`ParsingOptions`](https://github.com/UglyToad/PdfPig/blob/f131f642976936e06ee91cb19d3ed728f9dd18b6/src/UglyToad.PdfPig/ParsingOptions.cs)
defaults `UseActualText` to false; toggling it has no effect on this PDF
because the source PDF contains no `/ActualText`.

The mapping is **not demonstrated defective**: each observed CID maps to the
right Unicode scalar, and PDFKit obtains the logical words from the same
PDF. The two PdfPig APIs make different ordering decisions. This is evidence
of a pinned extractor capability limit for per-character positioned RTL words
and an API-specific oracle conflict, rather than proof that the product's
semantic mapping is correct in every other PDF. The 2026-10-30 explicit
[reader-contract decision](r22-mixed-paragraph-validation.md#reader-contract-update-2026-10-30)
supersedes the earlier use of PdfPig 0.1.15 default `GetWords()` and
`Page.Text` as RTL logical-order pass gates. Their unmodified results remain
versioned compatibility evidence; independent character mapping and geometry
checks remain required. This does not alter any recorded diagnostic result.

## Different, falsifiable encoding hypothesis and result

Try **one visible word glyph per shaped RTL word**, whose own character code
maps through `/ToUnicode` to the complete logical word. In a PDF-only
diagnostic, a Type 3 character program can draw the exact original shaped
word outlines at the existing word origin and advance. This is distinct from
the earlier hidden semantic CID plus visible outline, character-split
variants, marked-content tags, single `TJ`, or reordering the existing
operators: the visible glyph and Unicode mapping are one character unit.
The [Adobe PDF Reference 1.7](https://opensource.adobe.com/dc-acrobat-sdk-docs/pdfstandards/pdfreference1.7old.pdf)
defines Type 3 `CharProcs` as glyph descriptions keyed by names and selected
through the font's required `Encoding` entry. The proposed whole-word
`/ToUnicode` mapping still needs structural and reader validation; the
reference does not establish that the target readers will handle it as a word.

The predicted mechanism is narrow: PdfPig would receive one `Letter` whose
value is `שלום` or `מקום`, so its left-to-right letter sort cannot reverse
characters *inside* that word. PDFKit and Preview may still order words,
selection or Find differently; Type 3 font handling, word bounds, font
subsetting and accessibility may also fail. This is an experiment, not a
recommended production design.

Minimal test: build only the two-word and repeated-word one-page PDF
diagnostics at the same origins with the exact fixture outlines; record the
font dictionary, character program, multi-scalar `/ToUnicode` mapping,
content sequence and hashes. Compare page rasters against the shaped fpdf2
reference, then run pinned PdfPig 0.1.15 `Page.Text` and default/ActualText
`GetWords()` and PDFKit whole selection, logical/reversed Find counts and
distinct positive rectangles. A PDF structure or raster mismatch, any
wrong raw string/token, or missing Find/rectangle falsifies the candidate.
Only a passing API diagnostic merits a separately granted native Preview
copy/Find check and then a general product-path design. No generator or
reader gate should be changed to score it.

### Result of the bounded Type 3 diagnostic

The two-word PDF was generated under a separate exclusive build grant; its
SHA-256 is `65e1b4aba679d163a4e44473023167fa5885fa040f3c91902bf43204f1df2df1`.
It is one page with two visible Type 3 word character programs, each mapped
by `/ToUnicode` to its complete logical Hebrew word. The outlines came from
the original fpdf2 PDF's embedded shaped TrueType subset. No hidden duplicate
text or `/ActualText` was added. The 144 dpi page kept the visible arrangement,
but differed in 645 pixels confined to the Hebrew outline region; Type 3
rendering did not prove pixel or original-font identity.

Pinned PdfPig 0.1.15 returned the exact source `abc שלום מקום def` from
`Page.Text` in both ActualText settings and found each logical Hebrew word
once in `GetWords()` at distinct positive bounds. PDFKit 1451.5.3 also found
each logical Hebrew word once with positive distinct bounds and neither
reversed query. Its whole-page selection, however, was
`שלום abc\ndef מקום`, **not** the exact source. This is the first decisive
combined-gate failure, recorded in ignored
`artifacts/r22-next-phase/reference-20261029/type3-word-diagnostic-receipt.json`.
The repeated-word variant and native Preview of this candidate were not run.
One PDF character now represents each whole Hebrew word, so individual
character geometry and selection granularity are also lost. This result
rejects this diagnostic encoding for the current contract and does not
establish that every possible PDF encoding must fail.

The user has made that scope decision: validate and freeze a replacement
reader and independent character/word geometry method before judging a new
product candidate. The bounded [candidate protocol](r22-reader-replacement-candidate.md)
records the initial controls. Keep the writer fail-closed until the public API → OFD
write/read → PDF path has new evidence, including native Preview raw clipboard
and Find. The PDF→OFD converter still calls PdfPig's default `GetWords()` in
its semantic text layer and text fallback; treat resulting RTL import behavior
as a separate production defect, not as repaired by this export decision.
The present evidence does not prove every possible PDF encoding impossible.

## Separate PDF→OFD production defect and coverage boundary

`PdfToOfdConverter.AddSemanticTextLayer` calls `pdfPage.GetWords()` and writes
each returned token plus a space into an invisible OFD text element when a
page is rasterized and the text layer is enabled. `PdfToOfdOptions` enables the
invisible layer by default. The non-rasterized text fallback also joins
`pdfPage.GetWords()` tokens. Thus both production paths inherit this pinned
extractor's left-to-right position ordering for RTL words. The fixed shaped
fpdf2 reference shows reversed Hebrew tokens from that API, but this record
does not claim an end-to-end PDF→OFD import result for that reference or for
arbitrary imported PDFs. A dedicated import regression must check the exact
PDF hash, selected options, produced OFD text and geometry, and a positive LTR
control before a repair is claimed. No import input class is newly admitted by
the export reader-contract decision.
