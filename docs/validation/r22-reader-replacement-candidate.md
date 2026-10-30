# R22 replacement reader and geometry candidate

Status: **candidate validation protocol, not a frozen acceptance contract**.
The 2026-10-30 [reader decision](r22-mixed-paragraph-validation.md#reader-contract-update-2026-10-30)
removes pinned PdfPig 0.1.15 default `Page.Text` and `GetWords()` as RTL
logical-order pass gates. It retains their raw outputs as versioned
compatibility observations. The writer remains fail-closed for two words.

## Fixed controls

Use the exact DejaVu Sans R22 fixture, SHA-256
`d7441d42a24882db09640ea4bb768b17a106cdb31ee9749a52cb34a075b8fbcc`.
Do not regenerate or relabel these reference files:

| Role | Existing one-page PDF | SHA-256 | Declared observation |
| --- | --- | --- | --- |
| Shaped logical-reader positive | `fpdf2-two-word-cell.pdf` | `7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8` | PDFKit 1451.5.3 selected `abc שלום מקום def` and found each logical Hebrew word once; visual page correct. It is not product acceptance. |
| Repeated-word discrimination | `fpdf2-repeated-cell.pdf` | `08e7439bd2d02f60aa52957a14104073709266c275d51ba613b09d9300cc840e` | Existing reference for two occurrences of `שלום` at distinct positions. Recheck every rectangle. |
| Unshaped negative | `fpdf2-two-word-fixed.pdf` | `7b69e87771891ba440bde051af33284f3e008b376744751bc27064391c366ea5` | PDFKit logical Finds failed and the Hebrew appearance was wrong, despite PdfPig default words being logical. |
| Reader-order negative | `pango-two-word.pdf` | `c20d0d39456727ea7f967b679120cb6ffc0f52bb89cfc7816b746f48d7a99513` | Visual phrase correct; PDFKit logical Finds failed. |
| Word-carrier negative | `type3-word-two-word.pdf` | `65e1b4aba679d163a4e44473023167fa5885fa040f3c91902bf43204f1df2df1` | Both word Finds worked but PDFKit whole selection was `שלום abc\ndef מקום`; one PDF character per word loses character granularity. |
| Preview wrap negative | `fpdf2-between-rtl-zero-margin.pdf` | `9c40fc51de911841fecf8dc44dafe8c88bd392bf525a4f9f85112df8648eb3c1` | Preview twice showed a found-page banner for `מקום` without a usable result or highlight. An API match alone must not count as GUI success. |

These PDFs and their existing raw reader/Preview receipts are in the ignored
`artifacts/r22-next-phase/reference-20261029/` working set and the curated
[evidence archive](../evidence/r22-next-phase/README.md). Reference results
remain reference results, not output from the public API → OFD → PDF path.

## Candidate checks to validate before freezing

1. **Logical reading and word search:** use native macOS PDFKit's exact
   whole-page selection, `PDFDocument.findString` for each logical and reversed
   word, counts, pages and all selection bounds. Compare Unicode scalar arrays
   without reversing, stripping controls or whitespace normalization. For a
   product candidate, repeat raw clipboard copy and actionable Find/highlight
   in native Preview on the exact generated PDF under a separate GUI grant.
2. **Independent text and geometry:** keep PdfPig 0.1.15 raw `Page.Text`,
   default/ActualText `GetWords()` and per-letter values/bounds as compatibility
   data; test whether `Page.Letters` gives useful independently parsed semantic
   carrier bounds without treating its order as logical. Cross-check PDFKit
   [character bounds](https://developer.apple.com/documentation/pdfkit/pdfpage/characterbounds%28at%3A%29)
   and word Find bounds against Poppler 26.09.0 `pdftotext -bbox` rectangles.
   The structural checker must separately verify embedded font identity,
   `/ToUnicode` and visible original-GID draw points against the OFD source
   spans, positioned glyphs and independently rendered page. A word-level
   semantic carrier does not by itself prove character geometry.
3. **Controls and failure rules:** require the shaped two-word reference's
   exact logical PDFKit text and both separate positive word rectangles;
   repeated-word controls need two distinct positions. The unshaped, Pango,
   Type 3 and Preview-wrap negatives must retain their declared failures. If
   any API cannot expose stable character bounds, a structural mapping cannot
   associate source characters with the drawn original GIDs, or independent
   rectangles disagree beyond a predeclared tolerance, reject this method and
   do not freeze it. Record exact tool/API versions, every raw string and
   codepoint, all bounds, font/subset hashes, file hashes, and environment.

The first read-only Poppler 26.09.0 precheck of the existing shaped two-word
reference returned `pdftotext -bbox` word strings `םוקמ` and `םולש` at distinct
positive rectangles. The unshaped negative returned logical word strings
`שלום` and `מקום` despite incorrect Hebrew appearance. Poppler's word text
therefore cannot serve as the RTL logical oracle; its physical rectangles
remain a candidate independent geometry signal pending validation. No
replacement reader or geometry method is frozen by this precheck.

## Product and import boundary

After the method passes the controls, freeze its versions, APIs, exact
expectations and tolerance, then evaluate a general product candidate from
public API through OFD write/read and PDF. Preserve source intervals, original
font/GIDs, actual glyph origins, logical word Find and true Preview raw copy
and Find. Re-run the fixed corpus, shared regression, isolated package
consumer, CI, SES, review and visual gates on the final product head. The
`4d9f003a` packages and all earlier diagnostic PDFs keep their old identities.

PDF→OFD remains a separate defect: its default invisible semantic layer and
text fallback both consume PdfPig's default `GetWords()`. This export protocol
does not change or validate reverse import of RTL or unknown external PDFs.

## 2026-10-30 bounded method result

The six fixed input hashes and the font hash matched the request. The probe
ran at source `2f45eaa20990eff29cfbea4d1f4866d231517e50`, macOS 26.6.2,
.NET SDK 10.0.401, PdfPig 0.1.15, PDFKit bundle 1451.5.3, Poppler 26.09.0,
pypdf 6.21.2 and fontTools 4.68.0. Before inspecting the new character
results, the [preregistration](../evidence/r22-reader-method/README.md)
fixed bottom-left PDF coordinates, Poppler Y conversion, same-scalar
one-to-one association, 0.5 pt character X-center / 2 pt Y-center ceilings,
and the word-box overlap/center criteria. Its SHA-256 is
`62a0c89682f30d52fb03c73ce0c6af8c1634c668e9b5dbd600435f1a56947243`.
No tolerance was changed after the raw probes.

Both shaped positives passed PDFKit whole selection and logical Find counts:
`abc שלום מקום def` found each Hebrew word once, and
`abc שלום שלום def` found `שלום` twice at distinct positions. PDFKit Find
rectangles and independently parsed Poppler `-bbox` physical word boxes
matched the predeclared center and overlap criteria. Poppler's word *strings*
were reversed; they were never used as a logical-text oracle. The unshaped,
Pango and Type 3 negatives retained their declared logical or character
granularity failures. The wrap reference retained its known native Preview
Find failure regardless of its headless API result.

**The proposed PDFKit-character-index to PdfPig visible-letter association
failed decisively:** zero of eight Hebrew character positions met the
predeclared same-scalar tolerance in each shaped positive. For example, in
the two-word PDF, PDFKit's logical page string labels character index 4 `ש`
with a box centered near X 56.61 pt, while PdfPig identifies the visible
glyph near X 56.09 pt as `ם`. The nearest PdfPig box for the *same* `ש` is
many points away. Thus `PDFPage.characterBounds(at:)` paired with the logical
page-string index does not independently identify each visible Hebrew glyph
in this reference. Word-level agreement does not repair that character-level
failure. The candidate method is **falsified and not frozen**; the writer
remains fail-closed.

A separate read-only inspection after this failure decoded the shaped
references' `/ToUnicode`, `/CIDToGIDMap`, embedded TrueType subset and text
origins. For each of the eight Hebrew draws in the two-word and repeated
controls, the embedded subset outline and advance matched the corresponding
nominal original-fixture GID, and the raw PDF text origin matched exactly one
PdfPig letter baseline. This is a promising *different* structural character
check to preregister and validate on holdout files. It does not overturn the
failed PDFKit character-index method, prove product vector-path geometry, or
freeze a new acceptance contract. The product converter draws original GIDs
as paths and adds a word-level semantic carrier, so a structural checker must
also account for that representation before product use.

The [curated LFS evidence bundle](../evidence/r22-reader-method/README.md)
contains the preregistration, probes, all raw reader/structure JSON, fixed
criteria verdict and supplementary structural map with individual hashes.
No new PDF, native Preview action, product conversion or full regression was
run in this bounded batch. Raw clipboard copy and actionable Preview Find
remain required for a later exact product PDF.
