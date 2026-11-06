# R22 engine switch and F08: comparison and validation (Issue #109)

Source commit `29e8b7cabaec55b3b0d5485d5d6b1482f79b2a7c` (branch `t3/r22-engine-switch-f08`, on
`main` `ba58b5519b854e6eaf37e43cb39ed2a96dc79b45`). Ubuntu 24.04.4 LTS (WSL2), .NET SDK 10.0.401,
runtime .NET 10.0.12, poppler `pdftoppm`, ImageMagick `compare`/`montage`. Design:
[r22-engine-switch-and-f08](../architecture/r22-engine-switch-and-f08.md). Evidence:
[`docs/evidence/r22-engine-switch/`](../evidence/r22-engine-switch/README.md).

**Scope.** The engine is opt-in; defaults are unchanged. This record compares the existing
visual-gate inputs under both engines with the same DOCX, fonts and options, and validates the F08
output features. It does not flip the default and is not a release acceptance.

## Functional tests

| Suite | Result |
| --- | --- |
| Layout (incl. 17 F08 and 3 Flow/plain-text switch tests) | 410/410 |
| Converter.Docx (incl. 3 DOCX Native switch tests) | 363 passed, 1 skipped, 1 failed — the pre-existing `BuiltInPdfItalicRendererTests.DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam` 0.0001 mm rounding assertion (98.5816 vs 98.5815), reproduced identically on `main` `ba58b55` without this change |
| Converter.Pdf | 1,176 passed, 17 skipped (one CFF test updated: PDF export of a static CFF1 face now succeeds through the TrueType rendition instead of refusing) |
| Core 119, Packaging 1,683, Pdf.Vector 474, Skia 68, CLI 272 (1 skipped), Signatures 112, SesInterop 44, Crypto 107 | all passed |

Discriminating checks added: managed CFF outlines equal independent fontTools `RecordingPen` paths
for four CID-keyed Noto Sans SC subset GIDs and the four existing Source Code Pro GIDs; the CFF
subset keeps every GID slot with most CharStrings blanked and identical outlines for used GIDs;
supplementary characters count once in `CodeCount`; tab, decimal, center, end and hanging-indent
positions; synthetic stroke/shear in OFD XML; decoration paint order; plain-text line/page-control
output identical to the compatibility layout for edge cases; DOCX `w:tabs` positions in millimeters.

## Engine comparison (DOCX → native OFD → PDF → 144 dpi PNG)

Every scenario produced the **same logical text** under both engines (compact text equality per
document), and PDF text extraction equals the OFD text on every page.

| Input | Pages legacy / engine | Same page breaks | RMSE per page | Reviewed differences |
| --- | --- | --- | --- | --- |
| `generated-layout.docx` | 2 / 2 | yes | 0.080, 0.042 | Lines are about 0.3 mm lower/tighter (real ascent/descent instead of 1.3 em); table rows slightly shorter. Bold 量 in the table header renders with open counters (after fix below). |
| CN/EN headers sample | 5 / 5 | yes | 0.020–0.021 | Same headers, bodies and PAGE/NUMPAGES fields; sub-millimeter baseline shifts. |
| CN/EN inline JPEG sample | 1 / 1 | yes | 0.025 | The image paragraph keeps the compatibility layout (`DOCX_NATIVE_TEXT_ENGINE_FALLBACK`); the label line is about 0.4 mm shorter. |
| Flow sample (title, local bold/italic/red, 30 indented paragraphs) | 2 / 2 | no | 0.321, 0.231 | Legacy binds name-only SimSun with Arial-compatible widths; the engine shapes with explicit Noto Sans + Noto Sans CJK SC, so Latin wraps one word earlier and one paragraph moves to page 2. Bold/italic are synthesized. Expected. |
| Plain text (40 tab-separated rows, form feed) | 2 / 2 | yes | 0.258, 0.028 | Legacy measures a tab as four spaces; the engine aligns tabs to 12.7 mm default stops. Where text ends just before a stop the tab advance was near zero (default-stop semantics); Preview review in Issue #115 found columns touching (`数量 12说明`), so Flow/plain text now skip a stop closer than half an em (see the Issue #115 follow-up below). |

Defects found during the review and fixed before the final run: (1) `AtLeast` line height put the
extra space above the text, so the last body line touched the following table; it now goes below,
as in the compatibility layout. (2) Synthetic Han bold written as a 1%-em stroke rendered as a closed
blob in PDF (Issue #107); path glyphs and the historical one-to-one route now use the horizontal-only
emphasis for final layout output. (3) Decorations stepped between Latin and Han faces; they now use
the style's primary face metrics. (4) A CFF face could reach the PDF as raw CFF declared TrueType;
the alias now uses the TrueType rendition.

Output size: DOCX OFDs are within 1%; with the engine, Flow/plain-text OFDs embed font subsets
(356 KB / 336 KB instead of 4 KB name-only) and their PDFs draw outlines plus a semantic text layer
(1.9 MB / 632 KB instead of 112 KB / 76 KB).

## F08 output review (OFD → PDF → PNG)

`f08-visual-sample` (one page) and `flow-engine` (one page) were reviewed at 144/200 dpi with
zoomed crops: default and explicit tab stops (start 25 mm, center 80 mm, decimal 140 mm), synthetic
bold, oblique and bold-oblique (solid at 800 dpi), underline/strikethrough level across fallback,
highlight band, letter spacing, Ext-B 𠂇 (U+20087) and 𠀋 (U+2000B), Arabic joining and the CID-keyed
CFF Han face. No missing glyphs, overlap, clipping or duplicated text was seen; `pdffonts` lists only
CID TrueType fonts.

## Issue #115 follow-up (Preview review findings)

macOS Preview review of the merged switch (`6cbaaf5`, `OFD → PDF → Preview`) reported two defects;
both were fixed and re-checked from regenerated artifacts on Linux (OFD → PDF → Poppler PNG).

| Finding | Cause | Change | Re-check |
| --- | --- | --- | --- |
| A faint vertical line where `highlight 高亮` crosses from the Latin to the Han face | The highlight was two abutting same-color PathObjects (one per fallback run, serialized edges 0.0005 mm apart); anti-aliasing viewers draw a seam between separately filled shapes | The writer merges abutting decoration rectangles of one kind, color and vertical extent on a line into one path | `f08-visual-sample` page 1 now has 3 PathObjects (one highlight `101.971 30.855 29.156 6.129`, one underline, one strikethrough) instead of 4; 400 dpi crop shows one band |
| Plain-text rows `数量 12\t说明` with almost no gap | Text ended just before a 12.7 mm default stop; the tab advance was near zero | Flow and plain text skip a start-aligned or implicit stop closer than 0.5 em (`ParagraphLayoutOptions.WithMinimumTabAdvance`, CSS `tab-size` precedent); DOCX Native keeps Word's rule | `plain-text-engine` page 1: rows 04–40 move `说明` to the next stop with a visible gap; rows 01–03 keep the nearer stop, so the fourth column sits on two different stops (inherent to default stops) |

Regression tests: `AMinimumTabAdvanceSkipsAStopTooCloseToTheText`,
`PlainTextTabsLeaveAtLeastHalfAnEmBeforeTheNextColumn` (fails without the change: 1.77 mm < 1.85 mm),
`DecorationsUseFaceMetricsAndAreWrittenInPaintOrder` (one path per decoration across fallback; fails
without the change) and `DecorationsOfDifferentColorsOrSizesStaySeparate`. macOS Preview re-check of
the regenerated PDFs is still outstanding.

## Not completed

- **macOS Preview: not performed** (no GUI on this host). PNG review does not replace it; the switch
  needs Preview acceptance before any default flip.
- PDF Copy/Find and SVG selection were not verified in real readers.
- Real-world DOCX beyond the listed samples, Source Han Sans full fonts and other viewers (Suwell,
  Foxit OFD) were not checked. ZWJ/ZWNJ/variation-selector text is refused by PDF export by policy.
