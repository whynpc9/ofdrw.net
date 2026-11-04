# R22 paragraph layout validation (Issue #109, F01–F05)

Date: 2026-11-05. Source baseline: the Issue #109 branch on top of `main` `b1cef96`; the E2E receipt
records the exact commit in `sourceSha`. Design: [r22-paragraph-layout-engine](../architecture/r22-paragraph-layout-engine.md).
Reference: [paragraph layout engine](../content/reference/text-paragraph-layout.md).

Environment: Linux 5.15 (WSL2) x86_64, .NET SDK 10.0.401, HarfBuzzSharp 14.2.1.301 (native 14.2.1),
poppler `pdftoppm` 24.02.0. No GUI: macOS Preview acceptance was **not** performed.

## Summary

| Area | Result |
| --- | --- |
| Unicode conformance (full official files) | PASS — 0 failures |
| Paragraph engine and writer tests | PASS |
| Affected regression suites | PASS, except one pre-existing DOCX failure (identical on the base commit) |
| HarfBuzz E2E incl. new public-API probe | PASS |
| Automatic PNG rendering | PASS (3 pages) |
| PNG visual review | PASS for the pages listed below |
| macOS Preview acceptance | **Not done** (no GUI) — open |
| PDF Copy/Find, SVG selection/accessibility | **Not verified** — open |

## Functional tests

`UnicodeConformanceTests` (Unicode 17.0.0; decompressed SHA-256 checked against
`scripts/unicode/ucd-17.0.0-manifest.json`):

| File | Cases | Failures |
| --- | ---: | ---: |
| BidiTest.txt (every paragraph-level bit) | 770,241 | 0 |
| BidiCharacterTest.txt | 91,707 | 0 |
| LineBreakTest.txt | 19,338 | 0 |
| GraphemeBreakTest.txt | 766 | 0 |

`ParagraphLayoutTests` cover: per-cluster fallback across Noto Sans / LXGW WenKai / DejaVu / Noto Sans
Tamil with exact source coverage; lines ending only at UAX #14 opportunities and within the frame;
bidi visual run order and RTL right alignment; Arabic GSUB forms and line-end words equal to
standalone shaping (reshape at breaks); Tamil many-to-many clusters on grapheme boundaries; overflow
rejection vs. grapheme-only emergency breaks; justification (word spaces and CJK gaps); pagination,
paragraph spacing and first-line indent; style snapshots, paint run splits and the shared baseline;
rejection of tab, RLO, ZWJ and BEL with source offsets; missing coverage and missing bold face; hard
breaks, terminators, cancellation and line budget.

`ParagraphLayoutOfdTests` cover: one CGTransform per cluster with complete GID/code/glyph counts and
glyph-slot deltas; OFD read-back of logical text and of the exact original GIDs per object; a negative
control showing the same map without the writer marker stays fail-closed; PDF and SVG export with an
outline path per visible glyph and the logical text present; rejection of stroke paint and of a frame
wider than the page.

Regression suites run on this branch: Layout 387/387, Core 119/119, Packaging 1,683/1,683,
Converter.Pdf 1,176 passed + 17 skipped, Converter.Pdf.Vector 474/474, Graphics.SkiaSharp 68/68,
Cli 272 + 1 skipped, Signatures 112/112, Crypto.Password 107/107, Converter.Docx 360 passed,
1 skipped, 1 failed.

The DOCX failure is `BuiltInPdfItalicRendererTests.DifferentPageSizesKeepHeadersFieldsTablesAndImagesThroughDrawingSeam`
(`98.5816` expected vs `98.5815` at 3-decimal rounding). It fails identically on the unmodified base
commit `b1cef96` in this environment, so it is pre-existing and unrelated; it is not fixed here.

Review follow-up (Codex on PR #112): adjacent text with the same face, script, language and features
at the same level is now passed as HarfBuzz context, so a paint-only span inside an Arabic word keeps
its joined forms (regression test with `ببب`); grapheme emergency breaks binary-search rebuilt widths
instead of trusting the per-cluster estimate (regression test: DejaVu `ffi` ligature in a 2 mm frame);
an empty text yields one styled empty line. Layout tests 387/387 and PDF tests pass after the change;
the pages above were regenerated and re-reviewed with no visual change.

A first reader design (structural fallback for any complete CGTransform cover) failed 10 existing
fail-closed PDF tests (ambiguous/miscounted Hebrew visual GIDs, partial combining spans). It was
replaced by the writer-owned `ClusterMapV1` marker; those tests pass unchanged.

## E2E and rendering

`e2e/Ofdrw.Net.Text.HarfBuzz.E2E` (`ParagraphLayoutProbe`) uses public APIs only: layout → OFD →
`OfdReader` read-back (run texts in order, page count, and exact restoration of the source from line
runs + terminators) → `OfdToPdfConverter` and `OfdToSvgConverter`. CI renders the PDFs with
`pdftoppm -r 144`. Viewing chain for PNG review: **native layout → OFD → PDF → PNG** (no direct path).

Sample: Chinese, English, Hebrew, Arabic (with digits and an English word), Tamil (conjuncts க்ஷ,
ஸ்ரீ, pre-base vowels), mixed sizes (4.2/5 mm) and colors, a U+2028 line separator.

| Output | Bytes | SHA-256 |
| --- | ---: | --- |
| r22-paragraph-justify.ofd | 498,988 | `e6e653e7852def647c912d2a6ca6ca12738abe3d0c8d6c1e38b7351aee0fcb2c` |
| r22-paragraph-justify.pdf | 653,818 | `7bb0456e2ee9c1ad1ddd4b78e7429be6a5a831b18be8b697e30b9c6782ac3f97` |
| r22-paragraph-justify.svg |  | `904a8180c5e60c3c79e8883ba7f082d6c8317691237c34ba23ac4929a0826d37` |
| r22-paragraph-justify-page-1.png |  | `b9e068f7c2cfa9b963d15a4bc48c6f1a98f1c2bcf8ed5fddb4c2ce269a9d5f37` |
| r22-paragraph-justify-page-2.png |  | `8c44759490957d50189301a328c8045f9f22d0c75c333637b22dce442e226f7f` |
| r22-paragraph-rtl-center.ofd | 498,747 | `b973e5586a1566e2cc37f8b94c85616cd7ebb482545026fb6c1444d20d8ca8a9` |
| r22-paragraph-rtl-center.pdf | 674,675 | `f949d2add4ca24650eaab9a0f7b1e1e6876de454a8028042fab5fd331e32e685` |
| r22-paragraph-rtl-center.svg |  | `efc4fa7acce9d7305deda80f6aa8287c2545d73d39fbf0d3cbeccc5efe1790c4` |
| r22-paragraph-rtl-center-page-1.png |  | `7781708419aea006a19878a8e4ab8ff5ea9b2e23060adcfbadb40b7a7d6d512e` |
| r22-paragraph-receipt.json |  | `6236fc964471b4c193324845fc5a71b35865bceb67bdcd60419b0091185aa4b6` |

Hashes identify this local run; outputs stay in ignored `artifacts/` and CI artifacts and are not
committed. Embedded fonts are GID-preserving subsets (WenKai 25.6 MB source → about 0.95 MB).

## PNG visual review

Reviewed all three pages at 144 dpi and the Arabic/Tamil lines of an equivalent start-aligned page at
300 dpi:

- Chinese and English: no missing glyphs or tofu; consistent baseline across 4.2/5 mm runs; red
  "English" keeps its color without spreading; ideographic comma/full stop never start a line.
- Justify: non-final lines fill the 90 mm frame; expansion is at word spaces and between CJK
  graphemes; paragraph-final lines and the line before U+2028 are not stretched; first lines are
  indented 6 mm on the start side (right side for the Arabic paragraph).
- Hebrew in an LTR paragraph: "שלום עולם" reads right-to-left in place, "(ver. 2.0)" stays LTR.
- Arabic: joined initial/medial/final forms, Arabic comma, LTR "123" and "English" embedded
  correctly, period at the left end; no broken joins at the line break.
- Tamil: pre-base vowel signs, க்ஷ and ஸ்ரீ conjuncts drawn correctly; no overlap or clipping.
- Forced RTL page: centered lines; neutrals at paragraph ends move to the left edge as UAX #9 requires.
- Pagination: page 1 ends after the Arabic paragraph, page 2 starts at the content top with no
  carried paragraph spacing; no blank pages, clipping or duplicated text observed.

Known visual limitation: Arabic justification widens word spaces (no kashida), so short justified
Arabic lines can show wide gaps.

## Open items

- macOS Preview acceptance of the generated OFD → PDF pages (requires a desktop run; keep this report
  open until done).
- PDF Copy/Find and selection geometry: logical text is drawn invisibly at each run's first glyph
  slot, so selection rectangles do not follow RTL glyphs; not verified with real readers.
- SVG selection/search/accessibility beyond the `title`/`aria-label` presence check.
- Flow/DOCX migration, stroked text output, synthetic styles, non-BMP output (F08), PDF→OFD RTL
  semantic import (F06). F07 is out of scope.
