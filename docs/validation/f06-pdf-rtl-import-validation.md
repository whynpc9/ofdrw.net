# F06 PDF→OFD RTL semantic import: validation

Issue [#109](https://github.com/whynpc9/ofdrw.net/issues/109) F06. Source baseline
`main` `6cbaaf5`; design in [`f06-pdf-rtl-import.md`](../architecture/f06-pdf-rtl-import.md).
Status: on `main` after merge, not released.

## Supported and rejected

| Input | Default (`Reject`) | `VisualLineProfile` |
| --- | --- | --- |
| Pages without RTL letters | unchanged legacy words | unchanged legacy words (byte-equal element shapes, tested) |
| Painted L/R/AL letters and spaces on horizontal lines, verified by UAX #9 | reject | logical words, per-character source origins |
| Line consistent with both base directions under `Auto` | reject | reject; explicit `LeftToRight`/`RightToLeft` resolves it |
| Digits, punctuation, NSM marks, presentation forms, non-BMP | reject | reject |
| Invisible/clipped/transparent text, optional content, form XObject text | reject | reject |
| `/ActualText`, `/ReversedChars`, rotated text or page, overlapping glyphs, superscript band, gap > 3 em | reject | reject |
| Unrasterized text fallback | reject | reject |

All rejections keep `PDF_TEXT_READING_ORDER_UNVERIFIED`, state the reason and
leave the output stream untouched. `TextLayerMode=None` with
`RequireRasterization=true` remains the image-only path.

## Functional results

`PdfRtlVisualLineImportTests` (fixtures in `Fixtures/F06`, SHA-256 checked):

| Fixture | Generator | Option | OFD text |
| --- | --- | --- | --- |
| `D3/fpdf2-two-word-cell` | fpdf2 (visual content order) | Auto | `abc שלום מקום def ` = declared source |
| `pango-two-word` | Pango/Cairo (logical content order) | Auto | `abc שלום מקום def ` |
| `fpdf2-arabic-ltr-cell` | fpdf2+uharfbuzz | Auto | `abc مرحبا لا def `; lam-alef chars share one origin |
| `fpdf2-arabic-rtl-cell` | fpdf2+uharfbuzz | Auto | `مرحبا بالعالم لا ` |
| `fpdf2-hebrew-rtl-wrap` | fpdf2+uharfbuzz | RightToLeft | `שלום עולם abc מקום טוב ` (Auto rejects line 1) |
| `pango-after-rtl-wrap`, `fpdf2-between-rtl-zero-margin` | Pango, fpdf2 | LeftToRight | `abc שלום מקום def ` (Auto rejects) |
| `libreoffice-rtl` | LibreOffice 24.2 | Auto | three paragraphs (RTL Hebrew, RTL Arabic, LTR), each exact |
| `D3/fpdf2-two-word-fixed` (unshaped negative) | fpdf2 | Auto | `abc םוקמ םולש def `: what the page shows, not its source string |
| `pango-between-rtl-wrap` / digits / niqqud | — | — | reject: `/ActualText` / EN / NSM |

`GetWords()` is wrong on both D3 controls in opposite directions; the profile is
right on both with respect to the painted page, so it does not reproduce the
negative's coincidental string match.

Readback: every non-space character's OFD origin (`Boundary` + `TextCode X` +
cumulative `DeltaX`) equals its source glyph origin within 0.003 mm; RTL deltas
are negative. Re-export through `OfdToPdfConverter` (DejaVu bound by
`OfdNameOnlyFontBindings`) writes the invisible letters in logical content order
at the source glyph origins (±0.05 pt). Re-importing that re-export is rejected
(clipped invisible layer), as is the repository's own final-layout RTL export.
Without the painted-glyph gate, that export was misread as reversed Arabic
during probing; it has a dedicated regression test.

Fourteen synthetic PDFs isolate each guard (render mode 3, clip outside the
line, non-rectangular clip, `ca 0`, `/ReversedChars`, optional content,
overlap, rotated glyph, wide gap, superscript, presentation form, mixed-direction
glyph, `/Rotate 90`, no consistent order). The tests also cover a containing
rectangular clip and `ca 1`, which are accepted, plus content-order independence,
explicit direction, an LTR line on an RTL page, budget, cancellation and option
validation. Existing D3 tests are unchanged and pass.

## Regression

`dotnet test Ofdrw.Net.sln` (Linux, .NET SDK 10.0.401): Core 119, Packaging
1700, PDF 1218 (+17 skipped), Signatures 112, CLI 272 (+1), Layout 412,
SesInterop 44, Skia 68, PDF Vector 474, Crypto 107 pass. DOCX 364 pass, 1 skip
and 1 failure: `BuiltInPdfItalicRendererTests.DifferentPageSizes…`, the
pre-existing 0.0001 mm rounding assertion that also fails on `main` (#109
comments), unrelated to PDF import.

`scripts/run-converter-package-e2e.sh` built and validated 12 local
`0.1.0-preview.8` candidate packages. The isolated consumer ran the D3 checks and
the new F06 step: four profile imports with exact text, re-export, and a
rejection of the digits fixture. Receipt: `f06-import-receipt.json` (loaded PDF
DLL SHA-256 `2face74e…40c76304`). These are local candidate packages, not a
public release.

## Visual review

Chain: source PDF → profile import → OFD → `OfdToPdfConverter` → PDF →
`pdftoppm -r 144` PNG, for `fpdf2-two-word-cell`, `fpdf2-arabic-rtl-cell`,
`fpdf2-hebrew-rtl-wrap` and `libreoffice-rtl` (one page each, all pages).

- PNG review (each page viewed against its source): text, Arabic joining,
  lam-alef, line placement and right alignment match the source. There is no
  visible duplicate text from the transparent layer, no clipping and no blank
  page. The raster is slightly softer. Pixels differing by more than 64/255:
  0.13–0.30 % per page.
- macOS Preview: **not performed** (no GUI). Copy/Find in Preview or other
  readers is not verified.

Evidence: [`docs/evidence/f06-pdf-rtl-import`](../evidence/f06-pdf-rtl-import/README.md).

## Limitations

Neutral/weak characters, combining marks, presentation forms and non-BMP text
stay out of scope. Paragraph direction is per line under `Auto`. The
re-export needs fonts covering the script. Real-reader text semantics are
unverified. The final-layout export's invisible RTL text geometry is an
F05/export-side observation, not changed here.
