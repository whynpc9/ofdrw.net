# F06: bounded PDF→OFD right-to-left semantic import

Issue [#109](https://github.com/whynpc9/ofdrw.net/issues/109) F06 (former N08).
Status: on `main` after merge, **not released**. The default behaviour is unchanged.

## Problem

PdfPig 0.1.15 `GetWords()` orders letters geometrically and does not apply
inverse bidi. On the shaped D3 reference it returns `abc םוקמ םולש def` for the
source `abc שלום מקום def`; on the unshaped negative it returns the source
string while the page shows the Hebrew reversed. Both are wrong evidence. PR #79
therefore rejects every page with R/AL letters or bidi controls
(`PDF_TEXT_READING_ORDER_UNVERIFIED`) and keeps an explicit image-only path.
F06 keeps that default and adds an opt-in profile that has to prove each line.

## API

```csharp
new PdfToOfdOptions
{
    RightToLeftTextImport = PdfRightToLeftTextImport.VisualLineProfile, // default Reject
    RightToLeftBaseDirection = PdfTextBaseDirection.Auto,               // or LeftToRight / RightToLeft
};
```

The profile applies to the transparent text layer over a rendered page only.
Pages without RTL letters take the unchanged legacy word path. The unrasterized
text fallback (`RequireRasterization=false` with no rasterizer) keeps rejecting
RTL: it draws visible text without bidi. `PdfVectorToOfdConverter` copies both
options into its `RasterizePage` fallback; its native vector text path is
unchanged.

## Evidence for each accepted line

1. **Painted geometry.** `PdfPaintedTextEvidence` scans the page's own content
   operations and maps every show-text operator to `Letter.TextSequence`. A
   glyph is evidence only when its rendering mode paints (fill/stroke), it is
   not under a non-rectangular clip, it lies inside every rectangular clip, and
   its resolved `ExtGState` alpha is above zero. Text inside form XObjects (the
   sequence count no longer matches), unresolvable transparency, soft masks,
   text clip modes, optional content, `/ActualText` and `/ReversedChars` reject.
   Invisible layers are a producer's claim about text, not visible glyphs.
2. **Geometry → visual order.** Glyphs are grouped by baseline (0.05 em). A
   different baseline within 0.6 em (superscript, overlap), rotated/vertical
   glyphs, a rotated page, non-positive advances, overlapping origins (step
   below half an advance) and gaps above 3 em (columns) reject. Word spaces come
   from gaps ≥ 0.15 em, not from `GetWords()`. Content-stream order is ignored.
3. **Character mapping.** Each glyph needs a decoded value. Only BMP letters
   of bidi class L, R or AL are admitted; presentation forms, digits,
   punctuation, combining marks and controls reject. A glyph mapped to several
   characters of one direction (lam-alef) stays one unit whose logical string is
   kept.
4. **Logical order, verified forward.** For each admissible base level a
   candidate logical string is derived (N1/N2 levels from the strong neighbours,
   then L2). It is accepted only when `UnicodeTextAnalysis.ResolveBidi` — the
   repository's Unicode 17 UAX #9 implementation that passes all 861,948
   BidiTest/BidiCharacterTest cases — reorders it into exactly the observed
   glyph permutation (not merely an equal string). Under `Auto` the candidate's
   own P2/P3 level must also match; a line consistent with both directions, or
   with neither, rejects.

Accepted lines are written as one transparent `TextObject` per logical word
(followed by one space, like the legacy layer). `TextCode` is logical; every
character keeps the origin of its source glyph through `DeltaX`, which is
negative inside right-to-left runs. Characters of one ligature glyph share an
origin.

## What it does not claim

- An input painted in the wrong visual order (the D3 unshaped negative) imports
  as what the page shows, not as its generator's source string.
- Per-line `Auto` cannot know a paragraph's direction from an earlier line;
  wrapped paragraphs usually need an explicit base direction.
- Real-reader Copy/Find of the re-exported PDF, other OFD readers and macOS
  Preview were not verified.
- Imported text keeps name-only fonts. Re-export through `OfdToPdfConverter`
  needs a host font covering the script or `OfdNameOnlyFontBindings`; otherwise
  the exporter fails closed as before.
- The repository's own OFD→PDF export writes clipped (invisible) semantic text,
  so its PDFs are rejected rather than re-read. While probing, the final-layout
  export of an RTL paragraph was found to place that invisible text in logical
  order at left-to-right advances; without the painted-glyph gate it was misread
  as reversed Arabic. This export-side geometry belongs to F05 (PDF Copy/Find)
  and is recorded, not changed, here.

Validation: [F06 validation record](../validation/f06-pdf-rtl-import-validation.md).
Fixtures: [`Fixtures/F06`](../../tests/Ofdrw.Net.Converter.Pdf.Tests/Fixtures/F06/README.md).
