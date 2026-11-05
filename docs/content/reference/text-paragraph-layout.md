---
title: "Paragraph layout engine (R22)"
description: "Unicode 17.0.0 双向、断行与字素簇，字体回退、逐行重塑形、对齐/两端对齐、分页，以及直接消费最终排版结果的 OFD writer。"
sidebar:
  order: 130
  label: "段落排版引擎"
---

> Available since [0.2.0](https://whynpc9.github.io/ofdrw.net/releases/020/).

`Ofdrw.Net.Layout.Text.Paragraphs` lays out styled multi-paragraph text into final lines and
pages, and `ParagraphLayoutOfdWriter` writes that exact result to OFD. It is a new entry point:
`FlowDocument`, `PlainTextOfdConverter`, DOCX conversion and the earlier R21/R22 queries keep
their existing behavior. Nothing is migrated implicitly.

```csharp
var noto = new TextFontFace("Noto Sans", File.ReadAllBytes("NotoSans-Regular.ttf"));
var wenkai = new TextFontFace("LXGW WenKai", File.ReadAllBytes("LXGWWenKai-Regular.ttf"));
var dejavu = new TextFontFace("DejaVu Sans", File.ReadAllBytes("DejaVuSans.ttf"));
var style = new TextStyle(new[]
{
    new TextFontFamily("Noto Sans", noto),
    new TextFontFamily("LXGW WenKai", wenkai),
    new TextFontFamily("DejaVu Sans", dejavu),
}, fontSizeMillimeters: 4.2);
var request = new ParagraphLayoutRequest("排版 layout שלום مرحبا\nSecond paragraph.", style,
    new ParagraphLayoutOptions(widthMillimeters: 90, pageHeightMillimeters: 250,
        alignment: TextAlignment.Justify, lineHeight: TextLineHeight.Multiple(1.2)));
var layout = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(request);
var package = ParagraphLayoutOfdWriter.CreateDocument(layout, 110, 280, 10, 15);
await new OfdPackageWriter().WriteAsync(package, output);
```

## Pipeline

1. **Admission.** Unpaired surrogates, tabs, C0/C1 controls other than line/paragraph separators,
   explicit bidi formatting characters (LRE…PDI, LRM/RLM/ALM), other format controls (ZWJ, ZWNJ,
   soft hyphen, word joiner, BOM) and variation selectors are rejected with a source interval
   before any shaping. Style spans must tile the text and may not split a grapheme cluster.
2. **Paragraphs** split at UAX #9 paragraph separators admitted here: LF, CR, CRLF, NEL, U+2029.
   VT, FF and U+2028 are hard line breaks inside a paragraph. Terminators are part of a line's
   source interval (`TerminatorLength`) but have no glyphs.
3. **Unicode analysis** with frozen Unicode 17.0.0 tables: UAX #9 rev 51 (P2–P3, X1–X10, W1–W7,
   N0–N2, I1–I2, L1–L2), UAX #14 rev 55 default line breaking, UAX #29 rev 47 default extended
   grapheme clusters. No tailoring, dictionary segmentation or hyphenation. See
   [conformance](#unicode-conformance).
4. **Itemization** by resolved bidi level, script (common/inherited characters take the
   neighbouring script), selected face and style. Each grapheme cluster takes the first family in
   the style's fallback chain whose face with the requested bold/italic declaration covers every
   scalar of the cluster; common-script clusters stay in the previous face when it covers them.
   No host font discovery and no synthetic bold/italic: a missing styled face fails with
   `UnsupportedFontProfile`, missing coverage with `MissingGlyph`.
5. **Shaping** each item with `ITextShaper` (HarfBuzz by default) for measurement.
6. **Line breaking.** Greedy first fit over UAX #14 opportunities that are also grapheme
   boundaries. Trailing spaces hang. The chosen line is then rebuilt: runs cut by the line edges or
   by L1 level changes are **reshaped** and the real width is checked; on overflow the previous
   opportunity is tried (bounded to 32 attempts, then `BudgetExceeded`). Because the estimate charges
   cross-break kerning to the left side, up to three further legal breaks estimated within half an em
   over the frame are also rebuilt before the estimate is trusted. A segment without an
   opportunity fails with `LayoutOverflow` (`TextOverflowPolicy.Reject`, default) or is broken at
   grapheme boundaries (`BreakGraphemes`); never inside a cluster. Candidate cuts are rebuilt and
   reshaped. Rebuilt widths are not monotonic (an Arabic lam-alef ligature is narrower than lam alone),
   so a binary search is followed by a bounded upward scan that stops after four consecutive
   overflowing cuts. The chosen cut is the largest fitting one found, not proven globally largest.
7. **Reordering.** L1 per line, then L2 over runs. Within a run the shaper's visual glyph order is
   kept.
8. **Metrics and alignment.** All runs share one baseline. Normal line height uses each face's
   ascender, descender and line gap at its size; `Multiple` and `Exact` are available.
   `Start`/`End` follow the paragraph direction; the first-line indent is on the start side.
   `Justify` stretches non-final lines at word spaces and between graphemes next to Han, kana,
   Hangul or Bopomofo (not before closing punctuation); the extra advance is recorded per glyph.
9. **Pagination** into pages of `PageHeightMillimeters` (or one unbounded page). Paragraph spacing
   is dropped at a page break; a line taller than the page fails with `LayoutOverflow`.

## Final result contract

`ParagraphLayoutResult` is immutable and complete; failures and cancellation return no partial
result. Coordinates are millimeters relative to each page's content box, Y down, converted from
design units once.

| Level | Carries |
| --- | --- |
| Result | request (immutable styles), pages, `UnicodeVersion` = 17.0.0, shaper `EngineVersion`, shape-call count |
| Page | index, lines, used height |
| Line | absolute UTF-16 source interval incl. trailing spaces and terminator, paragraph/page index, base level, hard-break flag, top, baseline, height, ascent, descent, aligned content left/width, ink bounds, runs in visual order |
| Run | logical source text and start, exact `TextShapedRun` (face byte identity, script, language, direction, features), resolved bidi level, `TextResolvedStyle`, origin, baseline, advance, ink |
| Glyph | original GID, absolute cluster start/length (many-to-many), origin X/Y, advance incl. justification, justification amount, ink |

`TextResolvedStyle` snapshots paint (fill/stroke colors in DeviceRGB with alpha, stroke width),
size, language, requested vs. face bold/italic, the family and its fallback index, and
`Synthetic = None`. A writer never consults the request style again, reselects a font or reshapes.

## OFD output

`ParagraphLayoutOfdWriter.CreateDocument(result, pageWidth, pageHeight, left, top)` creates one OFD
page per layout page and one `TextObject` per run:

- the `TextCode` holds the run's unmodified logical text;
- one `CGTransform` per source cluster, in logical order (`CodePosition`/`CodeCount` count BMP
  characters), lists the original GIDs;
- glyph slots follow that order and every slot has an explicit origin (complete `DeltaX`/`DeltaY`),
  so RTL runs use negative deltas instead of reordered text;
- each distinct face is one font resource, embedded as a GID-preserving subset (or in full with
  `embedFullFonts: true`); OS/2 embedding restrictions are honored;
- fill color and alpha are written; the run's face supplies bold/italic.

Rejected before a package is created (`UnsupportedOutput`): stroke or no fill, supplementary
(non-BMP) characters, frames or ink that leave the page, and lost source/GID identity.

The model exposes this as `OfdPositionedGlyphMapping.ClusterMap`. The writer marks such a
TextObject with the writer-owned hint attribute `ClusterMapV1="true"` in the
`https://ofdrw.net/layout-hints` namespace (other readers ignore it). When reading, the historical
bounded CGTransform profiles keep precedence; a marked TextObject they reject is accepted only as a
complete cluster map (one TextCode, every code unit in exactly one ascending CGTransform, an origin
for every glyph slot). Unmarked foreign CGTransforms keep the historical fail-closed rules. PDF and
SVG draw every glyph of an accepted map from its original outline and keep the logical text as
invisible PDF text / SVG `title`/`aria-label`.

## Unicode conformance

Tables are generated by `scripts/unicode/generate_unicode_tables.py` from UCD 17.0.0 files pinned by
SHA-256 in `scripts/unicode/ucd-17.0.0-manifest.json` (Unicode License v3). The full official tests
run in `UnicodeConformanceTests`:

| File | Cases | Failures |
| --- | ---: | ---: |
| BidiTest.txt | 770,241 | 0 |
| BidiCharacterTest.txt | 91,707 | 0 |
| LineBreakTest.txt | 19,338 | 0 |
| GraphemeBreakTest.txt | 766 | 0 |

`UnicodeTextAnalysis` exposes the same algorithms (grapheme boundaries, line break opportunities,
bidi levels and visual order). L3/L4 are not applied by the bidi API; mirroring is the shaper's.

## Limits and what is not claimed

- Budgets: `MaximumTextLength` (262,144 UTF-16), `MaximumLines` (10,000), `MaximumShapeCalls`
  (100,000), 4,096 UTF-16 per shaped run, ±1,000,000 mm coordinates.
- Verified scripts for output are those in the validation report: Latin, Han, Hebrew, Arabic and
  Tamil with the listed fonts. Other scripts lay out through the same pipeline but are not claimed.
- Not in this profile: tabs, ZWJ/ZWNJ/variation selectors, vertical text, kashida justification,
  hyphenation, language tailoring, synthetic styles, stroked text output, non-BMP output.
- The 0.2.0 candidate passed macOS Preview checks for the three fixed paragraph output pages. PDF selection geometry, Copy and Find remain unverified. See the [release notes](https://whynpc9.github.io/ofdrw.net/releases/020/).
