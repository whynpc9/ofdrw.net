# R22 engine switch and F08 font/text items (Issue #109)

Status: implemented on the Issue #109 engine-switch branch; not released. Builds on the
[paragraph layout engine](r22-paragraph-layout-engine.md) (F01–F05). User guide:
[final text layout](../content/guides/final-text-layout.md); reference:
[paragraph layout engine](../content/reference/text-paragraph-layout.md); comparison:
[r22-engine-switch-comparison](../validation/r22-engine-switch-comparison.md).

## Goal and sequencing

Nobody benefits from the paragraph layout engine until Flow, plain text and DOCX Native use it, but
those paths carry years of visual baselines. The plan from Issue #109:

1. Add an **opt-in** switch (Flow / plain text, then DOCX Native). The default stays the
   compatibility layout. *(this change)*
2. Run the existing visual gates through both engines and compare page by page. *(this change, see
   the comparison record)*
3. Flip the default only in a later release, after differences are understood or removed.

F08 is not one package. This change pulls in the items the switch needs, in the priority order
agreed on the issue: CFF/OTF faces; tabs, synthetic bold/italic and non-BMP output (the switch
blockers); underline/strikethrough/highlight; and ZWJ/ZWNJ/variation selectors (nice to have).
Vertical text, variable fonts and color fonts stay out until a real document needs them.

## Switch design

Flow paragraphs, Flow table cells, DOCX body paragraphs, table cells and headers/footers all funnel
through `FlowParagraphLayout.Layout(...) → IReadOnlyList<FlowLine>`, and their pagination places
`FlowLine`s. The switch therefore sits at that single point:

- `FlowFinalTextLayout` maps inlines to a `ParagraphLayoutRequest` and returns `FlowLine`s whose
  `Final` property holds the engine's `ParagraphLayoutLine` (glyph list empty, height = final line
  height). Existing pagination, keep-together, table row measurement, header/footer layout and
  `OnElement` callbacks are unchanged.
- Line controls keep the compatibility semantics exactly: CR/CRLF/LF end a line inside the
  paragraph (first-line indent applies once), trailing line controls reserve terminal line heights,
  `\f` is a page control. Each line-delimited chunk is one engine request with `PageHeight = null`,
  so per-request budgets stay small for large plain text.
- Drawing a `Final` line calls `ParagraphLayoutOfdWriter.AppendLine` (the same code as
  `CreateDocument`) at the placed position; `FlowFinalTextWriter` owns one embedded font resource
  per face SHA-256.
- Paragraphs with inline images return `null` from the adapter and keep the compatibility layout
  (DOCX reports `DOCX_NATIVE_TEXT_ENGINE_FALLBACK` once). Inline objects in the engine are later work.
- Public surface: `FlowDocumentOptions.TextEngine` (`FlowTextEngine`: engine, explicit families,
  synthetic policy, language) and `DocxConversionOptions.NativeParagraphLayoutEngine`. The HarfBuzz
  adapter depends on Layout, so the engine is injected rather than referenced by Layout/DOCX.
- DOCX font resolution reuses `DocxFontCatalog` (FontDirectories / platform CJK directories and
  CJK alias groups); the requested family comes first, then configured and CJK substitute families
  that exist. Faces outside the engine profile are skipped with a diagnostic.

## F08 items

| Item | Decision |
| --- | --- |
| CFF/OTF | `TextFontFace` admits static OpenType/CFF1, plain or CID-keyed (FDArray/FDSelect, per-FD local subrs). `CffFontProgram` parses the program; `CffCharStringOutlines` is a managed Type 2 interpreter (all path, hint-mask, flex, subroutine and arithmetic operators except `random`; seac endchar rejected) with depth, operation and point budgets. Its outlines match independent fontTools `RecordingPen` oracles for Source Code Pro (plain) and a Noto Sans SC subset (CID-keyed). `GlyphOutlineSources.Open` selects TrueType or CFF for the OFD embedding plan and the PDF/SVG path renderers; the shared outline model gained a cubic verb. |
| CFF subsetting | `CffSubset` keeps every GID slot and rewrites unused CharStrings to `endchar` and unused subroutines to `return` in place, zero padded, so no CFF offset changes; OFD deflate removes the padding. The cmap is reduced like the TrueType subset. |
| CFF in PDF | PDFsharp embeds only TrueType (FontFile2). `CffTrueTypeConversion` builds a glyf rendition with identical GIDs, cmap and horizontal metrics (each cubic split into quadratics within one design unit) used only by PDF export; explicit-GID outlines still come from the CFF bytes. |
| Tabs | Tabs are their own items and never shaped. After L1 they sit at the paragraph level, so tab advances are computed in logical order from the start edge: explicit `TextTabStop`s (start, center, end, decimal), then implicit stops at `origin + k × interval`, clamped to the frame. A tab is written as the face's space glyph with the computed advance, keeping U+0009 in TextCode. Justification never stretches text before the line's last tab. DOCX maps `w:tabs`/`w:defaultTabStop` (margin-relative via the origin) and treats a hanging indent as a first-line stop. `WithMinimumTabAdvance(em)` skips a start-aligned or implicit stop closer than em × the tab's size, as CSS `tab-size` does; Flow and plain text use 0.5 em so columns never touch (Issue #115), DOCX keeps Word's zero minimum. |
| Synthetic styles | `ParagraphLayoutOptions.WithSyntheticStyles(Synthesize)` picks the closest declared face and records `TextSyntheticStyle.Bold`/`Oblique`. Bold is a fill-colored outline stroke (Han 1% em, otherwise 2.5% em, the existing Native OFD geometry); oblique is a 0.2 shear about the baseline. Advances do not change; ink includes stroke and shear. Output: `Stroke`/`LineWidth`/`StrokeColor` and a shear `CTM`, without `Weight`/`Italic`, so viewers cannot apply a second synthesis. PDF export renders a Han synthetic-bold stroke with the Issue #107 horizontal-only emphasis. |
| Non-BMP | `CodePosition`/`CodeCount` of the ClusterMapV1 profile count Unicode scalars (XML characters), which is the reading of GB/T 33190 "字符" and identical for BMP text. The reader maps them back to UTF-16; PDF semantic text uses the existing supplementary PUA alias. |
| Decorations | Highlight, underline and strikethrough are filled rectangles in the final result, using the style's primary face metrics (`post`, `OS/2`) so a decoration stays level across fallback. Written as PathObjects: highlight before the line's text, lines after. Abutting rectangles of one kind, color and vertical extent on a line are merged into one path, because anti-aliasing viewers (macOS Preview) show a faint seam between separately filled shapes (Issue #115). |
| Letter spacing, AtLeast, hanging indent | Needed by Flow/DOCX parity: per-grapheme spacing, `TextLineHeight.AtLeast` (extra space below the text, as in the compatibility layout), and a negative first-line indent. |
| ZWJ/ZWNJ/VS | Admitted inside grapheme clusters and shaped by HarfBuzz; cmap coverage is not required for them. OFD and SVG round-trip; PDF export fails closed because PDFsharp's text layer cannot keep them (existing `PdfTextControlPolicy`). |
| Vertical, variable, color | Not implemented; still rejected by `TextFontFace` and the admission rules. |

## Compatibility

- Defaults are unchanged; no existing visual baseline is affected unless the switch is set.
- The OFD writer and reader changes are additive: ClusterMapV1 BMP output is byte-identical to
  0.2.0 for the same layout; a shear CTM and producer stroke are accepted only on ClusterMapV1
  output; foreign CGTransforms keep the historical fail-closed profiles.
- `ParagraphLayoutOptions` and `TextStyle` gained immutable `With…` methods instead of new
  constructor parameters, keeping the 0.2.0 constructors binary compatible.
