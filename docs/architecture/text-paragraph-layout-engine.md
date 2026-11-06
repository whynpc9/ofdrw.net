# Paragraph layout engine (Issue #109, F01–F05)

Status: released in 0.2.0. The engine switch and the F08 admission changes (tabs, ZWJ/ZWNJ and
variation selectors, synthetic styles, non-BMP output, CFF faces) are recorded in
[text engine switch](text-engine-switch-and-f08.md); the admission notes below describe 0.2.0. Reference:
[paragraph layout engine](../content/reference/text-paragraph-layout.md). Validation:
[paragraph layout validation](../validation/text-paragraph-layout-validation.md).

This design implements the paragraph layout pipeline
`font resolution/fallback → paragraph analysis → run shaping → final lines and pages → OFD → exports`
as one new entry point. Earlier bounded queries (R21 prelayout, mixed paragraph, span/font/size
queries) and Flow keep their behavior and admission rules.

## F01 — Unicode algorithms and admission

Decision: a managed, dependency-free implementation of UAX #9 rev 51, UAX #14 rev 55 and UAX #29
rev 47 on Unicode 17.0.0 data, without tailoring.

- Data: `scripts/unicode/generate_unicode_tables.py` reads DerivedBidiClass, BidiBrackets,
  LineBreak, EastAsianWidth, DerivedGeneralCategory, emoji-data (Extended_Pictographic),
  GraphemeBreakProperty, DerivedCoreProperties (InCB), Scripts and PropertyValueAliases. Every
  input is checked against the SHA-256 in `ucd-17.0.0-manifest.json` before parsing. `@missing`
  defaults are applied first. The generated `UnicodeData.g.cs` (range tables, about 125 KB) is
  deterministic; regenerating with the same inputs is byte-identical. Unicode License v3 applies
  (`docs/licenses/UNICODE-LICENSE.txt`, shipped with every package).
- Line break flags needed by tailorable rules are precomputed per code point: Pi/Pf for QU
  (LB15a/b, LB19), East Asian F/W/H (LB19a, LB30), unassigned Extended_Pictographic (LB30b),
  Mn/Mc for SA resolution (LB1).
- Rules map one-to-one to code: `LineBreaker.Decide` lists LB4–LB31 in specification order;
  `GraphemeBreaker.NoBreak` lists GB3–GB999; `BidiAlgorithm` follows the rule numbers.
- Conformance: the complete official test files (compressed in LFS under
  `tests/Ofdrw.Net.Layout.Tests/testdata/unicode-17.0.0/`, decompressed SHA-256 checked) run in CI:
  BidiTest 770,241, BidiCharacterTest 91,707, LineBreakTest 19,338 and GraphemeBreakTest 766 cases,
  all passing. The algorithms are complete through L2; the layout entry adds a narrower admission
  profile (below) and that profile is what is claimed for output.
- Why not ICU/.NET: `StringInfo`, OS segmenters and HarfBuzz cluster levels do not give a
  versioned, testable UAX conformance claim, and ICU would add a native dependency.
- Admission: explicit directional formatting characters, other Cf controls (ZWJ, ZWNJ, SHY, WJ,
  BOM), variation selectors, tabs and unadmitted controls are rejected with their source interval.
  The bidi implementation itself supports explicit controls (they are exercised by the tests);
  admitting them for output is separate work.

EGC, UAX #14 opportunities and shaper clusters are distinct. Lines end only at positions that are
both a UAX #14 opportunity and a grapheme boundary; emergency breaking (opt-in) uses grapheme
boundaries only; output clusters come from the shaper.

## F02 — final result contract

`ParagraphLayoutResult` → pages → lines → runs → glyphs, immutable, all in millimeters with Y down,
design units converted once. Every run keeps the exact `TextShapedRun` (face bytes identity, GIDs,
clusters, script, language, direction, features); every glyph its absolute many-to-many cluster.
`TextResolvedStyle` is a snapshot of paint (fill/stroke color + alpha in DeviceRGB, stroke width),
size, language, requested vs. actual face bold/italic and `Synthetic = None`. Styles are immutable
classes, so later caller changes cannot alter a result. Budgets (text length, lines, shape calls,
coordinates) and cancellation fail without a partial result. New error: `TextLayoutError.LayoutOverflow`.

A grapheme cluster cannot carry two styles: a span boundary inside a cluster is rejected, which also
rules out conflicting paint inside a ligature formed from one cluster. Paint changes at cluster
boundaries split runs, and those runs are shaped separately.

## F03 — paragraph layout

See the reference for the pipeline. Key decisions:

- Font fallback is per grapheme cluster over explicit families; common-script clusters stay in the
  previous face when covered. Exact bold/italic face match only.
- Measurement uses item shaping; the chosen line is rebuilt and runs cut at line edges or at L1
  level changes are reshaped, then the real width decides (bounded back-off). This is what keeps
  Arabic joining correct at line ends (a broken word is shaped as the line sees it).
- Line metrics: shared baseline, max ascent/descent (+ half line gap) per run face and size.
- Justification expands word spaces and gaps next to Han/kana/Hangul/Bopomofo; Arabic kashida is
  not implemented.
- Pagination is line-granular; no widow/orphan control.

## F04 — writer

`ParagraphLayoutOfdWriter` consumes the result directly. The new `OfdPositionedGlyphMapping.ClusterMap`
writes one CGTransform per source cluster in logical order with explicit origins for every glyph
slot. This keeps logical text in TextCode (search/copy semantics) and makes the drawn glyphs fully
explicit, so neither our reader nor other readers must infer GIDs from a cmap or reorder text.

`CodePosition`/`CodeCount` count characters of BMP-only TextCodes, where UTF-16 units and code
points coincide; supplementary characters are rejected by the writer until a counting profile is
verified (F08).

Reader side: historical CGTransform profiles keep precedence for anything they accept, so existing
outputs are unchanged. The writer marks its TextObjects with `ClusterMapV1="true"`
(`https://ofdrw.net/layout-hints`, registered as a known writer-owned hint like
`FauxItalicMatrixV1`). Only a marked TextObject that the historical profiles reject is reconsidered,
as a complete cluster map needing no cmap inference. A structural fallback without the marker was
tried first and rejected: it accepted foreign complete-cover inputs (for example ambiguous or
miscounted Hebrew visual-order GIDs) that existing tests deliberately fail closed. PDF/SVG export draw such runs from original outlines and attach the
logical text invisibly (PDF) or as `title`/`aria-label` (SVG). The embedding plan accepts empty
outlines only for whitespace clusters of cluster-mapped runs and keeps actual GIDs when subsetting.
`TrueTypeGlyphOutlines` now accepts faces up to 32 MiB (the `TextFontFace` budget), needed for
full CJK fonts such as LXGW WenKai.

Not written: stroke, synthetic styles, non-BMP text. Existing writers (`TextPrelayoutOfdWriter`,
`TextMixedParagraphOfdWriter`) and their bounded profiles are unchanged.

## F05 — scripts and exports covered

Verified with deterministic fixtures already in the repository: Latin (Noto Sans), Han
(LXGW WenKai), Hebrew and Arabic (DejaVu Sans 2.37 derivative) and Tamil as the selected Indic
script (Noto Sans Tamil). Checks: GSUB joining forms differ from cmap glyphs; line-end Arabic
words equal their standalone shaping; Tamil many-to-many clusters stay inside graphemes; OFD
round-trip of text, GIDs and cluster mapping; PDF/SVG generation; page PNG review. PDF
Copy/Find, SVG accessibility and macOS Preview acceptance are not verified in this change.

## F06, F07, F08

- F06 (PDF→OFD RTL semantic import) is not part of this change; PR79's rejection of unverified RTL
  semantic import stays.
- F07 (archival coverage) is excluded by the requester.
- F08: tabs, ZWJ/ZWNJ, variation selectors, vertical text, variable/CFF/color fonts and non-BMP
  output remain rejected by this profile and need their own items.
