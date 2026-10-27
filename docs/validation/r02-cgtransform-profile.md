# R02 explicit CGTransform glyphs: bounded candidate

This source candidate renders a **single `TextCode` with a complete one-to-one
`CGTransform` mapping** of independent, nonblank BMP characters to original
glyph IDs in an embedded, static monochrome TrueType `glyf` face. `CodePosition`
is zero based; `CodeCount` and `GlyphCount` default to one. Each mapped glyph
must have an outline. A supplied `DeltaX`/`DeltaY` must describe every gap;
absent deltas leave successive glyphs at the same origin. The original
`TextCode`, resource, IDs, and `SourceXml` remain in the OFD model.

SVG paints original-GID outlines as paths and exposes the logical text in
`aria-label`, `data-ofd-text`, and `title`. It does not add duplicate visible
Unicode text. PDF embeds a face with a narrow Unicode-to-original-GID alias,
paints the mapped glyphs, and keeps extractable logical text. PDF embedding
requires an OS/2 table with an admitted embedding policy; SVG paths do not
embed the font. Missing glyphs, ambiguous GID-to-Unicode mapping, unsupported
positioning, and unsupported face types fail explicitly. Cross-`TextCode`,
many-to-one or one-to-many, non-BMP, combining, CFF, color/variable, and
synthetic emphasis remain outside this profile.

OFD hexadecimal escapes in the source `TextCode` are decoded for rendering and
logical extraction; the raw `SourceXml` is retained. HTML preview admits the
generated original-GID paths and their literal accessibility text. It reserves
outline generation capacity before creating paths, and refuses output when its
configured character budget is exceeded.

PDF positioning for this profile uses the validated character/GID slots. Two
characters that form one Unicode grapheme, such as a Hangul Jamo pair, retain
their separate original GIDs and the supplied inter-slot `DeltaX`/`DeltaY`.
Ordinary text without `CGTransform` keeps its grapheme-based positioning.
Compatible character-to-GID mappings using the same source face and style are
gathered across selected pages before PDF drawing and share one font alias.
Conflicting mappings use separate aliases. Selected aliases pass through the
complete-face FontFile2 restoration path, preserving nested composite outlines.

## Candidate evidence on base `889a31c61ca9ba5d07f19456b362a419e10d8434`

- Focused `ExplicitGlyphTransformTests`: 8/8 after review remediation. The deterministic OFL Noto Sans
  source swaps the GIDs for logical `AB`; PDF extraction remains `AB`, its
  embedded outlines match the original GIDs, and SVG paths name and draw the
  original GIDs. Negative tests cover missing and ambiguous GIDs, incomplete
  spans/deltas, and output-empty refusal. An externally shaped cmap-free OFD
  also converts to PDF and SVG by original GID. Escaped source characters and
  literal backslashes retain their decoded logical text through PDF/SVG;
  normal HTML preview renders the paths while an exceeded outline budget
  refuses output before writing to the caller.
- Existing PDF/PageBlock focused tests: 6/6. R08 original-GID outline tests:
  10/10. These runs precede the full regression and package-consumer gates.
- Synthetic source `artifacts/r02-cgtransform/first-candidate/explicit-gid-synthetic.ofd`
  was converted to PDF and SVG. The actual one-page PDF
  (`SHA-256 0c260fbe8ca89d1dfd9af4d8a10b6e17134bf73eaaeda890a6a67205526ec78c`)
  was reopened in macOS Preview after regeneration: original-GID **B A**
  visibly occupies the two expected positions, without clipping, overlap,
  missing glyphs, or duplicate visible text; Preview selection reports logical
  `A B`. The task's Preview window was closed after inspection. The page-one
  PDF and SVG PNGs were also inspected. This is one synthetic page only.
- An external Java OFDRW repository sample `doc1File.ofd` (`SHA-256
  92a655b3966aae0b44dc9ca666849f68ee26028e12f537669e58489271f06f95`)
  has original-GID mappings on eleven pages. Its `Layer ID="0"` is rejected by
  the existing source-identity loader before CGTransform rendering. For a
  **local diagnostic only**, eleven layer IDs were changed to unused positive
  IDs in an ignored copy (`SHA-256
  c42ac5a091fc1cff29d81eb9c2cc03cb281e5ac985349e0d21cc49254901df3d`).
  Our page-one SVG (`SHA-256
  c9f731f5c56663628510a4b1dfd0c7d76e2df1a66a93840ce26c71e2c9e17de4`)
  and zc310/ofd v0.1.5's PDF from the **same normalized input** (`SHA-256
  2188c76924aae0c40114807f49f84e56125b27c13b34a71a683e2fcafe8e8e1e`)
  were rasterized and visually compared on page one. Both show the same 22
  Chinese lines without apparent missing glyphs or clipping. The independent
  reader's PDF from the unaltered original (`SHA-256
  d8297fc72689d21c6d1eb4f9840f399c93b339ac2094336f4d4e630c5094a382`)
  has identical page-one pixels (144 dpi PNG, absolute-error count zero) and
  identical extracted text to its normalized-input PDF. This comparison covers
  our SVG only; the external sample's embedded font has no OS/2 table, so
  this PDF profile does not admit it. The sample and font remain ignored because
  redistribution rights for the font have not been established.

Full regression, local package consumption, exact-head CI/review, and final
package Preview are separate gates for each new source head. The external diagnostic is one page
of a locally normalized input; it does not establish unmodified foreign-file
acceptance or general CGTransform interoperability.
