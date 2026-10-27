# R02 explicit CGTransform glyphs: bounded candidate

This source candidate renders **one or more `TextCode` runs with a complete
one-to-one `CGTransform` mapping** of independent, nonblank BMP characters to
original glyph IDs in an embedded, static monochrome TrueType `glyf` face.
The later source-span candidate below adds a separate complete, unpositioned
many-to-one relationship. Its validation gates remain independent.
Within this candidate, `CodePosition` is a zero-based index in the concatenated
decoded runs; each run keeps its own starting point and deltas. `CodeCount` and
`GlyphCount` default to one. Each mapped glyph
must have an outline. A supplied `DeltaX`/`DeltaY` must describe every gap;
absent deltas leave successive glyphs at the same origin. The original
`TextCode`, resource, IDs, and `SourceXml` remain in the OFD model.

SVG paints original-GID outlines as paths and exposes the logical text in
`aria-label`, `data-ofd-text`, and `title`. It does not add duplicate visible
Unicode text. PDF embeds a face with a narrow Unicode-to-original-GID alias,
paints the mapped glyphs, and keeps extractable logical text. PDF embedding
requires an OS/2 table with an admitted embedding policy; SVG paths do not
embed the font. Missing glyphs, ambiguous GID-to-Unicode mapping, unsupported
positioning, and unsupported face types fail explicitly. Many-to-one or
one-to-many, non-BMP, combining, CFF, color/variable, and
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

## Multi-`TextCode` source increment after PR #57

The [published GB/T 33190-2016 text](https://tsgdag.jhun.edu.cn/_upload/article/files/71/fc/4ed26d5841caab5bc999e19f8f14/02a71fcf-7b6e-4661-9991-eb309c7614c6.pdf),
section 11.4, printed pages 68–70 (PDF pages 73–75), defines a zero-based
`CodePosition`, positive `CodeCount` and `GlyphCount` with defaults of one, and
the four mapping cardinalities. The examined PDF is 5,231,971 bytes with
SHA-256 `d0232b4d75340b885a5e068df477ad2b9f77e4fd23ee3c7e6ddc51b147b27ae8`.
That clause does not expressly settle how positions count across multiple
sibling `TextCode` nodes or how supplementary characters count. The globally
ascending BMP interpretation here is a bounded candidate, consistent with
the [zc310/ofd creator contract](https://pkg.go.dev/github.com/zc310/ofd@v0.1.5/pkg/creator),
and must not be presented as explicit wording in the standard. Position resets,
supplementary code points, and all non-one-to-one cardinalities still refuse.

On base `c71165ecc2c005ed92b60710ccfda0b30cb8f9d1`, focused Release
`ExplicitGlyphTransformTests` passed 15/15 using .NET SDK 10.0.401. A
deterministic two-run OFD was read back with text `ABCD`, exported to PDF and
SVG, and both one-page outputs were rasterized and visually checked. The
original-GID shapes read as two rows, `B A` and `D C`, with distinct run starts
and no observed overlap or missing shapes; PDF letter extraction returned
`ABCD`. The exact OFD/PDF/SVG and page PNGs are in ignored
`artifacts/r02-cgtransform/multicode-run/output/`. This is a synthetic
single-page PNG check only; macOS Preview, an independent reader for this
sample, an unmodified foreign file across all outlets, full regression, and
package-consumer verification have not been completed for this new source head.

The later exact PR #58 gate completed those bounded checks: 11 Release test
assemblies passed 4,267 tests with 17 skips and no failures; one fresh 12-package
feed passed package validation and converter E2E; a private PackageReference
consumer retained the source text and original GIDs. The same OFD's package
PDF matched the independent reader's PDF at 144 dpi with zero changed pixels.
That reader's SVG ignored the transform and was excluded as a visual oracle.
The package PDF was fresh-opened and checked in macOS Preview on its one page.
The Codex bot remained `running/null` after more than 15 minutes; the merge
used an independently recorded coordinator review, not a clean bot receipt.
PR #58 merged as `564e9cbe29100b904125d576cfc6ed74d91848ad`.

At the PR #58 stage, the remaining 2:1, 1:2, and 2:3 relationships illustrated in section
11.4, the standard's numbers are examples of GIDs in unspecified font files;
they are not reusable outlines. The current projection stores one GID per BMP
character, and the PDF path builds a one-Unicode-character-to-one-GID `cmap`
alias before drawing each character. Those structures cannot render the stated
cardinalities with exact glyph positions and one logical text stream. The next
implementation needs a source-character-span to original-GID-sequence model,
per-glyph positioning, and PDF Unicode mapping or marked-content semantics
verified by independent extraction and a reader. No source-backed foreign OFD
with those cardinalities and a redistributable matching embedded font has yet
been identified. The existing Java sample has only equal `CodeCount` and
`GlyphCount`, invalid zero layer IDs, and a font without OS/2; neither identity
nor embedding checks are weakened to make it pass.

## Complete source span to one original GID: next candidate

On merged base `564e9cbe29100b904125d576cfc6ed74d91848ad`, the next
candidate models a `CGTransform` source-character span and its original-GID
sequence explicitly. Its admitted many-to-one subset has **one `TextCode`, one
`CGTransform`, two to eight independent nonblank BMP characters, `CodePosition=0`,
`CodeCount` equal to the complete decoded source span, `GlyphCount=1`, one
nonzero original GID, and no `DeltaX`/`DeltaY`**. The unchanged one-to-one
profile still admits multiple globally indexed `TextCode` runs. Mixed clusters,
extra glyph slots, positioned many-to-one clusters, supplementary scalars, and
unknown font formats still refuse. Neither source IDs nor the PDF OS/2
embedding policy are relaxed.

The OFD keeps its original source XML and logical Unicode sequence. SVG and
HTML preview paint one original-GID outline path and expose the whole source
span as accessibility text. PDF draws one visible CID selecting the original
GID through a private cmap alias, then installs a `ToUnicode` mapping from
that CID to the exact UTF-16 source sequence. Aliases with incompatible
Unicode meanings for the same registered font identity refuse before output;
the PDF stroke overlay also refuses until it has a source-GID paint route.

The focused Release tests on the dirty implementation passed 20/20 using the
pinned OFL Noto Sans `fl` (GID 1653) and `ffi` (GID 1654), and a second pinned
OFL Noto Serif `fl` (GID 2135). They checked OFD readback, one SVG/HTML path,
PDF visible CID, exact `ToUnicode`, extracted source text, and embedded
original-GID outline. `fl` with a positional delta and a same-alias GID with
two different source sequences refused with zero caller output. Generated PDF
and SVG pages were rasterized and visually inspected for these small samples;
the actual viewer and final package gates remain separate.

The independent zc310/ofd v0.1.5 converter is useful only as a limited
semantic comparison for the derived synthetic `fl` OFD: its PDF embeds a
remapped Noto Sans ligature glyph and maps an invisible CID to `fl`, while its
visible page paints separate `f` and `l` paths. Its SVG also paints separate
characters. Neither output is an original-GID visual oracle for this
many-to-one case. The exact input, font and decoded PDF-stream findings are
recorded in ignored `artifacts/r02-cgtransform/next-cardinality/`.

This is not full R02 acceptance. One-to-many, many-to-many, local position
resets, non-BMP counting, general foreign OFD input, and complete reader
interoperability remain open. Full Release regression, candidate package
consumption, exact-head CI/review, and final macOS Preview are separately
required for this new source head.

## One source character to two original GIDs: bounded next candidate

After PR #59 merged as `4a6a7a4e0eca1b921f17f51d7bcd31c4ae1f81e1`,
this candidate admits one independent, nonblank BMP character in one
unpositioned `TextCode`, with one complete `CGTransform` at `CodePosition=0`,
`CodeCount=1`, `GlyphCount=2`, and two distinct nonempty original GIDs. The
embedded static TrueType must permit PDF embedding and cover the source
character for the PDF semantic layer. The admission rule is based on the OFD
mapping and font properties; it does not select Thai or fixed GID numbers.
Additional characters/clusters, sibling `TextCode` nodes, positioning deltas,
and unsupported font formats still refuse before output.

The primary fixture is an unmodified SIL OFL Noto Sans Thai static face pinned
by `e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/upstream-manifest.json`:
U+0E33 maps in its OFD source to original GIDs 59 and 86. Independent
HarfBuzz shaping produces that pair at a shared origin; the font's direct
Unicode cmap for U+0E33 instead points at GID 90. A second Noto Sans synthetic
fixture maps `A` to `B` and `C` original outlines; its first GID has a nonzero
natural advance, yet both exported outlines use the OFD origin. The second
case checks that glyph positioning does not silently follow font advances.

SVG/HTML paint two source-GID paths under one logical label. PDF paints those
same source outlines as paths, plus one clipped, invisible Unicode text glyph
from the embedded font for selection/search. This leaves exactly one extracted
source character without a visible private-use fallback. A trial with two
visible text CIDs and `/ActualText` was rejected because PdfPig extracted an
additional U+E001 from the second CID. The path-plus-semantic route passed
focused Release 35/35 `ExplicitGlyphTransformTests`; both PdfPig and Poppler
`pdftotext` extract only U+0E33 for the Thai page and only `A` for the
nonzero-advance page. Their PDF and SVG pages rendered at 144 dpi and were
visually inspected; pixel differences were limited to roughly 10 and 21
antialias-edge pixels, respectively. The exact focused outputs are in ignored
`artifacts/r02-cgtransform/next-glyph-sequence/focused/`.

This remains a local candidate. Independent OFD reader comparison, full
Release regression, package consumer evidence, current-head CI/review, and
fresh macOS Preview selection/visual inspection are separate gates. The
broader 1:N, many-to-many, mixed, positioned and foreign-file cases remain
outside this profile; one bounded 1:2 case does not close full R02.
