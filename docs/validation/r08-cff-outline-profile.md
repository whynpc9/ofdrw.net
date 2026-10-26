# R08 static CFF1 original-GID outline profile

The optional `Ofdrw.Net.Graphics.SkiaSharp` backend now has an internal,
operation-owned reader for one **original GID** in a standalone static,
monochrome OpenType/CFF1 face. It snapshots unhinted cubic path commands in
design units, the `hmtx` advance, units per em, and the original source SHA-256.
It does not infer Unicode from GID, shape text, or alter OFD `TextCode` and
`CGTransform` source XML. The CFF route remains separate from the existing
managed TrueType `glyf` reader.

Admission requires `OTTO`, a CFF1 table, valid `head`/`maxp`/`hhea`/`hmtx`,
and an exact original stream from Skia. The source is at most 16 MiB; a
returned glyph has at most 32,768 path commands and finite coordinates within
one million design units. Invalid GIDs, collections, CFF2, variable or color
tables, inconsistent native CFF CharString count, and non-cubic native path
verbs fail closed. A null or empty native path is ambiguous between an empty
glyph and an unavailable outline, so this first profile rejects both,
including space.
Cancellation is checked
before and after the native glyph call and during path enumeration; it cannot
interrupt an in-progress native Skia call. The caller still must ensure font
license rights and isolate untrusted native-font parsing as appropriate.

## Fixed source and independent oracle

The unmodified Adobe `SourceCodePro-Regular.otf` and its adjacent OFL 1.1
license come from `adobe-fonts/source-code-pro` `release` commit
`803b7e23ec97ae58b6232ea76519a76d428ba268`. The font is 131,128 bytes,
SHA-256 `9f9664e2edf6f045c11e774f9bd0be6993971f2544a39061a5ce478b96b051f8`.
Its `OS/2.fsType` is zero. The original OTF is Git LFS tracked; see the
[fixture manifest](../../tests/Ofdrw.Net.Graphics.SkiaSharp.Tests/testdata/cff/README.md).
No modified font using the reserved name is created.

Development-only fontTools 4.60.1 `RecordingPen` independently read the
original CFF charstrings. Each digest is SHA-256 of all path commands as one
line per command: `M:x,y`, `L:x,y`, `C:c1x,c1y:c2x,c2y:x,y`, or `Z`, with
coordinates formatted to three decimals and a final newline. The focused
regressions compare every command from Skia, not just glyph bounds:

| Glyph | Original GID | Commands | Advance | Oracle SHA-256 |
| --- | ---: | ---: | ---: | --- |
| `A` | 2 | 16 | 600 | `246cf95385f4ca91e6048151b21fce9f907b5afd2fc3d9af0cd23212e7df41a3` |
| `g` | 34 | 36 | 600 | `69931ee0a24a958361e221c911c3da6768046f405bf69da791267754c7e97a1a` |
| `é` | 298 | 21 | 600 | `3a69a6e9deda4131fd49479b389f30942c3ddad14bbf8a56cc20fc768cdd30ea` |
| `@` | 998 | 25 | 600 | `b8c5eb3e26740cc366973e0e86120ac2db1f1f759e150d7aa43d437ebe666292` |

The focused test can write its actual outline page when
`OFDRW_R08_CFF_ARTIFACT_DIR` is set. Rendering that SVG and the independent
fontTools `SVGPathPen` page through `rsvg-convert` produced byte-identical
600 × 220 PNGs (ImageMagick absolute error 0). The actual PNG was inspected
page 1/1 for complete shapes, accent, descender, overlap, and clipping. The
source baseline, input/output hashes, and distinct visual gates are in the
ignored R08 acceptance record. The final `0.1.0-r08.cff.1` local feed has 12
validated default packages and the optional Skia package. The existing
PackageReference-only Skia consumer passes; a separate fresh-cache consumer
loads the packaged CFF implementation and matches all four oracle digests.
Its actual SVG and PNG are byte-identical to the source page. The package
diagnostic PDF's one page was inspected in macOS Preview at normal and zoomed
views, then its task window was closed. This visual result applies to the
diagnostic outline page, not to CFF text export from OFD.

The count-mismatch regression reproduced a previously accepted nonexistent
GID; the admission guard now rejects it. The exact final source passed the
Release solution test run across 11 projects: 4,210 passed, 13 known skips,
zero failures. The Skia test project contributed 51/51 passes.

## Four outlets

| Outlet | Current boundary |
| --- | --- |
| OFD save | Existing original font and logical text preservation; this increment does not add typed `CGTransform` writing. |
| SVG | The diagnostic page proves only the new outline snapshot. OFD CFF text and direct GID SVG export are not accepted by this result. |
| PDF | CFF font embedding and direct GID PDF export remain outside the proven profile. A diagnostic SVG converted to PDF does not establish OFD→PDF compatibility. |
| Outline | The optional Skia backend reads the four fixed CFF1 original GIDs with independent cubic command agreement. Arbitrary CFF, bare Type1, and other fonts are not covered. |

R02 still needs a licensed foreign `CGTransform` document, source GID and
Unicode-sequence/positioning semantics, and independent reader acceptance.
