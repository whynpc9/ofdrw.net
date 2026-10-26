# R08 embedded CFF1 OFD to SVG text profile

This increment admits selected Unicode text backed by a **standalone, static,
regular, monochrome OpenType/CFF1 font** in OFD-to-SVG conversion. The SVG
retains the original `<text>` Unicode and embeds the original OTF as a
`font/otf` data URI under a content-hashed family name. It does not convert a
GID into Unicode, write `CGTransform`, subset CFF, or enable PDF CFF embedding.

Admission bounds the font at 16 MiB, rejects variable/color/CFF2/TrueType
tables, CID-keyed CFF, non-Type-2 CharStrings, synthetic emphasis, symbol cmap,
restricted embedding flags, invalid CFF1 header/INDEX/Top DICT offsets, an
empty CharString, and disagreement between `maxp.numGlyphs` and the CFF1
CharStrings INDEX count. Selected Unicode cmap GIDs and all non-default
format-14 UVS GIDs must be valid CharStrings IDs; zero non-default IDs refuse
admission.
The selected-text GID scan charges worst-case cmap lookup work to the shared
font parsing budget before each lookup, including repeated format-4 lookups.
Normal Unicode coverage and existing `CGTransform` refusals remain active.
This is structural admission; it is not a general Type 2 CharString validator
or a promise of arbitrary CFF font compatibility.

## Fixed acceptance

The unmodified Adobe Source Code Pro Regular OTF (OFL 1.1; SHA-256
`9f9664e2edf6f045c11e774f9bd0be6993971f2544a39061a5ce478b96b051f8`)
contains original GIDs 2 (`A`), 34 (`g`), 298 (`é`), and 998 (`@`). The
sample `A g é @` deliberately uses only that font's covered Latin repertoire;
it makes no CJK claim. OFD save/readback retains the exact OTF and Unicode
`TextCode`. SVG contains the exact OTF data URI and one logical `<text>` node.
The same OFD selected for PDF export fails with a CFF-specific refusal before
writing output. A mutated `maxp` count with the original CFF CharStrings INDEX
also fails SVG conversion without output.

The actual OFD-derived SVG was rendered at 756 × 340 through `rsvg-convert`
with `PANGOCAIRO_BACKEND=fc` and a private Fontconfig alias from its exact
content-hashed family to that same pinned OTF file. `fc-match` resolved the
alias to the fixture path; the file hash was checked. This explicit binding is
needed because the local `resvg` renderer ignores `@font-face` data URIs. An
independent fontTools `SVGPathPen` page, using the pinned font's own advance
widths and glyph paths, was rendered at the same size. The two white-page PNGs
had normalized mean absolute pixel error `0.000116835`. Both pages were
inspected for glyph shapes, the accent, descender, spacing, overlap, and
clipping. This verifies the selected face's rendered geometry in that offline
viewer; a separate browser's embedded-font behavior is not inferred.

| Outlet | Result for this fixed sample |
| --- | --- |
| OFD save/read | Original OTF bytes and Unicode text preserved; CFF remains full, not subset. |
| SVG | Actual Unicode text and original OTF embedded; bounded structural CFF1 profile and page visual checked. |
| PDF | Explicit CFF refusal before output; no CFF font embedding or PDF text extraction claim. |
| Outline | Prior optional Skia original-GID profile remains separate; see [outline acceptance](r08-cff-outline-profile.md). |

The sample does not establish bare Type1, arbitrary CFF, CFF2, variable/color
fonts, foreign `CGTransform`, or independent OFD reader rendering. R02's
glyph/Unicode-positioning semantics remain a separate gate.
