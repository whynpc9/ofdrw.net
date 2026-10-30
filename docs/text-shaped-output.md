# Typed shaped text output (R22, first profile)

`TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs` consumes the
completed R21 prelayout result. It writes one producer-owned OFD `TextCode` per
nonempty line with unchanged logical source and a complete `CGTransform` mapping
to the actual original GIDs. Glyph origins use the shaped pen advances and
offsets, converted once from font units to page millimeters. The public
`OfdPositionedGlyphRun` keeps the absolute UTF-16 source clusters, original
face SHA-256, GIDs, origins, and advances separate from the logical text.

This output profile accepts a regular, static, horizontal Latin run with
one to eight printable, nonblank ASCII characters per line. Each character must
have one glyph in source order, or a complete two/three-character ligature must
have one glyph. It also admits an ASCII Latin letter followed by one nonspacing
mark in U+0300–U+036F when HarfBuzz returns one complete two-code-unit cluster;
it preserves the source without normalization. Separate mark glyphs, other
combining ranges and longer combining sequences remain unsupported. The selected face is embedded and the TrueType subset retains
the actual shaped GIDs in their original slots. The writer rejects a face hash
that does not match the selected embedded resource, imported `SourceXml`, and
unsupported source/cluster/position shapes before writing the output stream.
One-to-one output requires a one-to-one source-character/GID relationship for
distinct values: repeated GIDs must denote the same character, and repeated
characters must retain the same GID. This keeps PDF ToUnicode, font aliases
and SVG outline semantics unambiguous.
The older `CreateSinglePage` contract stays strict about nominal one-to-one
ASCII glyphs and full embedding.

The candidate package consumer shapes `AV` with kerning enabled/disabled,
`ffi` with the `liga` feature enabled/disabled, decomposed `e` + U+0301 and
decomposed `a` + U+0308.
It verifies the second `AV`
origin changes and checks OFD text readback, SVG original-GID paths,
PDF text extraction and rendered PDF pages separately. These checks cover the
bounded samples only. Mixed direction, Arabic/Indic shaping, general combining marks,
multi-`TextCode` composition, supplementary characters, vertical output and
general paragraph bidi remain outside this producer profile. Imported OFD
content retains its own source-preservation and R02 acceptance boundaries.

Generated PNG or SVG output is not a macOS Preview or independent-reader
acceptance result. Each visual outlet needs its own evidence before release.

The separate mixed writer and its repeated-letter candidate are reconciled in
[the R22 remaining-scope ledger](validation/r22-remaining-scope.md). The Latin
producer contract above does not imply general paragraph or SVG text semantics.
