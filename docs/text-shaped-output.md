# Typed shaped text output (R22, first profile)

`TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs` consumes the
completed R21 prelayout result. It writes one producer-owned OFD `TextCode` per
nonempty line with unchanged logical source and a complete `CGTransform` mapping
to the actual original GIDs. Glyph origins use the shaped pen advances and
offsets, converted once from font units to page millimeters. The public
`OfdPositionedGlyphRun` keeps the absolute UTF-16 source clusters, original
face SHA-256, GIDs, origins, and advances separate from the logical text.

This first output profile accepts a regular, static, horizontal Latin run with
one to eight printable, nonblank ASCII characters per line. Each character must
have one glyph in source order, or a complete two/three-character ligature must
have one glyph. The selected face is embedded and the TrueType subset retains
the actual shaped GIDs in their original slots. The writer rejects a face hash
that does not match the selected embedded resource, imported `SourceXml`, and
unsupported source/cluster/position shapes before writing the output stream.
The older `CreateSinglePage` contract stays strict about nominal one-to-one
ASCII glyphs and full embedding.

The candidate package consumer shapes `AV` with kerning enabled/disabled and
`ffi` with the `liga` feature enabled/disabled. It verifies the second `AV`
origin changes and checks OFD text readback, SVG original-GID paths,
PDF text extraction and rendered PDF pages separately. These checks cover the
bounded samples only. Mixed direction, Arabic/Indic shaping, combining marks,
multi-`TextCode` composition, supplementary characters, vertical output and
general paragraph bidi remain outside this producer profile. Imported OFD
content retains its own source-preservation and R02 acceptance boundaries.

Generated PNG or SVG output is not a macOS Preview or independent-reader
acceptance result. Each visual outlet needs its own evidence before release.
