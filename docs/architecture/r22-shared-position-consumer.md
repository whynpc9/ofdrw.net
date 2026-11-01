# Shared positioning arithmetic in the typed prelayout consumer

`TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs` and
`TextPositionedRun.Create` use the same internal, immutable scalar projection:
size/UPEM, pen plus offset, Y inversion, advances and the next signed 64-bit pen.
The helper makes no admission decisions and constructs no public query result.
Both callers keep their own validation and Core DTO construction. There is no
new public API and no externally supplied positioned-result writer overload.

The strict query retains its input, representable-scale, cancellation, per-glyph,
ink, intermediate-pen and total checks in their existing order. Computing the
next pen does not adopt it before the current glyph's ink checks. Y advance is
converted to double before negation, including `int.MinValue`.

The typed writer retains its existing page, source, face, profile and backend
guards. Its admission remains a short regular horizontal Latin line: at most
eight source units/glyphs, the existing one-to-one/complete ligature or bounded
combining-pair profiles, and the existing source/GID alias restrictions. It keeps
boundary-adjusted origins, adjacent-origin deltas and its positive zero vertical
advance. Eight signed-int advances sum exactly as either double or signed long;
the change to long does not round the pen in this profile. Floating-point
expression grouping remains unchanged. The legacy `CreateSinglePage`, mixed
writer, Flow and graphics paths are untouched.

These callers deliberately have different contracts. A controlled replacement
shaper can produce three advances of 750000, 750000 and -750000 mm whose offsets
keep glyph origins at 0, 0 and 750000 mm. The existing typed writer can serialize
this model; the strict query rejects the intermediate pen of 1500000 mm. The
shared helper preserves that distinction without a permissive public query or
exception fallback. This is a compatibility example, not native-font or visual
acceptance evidence.

Seven native baseline comparisons already agreed before this change. This work
adopts shared arithmetic in a real consumer; it does not claim to fix a native
geometry defect. Directed regressions compare observable DTO bits against the
old writer formulas, preserve source/GIDs/deltas through serialization/readback,
and keep the compensated-pen distinction. The package E2E adds three real typed
OFD → PDF/SVG/PNG pages covering seven native lines: AV kerning, ffi/fi ligatures,
a combining pair, negative j ink, multiline source offsets, two sizes and two
UPEMs. These outputs require their own visual review.

An independently observed tiny-positive-size issue remains open: existing
`0.###` serialization can write `Size="0"`, including a size with a representable
positive projection scale. Preserving current numeric compatibility does not
endorse that output or repair the formatter. This change does not resolve paint,
weight/italic synthesis, the CJK quantity counter failure, final-line/bidi layout,
pagination, Arabic/Indic writer admission, or the complete R22/C3 scope.
