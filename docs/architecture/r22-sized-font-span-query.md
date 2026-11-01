# Caller-selected face and size query

`HarfBuzzTextPrelayout.CreateSizedFontSpanShaping()` queries a nonempty source
partition with explicit immutable font faces and em sizes in millimeters.
`TextSizedFontSpan` retains the exact existing `TextFontSpan` and adds its size.
The original source interval, face and opaque `StyleKey` are unchanged. This
query never infers a size, color, weight or font program from a style key.

```csharp
var request = new TextSizedFontSpanShapingRequest("AV", 20, new[] {
    new TextSizedFontSpan(new TextFontSpan(20, 1, face, 7), 4),
    new TextSizedFontSpan(new TextFontSpan(21, 1, face, 9), 6)
});
var result = HarfBuzzTextPrelayout.CreateSizedFontSpanShaping().Shape(request);
// Two source items, each with its exact shaping result and size projection.
```

This is an explicit query policy: a face hash or numeric size change starts a
new shaping item. Equal-face/equal-size spans coalesce even when their face
objects, display names or opaque keys differ. A glyph whose source slices
have conflicting keys still refuses; the query does not pick a leading key.
Compatible base/mark spans can therefore retain shared occurrences and zero
advances, while a size change inside a complete-source .NET text element
refuses before any shaping call. All original face/size boundaries are checked,
including those outside a selected range.

With the fixed DejaVu fixture, equal-size `AV` retains its native kerning and
total advance 2671 design units. Explicit `A` 4 mm / `V` 6 mm produces two
contextual items with advance 1401 each: their separate advances sum to
6.8408203125 mm. Scaling glyphs from the complete kerned run by two external
size labels instead gives 6.5849609375 mm. These distinct policies must not be
confused. Similarly `f` 4 mm / `fi` 6 mm yields original GIDs 73 and 5042;
equal-size `ffi` coalesces to GID 5044 when ownership agrees.

Every item is freshly shaped with the complete original Unicode, direction,
script, language, selected face and ordered feature intervals retained by the
existing contextual query. `CreateContextualItem(original, range)` retains the
exact original sized request, clips source spans with their sizes/keys/faces,
and preserves unchanged span objects. Nested selected requests refuse. Features
are intersected and rebased using the existing contextual factories; the full
original feature sequence remains in native source context. The pinned
[HarfBuzz 14.2.1 UTF-16 implementation](https://github.com/harfbuzz/harfbuzz/blob/14.2.1/src/hb-buffer.cc)
receives the complete source with an item offset/length and retains bounded
native pre/post context. Retaining the complete managed request is not a claim
of unlimited native context. Size itemization here does not alter native font
scale, PPEm or hinting.

Results remain in logical source order. Each item preserves native glyph
order, original GIDs, absolute UTF-16 clusters, design-unit advances, offsets,
ink, flags, coverage slices and per-occurrence style keys. Its `PositionedRun`
reuses [the exact shaped-run projection](r22-positioned-run-query.md) at that
item's declared size and origin/baseline `(0,0)`, retaining the exact shaped
object. Each face uses its own UPEM. Separate Arabic items retain joining
context and their internal native order; this is no combined RTL visual line.
There is no aggregate positioned line, line advance, caret or pagination API.

The existing font partition engine supplies shared planning and global guards:
4096 UTF-16 units, 1–256 spans, 256 features, eight distinct faces, 64 MiB total
selected face bytes, and 1–256 calls (default 256). Each planned item charges
`2 * faceBytes + 64 * completeOriginalUtf16Length` against a cumulative 512 MiB
work budget before any native call. Selected ranges do not reset this charge.
Outputs share limits of 16384 glyph occurrences and 32768 source slices across
all items. Each millimeter projection additionally enforces its coordinate
budget. Sizes must be finite, positive and at most 1000 mm, with a representable
positive size/UPEM scale. Cancellation and any later failure return no partial
result; synchronous native calls retain their existing interruption limits.

Directed tests cover coalescing, size switches, AV/ffi, mark boundaries,
ordered overlapping features through selected offsets, Arabic joining,
different UPEMs, non-BMP maximum source endpoints, exact object retention,
immutable snapshots, bounded hostile collections, global work/call/glyph/slice
limits, projection failures and cancellation. Both isolated package drivers
copy the same probe, emitting 11 accepted, five refused and one canceled case
in `r22-sized-font-span-query.json`. Source tests, exact package identities,
independent geometry checks, fresh diagnostic pages, actual Preview, CI and
final-main evidence are separate gates; this document does not certify them.

No writer consumes this result. Automatic font/style analysis, synthetic
weight, color/paint interpretation, final-line joining/bidi/line breaks,
legacy Flow migration, general pagination, Arabic/Indic output and imported
OFD mapping profiles remain open. The original-font quantity counter and
terminal-date Find issues remain separate; this query does not resolve them.
