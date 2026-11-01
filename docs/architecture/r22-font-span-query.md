# Caller-selected font partition and style ownership query

`HarfBuzzTextPrelayout.CreateFontSpanShaping()` queries a complete nonempty
source with a contiguous, caller-selected partition of immutable font faces.
Each `TextFontSpan` also has an integer `StyleKey`. The key is only an opaque
equality declaration supplied by the caller. The query does not parse colors,
font sizes, weights, letter spacing or other paint/style values, and it does
not decide whether two actual style objects deserve the same key.

```csharp
var request = new TextFontSpanShapingRequest("ffi", 20, new[] {
    new TextFontSpan(20, 1, face, styleKey: 7),
    new TextFontSpan(21, 2, face, styleKey: 7)
});
var result = HarfBuzzTextPrelayout.CreateFontSpanShaping().Shape(request);
// One font item; the original ffi glyph has the unique declared key 7.
```

The fixed original DejaVu fixture demonstrates the need for coalescing:
complete `ffi` produces original GID 5044/advance 1980, whereas independently
shaped `f` and `fi` produce 73+5042/advance 2011. Adjacent spans with equal
`FaceSha256` therefore shape together, even when their display names or
`TextFontFace` instances differ. Font names never select an implicit host face.
Every input span retains its exact original face object in the immutable
request. A coalesced item's shaping request uses its first span's face; all
coalesced spans have the same selected program bytes.

Direction, script, language and source-ranged features are uniform caller
inputs. Feature intervals are relative to the complete source. This is not
automatic font fallback or script/bidi/style itemization: the caller already
chooses every face and interval. Existing compatibility Flow still has its
existing leading-scalar style policy; this query does not change Flow.

Every font item uses the [contextual item query](r22-contextual-item-query.md),
with the complete original source and that item's selected face. Only the
item's glyphs are emitted. A face is not rejected for lacking characters
owned by another font in the surrounding context. Actual item shaping still
fails on unmapped output glyphs. Font switches must lie on complete-source
.NET text-element boundaries, using the contextual factory's documented
runtime segmentation contract. A later switch inside a base/mark sequence
rejects before an earlier valid item can enter the shaper.

`IContextualTextShaper` is required explicitly by the engine constructor.
Every group, source partition, font boundary and call/work budget is checked
before the first contextual shaping call. A replacement shaper must return
the exact item request; retaining only equivalent text or another contextual
request is insufficient.

Groups are returned in **logical source order**. Glyph occurrences inside
each group retain the native order, original GIDs, source clusters, advances,
offsets, ink and break flags. Coverage slice indices are local to the group;
`FirstSpanIndex` relates them to the original request's source spans. Different
groups keep their own face design units/UPEM. There is no combined advance,
global visual-line order, line origin, caret or pagination result. For example,
Noto `A` at UPEM 1000 and DejaVu `V` at UPEM 2048 remain distinct items.

Every glyph receives exactly one `GlyphStyleKeys` entry only after **all** its
source slices agree. If `f|fi` has keys 7 and 9, the shared ffi occurrence
rejects with `InvalidRequest` and its absolute source cluster interval. The
query never chooses the leading style, duplicates a glyph, divides an advance
or automatically disables ligatures to force an assignment. An explicit
caller `liga=0` succeeds only if the actual native output now has compatible
separate clusters. It cannot force combining clusters to split. Same-key
base/mark clusters retain every glyph occurrence, including zero advances.
Arabic joined forms with separate source clusters may carry separate keys
without shaping each letter independently.

Finite limits apply to the complete request:

- 4096 UTF-16 units, 256 positive source spans and 256 feature intervals.
- Eight distinct selected program hashes and 64 MiB total selected face bytes.
- At most 256 shaping calls, optionally reduced by `MaximumShapeCalls`.
- 512 MiB cumulative work, charging two selected-face copies and 64 times the
  complete context length per item, checked before the first call.
- 16384 total glyph occurrences and 32768 total source slices. Repeated M:N
  cluster occurrences consume their own budget, across all items. The existing
  per-item glyph/slice limits also apply. A provisional item's associations
  may be created before an aggregate limit refuses it; no partial result is
  returned.

Cancellation is checked through planning, shaping and association. As with
the existing adapter, a synchronous native call is not interruptible at an
arbitrary instruction. No request, face, glyph or feature range is mutated.

The package consumer records seven accepted and two refused cases in
`r22-font-span-query.json`, retaining original font/source identities and all
per-item glyph placements, coverage and style keys. Diagnostic glyph-path
pages can apply an external illustrative palette to these keys; that palette
is not a product style resolver. Functional, exact-package and actual-page
acceptance remain separate from general document output support.

No new writer accepts this result type. This candidate does not admit a
general styled line, Flow integration, font discovery, Arabic/Indic output,
native reader semantics or arbitrary imported OFD mappings. PR #78's original
font quantity-counter readability and complete terminal-date Find remain open.
