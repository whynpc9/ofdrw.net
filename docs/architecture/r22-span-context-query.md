# Bounded source-span shaping query

`HarfBuzzTextPrelayout.CreateSpanShaping()` shapes one complete, caller-selected
`TextShapeRequest` once. `TextSpanShapingRequest` adds a contiguous partition of
that same source, using absolute UTF-16 offsets in the run's coordinate system.
It preserves source ownership without shaping each fragment independently.

```csharp
var run = new TextShapeRequest("ffi", 20, face);
var request = new TextSpanShapingRequest(run, new[] {
    new TextSourceSpan(20, 1),
    new TextSourceSpan(21, 2)
});
var result = HarfBuzzTextPrelayout.CreateSpanShaping().Shape(request, token);
// result.ShapedRun retains the whole-run GIDs, positions and total advance.
// A ligature may have one coverage entry containing slices from both spans.
```

The fixed DejaVu fixture demonstrates the need for this contract: whole `ffi`
uses original GID 5044/advance 1980; independently shaped `f` and `fi` use
73+5042/advance 2011. Whole Arabic `ببب` uses final/medial/initial GIDs
5258/5260/5259, while independently shaped letters use isolated GID 1366.
Changing `SourceStart` labels does not supply missing shaping context.
These are finite query controls, not Arabic paragraph/output acceptance.

Every glyph occurrence has its own `TextSpanGlyphCoverage`, in the shaper's
original order. Its slices remain in logical source-span order. RTL glyph
order therefore does not reverse source ownership. Multiple glyphs sharing
the same exact cluster retain every occurrence, original offsets and zero
advances; `ClusterGlyphCount` reports that multiplicity. `CrossesSpanBoundary`
reports a cluster intersecting multiple spans. No unique style, per-span
advance, caret, selection rectangle or rendering decision is inferred.

The request snapshots at most 256 positive source spans. They must cover the
complete run without gaps, overlap, reordering, out-of-range arithmetic or
surrogate-pair splits; empty text requires an empty partition. A span boundary
may lie inside a combining cluster, and the result exposes the shared source
association. The existing run contract still owns its font identity, original
text, features, direction/script/language, glyph/source validation and
4096-code-unit/expansion budgets.

The query requires the shaper to return the exact input request. It counts
all source intersections before allocating output, rejects more than 32768
slices, uses bounded indexed input snapshots and binary span lookup, and
checks cancellation before/after shaping and throughout association. A native
synchronous call cannot be interrupted at arbitrary instructions. No partial
association result is returned on refusal or cancellation.

This is one-face source-association querying. Callers still own compatible
font/style itemization and decisions for a glyph spanning conflicting styles.
It does not add surrounding context outside the supplied full run, general
bidi analysis, font fallback, cross-face shaping, line-level reshaping,
pagination, legacy Flow behavior, writer admission or PDF/SVG semantics.
Broader C3 and the Arabic/Indic end-to-end corpus remain open.

The separate [contextual item query](r22-contextual-item-query.md) retains
surrounding text outside an explicitly selected item. Its independent shaping
is not equivalent to slicing these whole-run glyphs or assigning their styles.
The separate [font partition query candidate](r22-font-span-query.md) adds
caller-selected face grouping and refuses conflicting style-key ownership.
It does not interpret actual paint/size properties or change compatibility Flow.
The separate [positioned-run query](r22-positioned-run-query.md) projects an
exact shaped result at an explicit size without reshaping or writer admission.
The [caller-selected face/size query](r22-sized-font-span-query.md) explicitly
groups compatible sizes before shaping and projects each exact item separately.
It retains original context without resolving paint or composing a visual line.

Focused tests cover actual ffi/feature-off, Arabic and combining behavior,
generic many-to-many/zero-advance occurrences, source snapshots, invalid UTF-16
partitions, integer endpoints, association/span budgets, cancellation and
foreign replacement-shaper requests. The package consumer repeats the fixed
font cases and emits `r22-source-span-query.json`. Candidate-specific full
regression, API, exact packages, diagnostic-page and CI evidence remain
separate acceptance gates.
