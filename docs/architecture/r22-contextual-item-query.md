# Bounded contextual item query

`TextShapeRequest.CreateContextualItem(context, item)` selects one nonempty
absolute UTF-16 interval from an existing complete request. It retains that
exact immutable request in `Context`, including the unmodified full source and
selected face. The item inherits direction, script and language. An ordinary
request still has null `Context`; its existing public constructor is unchanged.

```csharp
var context = new TextShapeRequest("ببب", 100, face,
    TextDirection.RightToLeft, "Arab", "ar");
var item = TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(101, 1));
var result = new HarfBuzzTextShaper().ShapeContextualItem(item);
// Only the middle source character is output; Context still owns all three.
```

The fixed original DejaVu fixture makes the defect observable: the middle
character without context is isolated GID 1366. With both neighbors it is
medial GID 5260; removing the preceding or following neighbor yields initial
5259 or final 5258. All results own exactly the selected source interval.
Changing `SourceStart` on an isolated substring cannot supply these neighbors.

The adapter uses `Buffer.AddUtf16(fullText, itemOffset, itemLength)` with
explicit default flags and the existing monotone-grapheme cluster level.
Native cluster and feature coordinates refer to the full UTF-16 string;
returned clusters are rebased to the full request's absolute source start.
The final cluster ends at the item endpoint, not the context endpoint. Native
output outside the item rejects. Context can influence shaping without
contributing output glyphs. The finite native control with BOT/EOT flags also
retained the medial form; no claim that these flags suppress supplied context
is made. The pinned adapter continues to enforce native version 14.2.1.

Features on `context` use the existing full-run-relative coordinates. The
factory intersects each positive feature range with the item, rebases it to
item-relative coordinates and preserves order. An outside or zero-length
feature has no intersection. The native adapter then translates these item
ranges to full-string cluster coordinates exactly once. Neither request is
modified. Item-relative callers can inspect the resulting `Features`.

Both item endpoints must be boundaries from
`System.Globalization.StringInfo.ParseCombiningCharacters` on the complete
source, or the final source endpoint. This is the current .NET text-element
algorithm of the calling runtime, not a new pinned Unicode segmentation
table, a glyph boundary or a caret promise. Surrogate pairs and tested
base/combining-mark sequences cannot be split. Nested contextual requests
reject; select each item directly from the original complete request. The
full context must first pass the existing 4096-code-unit, valid UTF-16,
control, feature and selected-face validation. Item glyph expansion and
source coverage retain the existing budgets and cancellation checks around
the synchronous native call. No arbitrary native instruction is interruptible.

An item is shaped independently with surrounding context. This is **not**
equivalent to cropping a whole-run result: whole `ffi` produces GID 5044,
whereas item `fi` produces original GID 5042. The existing span query still
shapes compatible complete runs together and reports shared clusters. A
cluster shared by conflicting styles has no default style ownership: callers
must refuse that assignment rather than duplicate the glyph, divide its
advance or select the leading style. This factory does not take style values,
resolve such conflicts, select faces, itemize fonts or promise that independently
queried items can be stitched into a general styled line.

`IContextualTextShaper.ShapeContextualItem` is an explicit capability, separate
from ordinary `ITextShaper.Shape`. The HarfBuzz adapter rejects contextual
requests at its ordinary entry point and ordinary requests at its contextual
entry point. Existing custom shapers do not silently acquire this capability.

`TextSpanShapingEngine` can associate source spans inside a contextual item;
it requires this opt-in and dispatches to the contextual method, refusing an
ordinary-only shaper before calling it. Its result retains the original context through the exact item request. No
new API converts this query result into an OFD/PDF/SVG document. Existing
prelayout and mixed-paragraph engines require the exact requests they create,
and reject a replacement shaper's contextual request before writer admission.
Legacy Flow, pagination, Arabic/Indic output profiles, reader semantics,
general bidi and PR #78's readability/Find gates remain open and unchanged.

Focused regression first observes the unchanged adapter returning isolated
forms for all three connected controls. The repaired path covers both context
sides, first/last RTL items, nonzero and maximum absolute offsets, non-BMP
prefixes, feature intersections, full-range compatibility, combining marks,
invalid boundaries, nesting, cancellation and existing output-engine refusal.
The package consumer emits `r22-contextual-item-query.json` with original
source, item intervals, features, original GIDs and all placements. Diagnostic
glyph-path pages remain query evidence, not script output or search acceptance.
