# Exact shaped-run positioning query

`TextPositionedRun.Create(shapedRun, sizeMillimeters, originX, baselineY, token)`
projects one existing immutable `TextShapedRun` into millimeters. It retains
that exact object, request, face, external context, feature ranges, engine
version, native glyph order and absolute UTF-16 clusters. It calls no shaper.
The resulting `OfdPositionedGlyph` list is immutable; its type reuses the
existing geometry DTO without creating an OFD element or invoking a writer.

This fills a query boundary in research v1.2 §12.4.4. The prelayout engine's
existing request-identity guard is intentional: it shapes its own requests
under its restricted Latin/Han layout profile. Replaying an external shaped
result through that API violates its contract. Constructing Core positioned
DTOs directly instead requires callers to calculate every coordinate. This
new query explicitly accepts the already shaped result, including a contextual
item, without claiming it is a legal or final line.

For scale `s = size / face.UnitsPerEm`, initial design-unit pen `(0, 0)` and
caller origin `(x, baseline)`, each glyph origin is
`(x + (penX + offsetX) * s, baseline - (penY + offsetY) * s)`.
Glyph advances and the final run advance are `(advanceX * s, -advanceY * s)`.
The pen accumulates advances in signed 64-bit design units; offsets never
contribute to final advance. The last advance is retained. Adjacent origin
deltas generally differ from glyph advances. Ink is independently translated
from upward-Y design rectangles to downward-Y millimeter rectangles, then
unioned. No ink yields null; an empty run has no glyphs, no ink and zero advance.

Size is finite and in `(0, 1000]` mm and must yield a representable positive
scale. Input origins are finite within ±1,000,000 mm. Derived glyph origins,
per-glyph and total advances, intermediate pen positions, ink edges and union
dimensions must meet the same finite coordinate budget. Invalid inputs use
`InvalidRequest`; derived budget violations use `BudgetExceeded`. The existing
shaped-run source/glyph expansion bounds still apply. Cancellation is checked
before allocation, for each glyph and before returning; no partial result is
published. This is bounded managed work, with no native interruption claim.

Each face/size is projected separately. The query does not resolve a font or
paint style, split a ligature across sizes, choose legal breaks, reset joining
at final line boundaries, combine bidi runs, paginate or change legacy Flow.
Existing prelayout/mixed writers still own their current arithmetic and
restricted admission rules; migrating them to a common result is future work.
In particular, accepting Arabic shaping input here does not admit Arabic
OFD/PDF/SVG output or complete C3 or the Arabic/Indic end-to-end corpus.

Tests cover signed 2D pens/offsets, nonzero baselines, final advance versus
origin delta, independent ink union, null ink, empty input, contextual request
identity and native order, ligatures/shared mark clusters, non-BMP maximum
source endpoints, separate face UPEMs/sizes, immutable results, cancellation,
and intermediate/ink/output coordinate budgets. The package consumer emits
`r22-positioned-run-query.json`; its constructed 2D oracle is explicitly
separate from actual native font shaping. Candidate packages, independent
geometry/outline checks, fresh diagnostic pages and actual Preview evidence
remain required before a bounded candidate acceptance.
