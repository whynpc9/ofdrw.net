# Filled text stroke and small counters

The final-main review at `a05badc212a0b653a0a10afc9648e24c50c3f11a`
found light tracks in CI-generated DOCX titles. The first PR78 head
`e84af7d6e437cea2543a55b4b895c8d5ec4e7914` filled the supplementary
contour and improved the title. Fresh review then found almost closed counters
in the small table-header character 量. The actual public OFD→Image export at
its default 144 DPI also reproduces that problem; it is not solely an external
Poppler inspection issue. Both earlier generations remain unaccepted.

The deterministic source is
`e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx` (SHA-256
`17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`).
Inspected title outlines and the CI 量 contour topology match their original
fonts. This does not establish a baseline displacement or validate arbitrary
fonts. Preview, Poppler, PDFium and Ghostscript differ at small display scales;
a clear high-resolution image does not waive ordinary-size readability.

For fully opaque, identical-RGB fill and stroke with a positive width and an
available outline font, PDF export uses one visible original glyph fill and a
filled geometric region for its centered stroke. The semantic font string is
emitted once inside an empty clip; actual Preview selection, copying and Find
must verify that route separately from the earlier Tr3 diagnostic. Mixing a
hinted visible font with its unhinted supplementary contour is avoided in this
route. Other colors, partial alpha, outline-only text and zero-width hairlines
retain their existing paint behavior.

The geometric region unions positively oriented segment rectangles and corner
regions using nonzero fill. Closed contours retain the previous centered width,
miter join and PDF default miter limit 10 with bevel cutoff. Original glyph fill
uses nonzero winding. The geometry is built before the existing affine transform
and clip. Cubic centerlines are subdivided with local tolerance adjusted by the
current transform's largest singular value, bounding that curve approximation
to 0.001 PDF points in page coordinates. This is not a bound on every font or
rasterization error. Subdivision depth is limited to 20 and recorded points to
65,536 per glyph; invalid or exhausted geometry fails before PDF publication.
This PDF paint route adds no native dependency or replacement font. The separate
native DOCX counter analyzer uses managed Clipper2 as described below.

An explicit same-color stroke at a nonhistorical width must also avoid the old
native-pen overlay. With a 1.25%-em stroke and requested weight700, unchanged
`38b464f2` paints about25.22% more integrated opacity than the independent
original-path reference; its weight400 control passes. The two same-color
opaque centered strokes cover their wider region, so the exporter emits that
union once through geometry. The reference retains its0.98–1.02 interval.
Additional cases cover nonhistorical width under affine transform/clip and a
0.5%-em explicit stroke contained by a predetermined1%-em reference stroke.

For positioned mixed Han/Latin text, the PDF contour union uses the whole
TextObject's script-aware supplementary emphasis width. Recomputing from an
individual Latin glyph could raise a counter-constrained native width back to
2.5%. A public native/default DOCX fixture with24pt `量A` preserves its planned
approximately2.04% stroke for both characters; its PDF must match the same
original face, positions and declared stroke with only requested weight cleared.
This does not make arbitrary imported widths authoritative: pure Latin still
requires the2.5% supplementary width, and a mixed span's explicit stroke below
1% remains supplemented to1%.

The original seven title coverage cases retain their2x-versus8x area oracle and
0.98–1.02 interval. New short Latin/mixed path controls use both operands at8x
with the same interval and one-pixel boundary check. Their original2x-versus8x
ratios near0.979 and same2x comparisons remain diagnostic evidence; a passing
8x comparison does not establish low-resolution pixel equality. The short
controls reject the old per-glyph2.5% expression, while the public native fixture
also requires exact pixel equality at4x. Actual-page visual review remains a
separate requirement.

Fresh native DOCX with an embedded regular face can increase synthetic
emphasis towards3.5%-em, rather than limiting every Han-containing span to1%.
The earlier0f36cc3b policy used a2.5% cap and80% relative witness retention;
its fresh visual review still reported weak emphasis. The current optical
policy adds an absolute opening constraint instead of only a relative one. The original face, GIDs, text, advances and requested
weight remain unchanged. The planner inspects original upright static TrueType
outlines at the serialized font size. For each closed counter, it looks for an
empty witness point in that counter, excluding smaller containing counters;
an outer ring cannot borrow the inner white hole in a nested glyph such as回.
It measures a conservative original empty disk, then checks candidate widths
against the actual renderer's segment, miter10 and bevel polygons. Distance to
the original centerline alone is insufficient at a pointed join. An increase
must leave clearance of at least60% of each conservative witnessed radius
and at least0.4pt, plus the existing numerical allowance. That absolute floor
can refuse an increase for a small counter even when a relative rule permits
one. Failure retains the legacy1% Han stroke; that fallback does not promise
to meet the increased-width clearance requirement. The increase is rounded
down to the writer's0.001mm grid. A generated rectangular-counter fixture
checks6/12/18pt behavior analytically without adding a font payload.

Counter candidates now come from the whole original nonzero filled outline's
polygon union, using Clipper2 2.0.0. Classifying only opposite-oriented input
contours misses white regions enclosed by overlapping clockwise pieces. A
generated four-bar control has the same filled opening as the existing explicit
rectangular counter:6pt retains legacy width, while12/18pt must still increase
and match that positive control. The original runtime font/GID and input bytes
remain unchanged. Topology output never replaces the original glyph painting;
witness membership, empty-disk distances and actual miter/bevel clearance still
use original analyzed contours.
The single standard glyph callback must match the bound font's original direct
cmap GID; a substituted/multiple/missing glyph retains legacy width. Cmap
parsing/lookups share the existing64MiB work budget within one DOCX render and
propagate cancellation. Package subsetting may remove unused outlines, but the
selected original glyph bytes/GID must remain, apart from zero alignment padding.

The original pinned 目 has three overlap-built rectangular openings with minimum
height162/1000em. Native and default6pt DOCX now retain0.060mm (about2.83%-em),
rather than the unsafe former0.074mm. The next writer-grid width0.061mm would
already miss the fixed opening allowance. Original cmap/GID/glyph bytes are
checked through packaging, allowing only zero alignment padding; read/write and
PDF export must preserve the declared constrained stroke. The generated100-unit
counter's6pt fallback above is a different geometry and is not used to demand
fallback for every real glyph.

The shared fixture's first section heading includes 格, whose overlapping pieces
enclose counters despite all input contours winding alike. At its selected white
witness, an original straight border bounds clearance below0.4pt for0.168mm,
the smallest writer-grid width satisfying the old3.4%-em lower expectation.
The corrected0.167mm retains0.400679pt at that witness, including the fixed
allowance. That derived fixture expectation changes; the original title pixel
coverage interval, quantity brightness100/upper-bar oracle and0.4pt/60% guards
do not. This is a witness-specific constraint, not proof that every possible
witness would forbid a wider stroke.

Within one native table row whose cells each contain exactly one paragraph,
upright synthetic CJK text with the same embedded
font resource, serialized size and RGBA color uses the smallest already planned
width in that group. This only reduces wider siblings; it never increases the
counter-constrained cell. All automatically wrapped lines of each sole
paragraph participate. A row with multiple paragraphs in any cell retains the
per-object policy; separate rows, outside paragraphs and pages never combine.
Pure Latin, italic, physical-bold, name-only and imported-paint routes
retain their existing contracts. The shared layout fixture checks that each
header still has more ink than its same-font/same-glyph regular reference;
ordinary-scale emphasis and counter readability remain separate visual gates.

Fresh upright, opaque, pure-BMP-Han members of such a group can also receive
an optional nonsemantic original-curve stroke towards3.5%-em. The original
TextObjects, requested Weight, font/GIDs, advances and already constrained base
LineWidth remain. Every additional path is clipped outside all enumerated
counter bounding rectangles in the complete equal-format group, including
counters belonging to other TextObjects or automatically wrapped lines. Each
rectangle expands by0.5pt plus0.0005pt numerical allowance, with its millimeter
coordinates rounded outwards to six decimals. Separate even-odd complements
intersect; overlapping boxes cannot XOR a protected region back into the paint.
This group-wide exclusion matters when adjacent formatted runs are emitted as
separate objects. Protecting only each path's own glyphs would leave a sibling's
counter exposed to the extra stroke.

The added paths retain original closed line/quadratic contours and their joins;
polygon union output is used only for counter enumeration. Original callback
GIDs must match the face's direct cmap. Placement uses the serialized font size,
boundary/run X/Y and producer DeltaX, with the same horizontal-metric baseline
correction as the PDF outline route. The group is constructed completely before
any path reaches the page. Each path follows its own text immediately, preserving
later cell backgrounds/borders and other paint order. Separate formatting groups,
rows, paragraphs and pages remain independent. Mixed Han/Latin, italic, physical
bold, name-only, imported, partial-alpha or unsupported/exhausted groups keep their
base policy. The optional path does not alter the uniform planner's0.4pt/60%
witness guard or make that guard a new guarantee for additional paint.

Per pagination pass, optional placement admits at most64 text objects and256
BMP Han characters per group,2048 glyph placements,131072 original-curve replay
commands,4096 combined counter visits/clip emissions and4MiB of formatted path
text, with512KiB per path stream. Repeated group-mask copies consume the clip
budget. Failure keeps the whole group's base stroke, and cancellation propagates.
Discarded pagination passes get fresh placement budgets; the existing document-wide
font/topology budgets and at-most-eight-pass pagination limit remain. These are
producer limits, not a bound on Clipper internals or arbitrary viewer performance.

The regular contrast reference now comes from independently reconverting a DOCX
whose header runs request regular weight. It must retain the same original font
payloads and text positions. Clearing Weight/Stroke on an OFD that still contains
supplementary paths is not a regular reference. Counter/coverage oracles remain;
more ink than regular is not sufficient visual acceptance. Flattening/quantization
can still affect tiny topology. Exclusion rectangles protect enumerated regions,
not a formal exact-curve arrangement completeness proof. Raw raster changes,
zero-sample holes, ordinary-scale readability and actual reader behavior remain
separate evidence, and failures cannot be erased by the geometric checks.

This is an optical fallback, not a physical700 face or a guarantee about every
open gap, neck, counter component, font or raster scale. Italic, variable,
color, non-TrueType or unavailable analysis retains the existing policy.
Analysis is local to one DOCX render: at most8 faces,32MiB per face/64MiB total,
512 glyph/size measurements,131,072 original/returned-hole points and50million bounded
geometry visits; each glyph admits64 contours/4096 analysis points with curve
depth20. Exhaustion retains the legacy width; cancellation still propagates.
Polygon input uses quadratic flattening at0.00025pt and a1e-6pt integer grid;
these approximations can affect tiny topology and are not an exact-curve proof.
Before the union call, absolute coordinates are limited to10,000pt and each
glyph to8192 potentially intersecting closed edge-box pairs, including adjacent
and touching pairs. These checks consume the existing visit budget. The
subsequent traversal rejects more than8192 nodes,65,536 output points,64 holes
or nesting depth64; those are post-allocation checks, not bounds on library
internal allocation. Input limits restrict admitted complexity but do not
establish a hard worst-case library work or memory bound. Clipper's synchronous
call has no token hook; cancellation is checked before and after it, including
a false result. Font parsing already in progress is also synchronous.
Unsupported/failed topology retains legacy width. The pinned managed package
has no native assets/transitive dependencies; its Boost1.0 notice and C#
maintenance-transition limitation are recorded in `THIRD-PARTY-NOTICES.md`.
Real selected bold faces,
strict missing-bold refusal, imported paint and name-only compatibility retain
their contracts. The runtime counter fixture is constructed from the existing
original test font; no new font payload is added.

The earlier Python counter-clearance probe only modeled circular dilation and
could borrow a nested hole. Its actual PNG observations remain diagnostic;
its80% deduction is not a miter-stroke proof. The production planner instead
uses the actual polygon-clearance test. Direct acute-triangle tests distinguish
a miter that reaches beyond half the line width from the limit10 bevel cutoff.

`FilledStrokeCoverageTests` reuses the pinned shared CJK bundle. At test-only
commit `80432418743d6b975c942e333fa531c32d23a80d`, unchanged e84 product code
fails the new default-image counter check: its two aperture row means are 22.3
and 18.4 against a required 100. The upper stroke must also remain present, so
blank output cannot pass. Semantic characters remain single. Direct/inherited
alpha, color and outline-only controls remain required.

The earlier coverage reference accidentally included the low-resolution native
stroke's expanded boundary in its opaque-interior mask. Actual same-process
BGRA records show all 572 disputed pixels outside the pure original-fill opaque
mask, with none missing inside it. The corrected check uses that fill mask,
plus bidirectional stroke coverage at 576 DPI within one boundary pixel, a
nonempty foreground requirement, and a 144-DPI integrated-opacity comparison.
The latter still rejects frozen pre-fix and e84 title outputs (about 24% and
27% excess respectively); it does not accept a changed oracle by assumption.
An affine-and-clip case checks the actual product path against original-font
geometry. Earlier failure records retain their original identities.

Full regression, API documentation, fresh exact-package consumers and actual
page review remain required. The new head and actual merged main must each bind
their own feed, outputs, public images and GUI receipts. Small-font fixture
results do not prove general Word fidelity, general bidi, R02, the remaining
R22 ledger or R20. The recurring native Skia availability risk is separate and
unresolved. No tag, Release or public NuGet publication is authorized here.


The SVG outlet applies the same single-emphasis principle only when the selected
embedded or explicitly bound face is a known non-bold static TrueType outline.
An opaque same-RGB filled stroke that already reaches the shared script-aware
emphasis width uses the original face with its declared stroke, without another
host font-weight synthesis. The comparison uses the existing writer grid: Han
spans use1%, other spans2.5%. The exporter does not increase the declared stroke.
Thin/hairline, different-color, partial-alpha, physical-bold and unknown name-only
cases retain their existing paths. Plain SVG text and original font bytes remain;
this does not establish native SVG selection or clipboard behavior. Actual
embedded-font visual review needs a renderer that proves it used those bytes.


Name-only missing-Noto titles previously received the same4.5%-em compatibility
stroke as viewer-local fonts. With matching Noto700 available, actual source
SVGs therefore combined a physical bold face and that compatibility outline;
this is separate from embedded-regular quantity/red-text synthesis. Fresh Native
DOCX now explicitly retains the name and requested Weight for face selection
without injecting this extra stroke for missing Noto CJK. It does not infer a
physical face from the name or promise identical host fallback. Fresh
SimSun/SimHei and the general typed writer's existing compatibility remain.

Imported SourceXml painting and CTM are authoritative for name-only as well as
embedded fonts. Packaging no longer injects or overwrites compatibility paint
on those envelopes, including omitted Stroke/Fill/LineWidth or an italic request
without CTM. This intentionally corrects the former imported name-only behavior;
it is not a heuristic that discards arbitrary explicit strokes. Tests cover
missing-Noto native/default production, two reader/writer cycles, fresh legacy
controls, clone-to-writer behavior, explicit color/alpha/outline/CTM paint and
omitted defaults. Actual host and matching-Noto output review remain separate.
