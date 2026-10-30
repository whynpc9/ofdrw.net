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
There is no new package, native dependency or replacement font.

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
