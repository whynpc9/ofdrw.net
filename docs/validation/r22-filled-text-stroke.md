# Filled text stroke follow-up

The final-main visual review at `a05badc212a0b653a0a10afc9648e24c50c3f11a`
found light interior tracks and a double-line appearance in CI-generated DOCX
titles. The source Title style requests bold blue text; its native OFD requests
both fill and stroke. This blocks overall visual acceptance even though the
text, original glyphs, source tests and package gates pass. Equality with an
older accepted raster does not waive the defect.

The deterministic source remains
`e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx` (SHA-256
`17bea68d57776b02e5c68c8915de11b42087809d5310f8fafcbb003db08d7bb6`).
For the ten inspected title characters, embedded PDF outlines match the OFD
font and the supplementary contour coordinates match the original font. This
evidence does not establish a baseline displacement or cover arbitrary fonts.
Separating the PDF text fill and boundary stroke isolates the visible seam;
filling the same original contour with its stroke restores a solid appearance.

PDF export therefore fills the supplementary contour when the resolved fill
and stroke have identical RGB values and both are fully opaque, including the
actual drawing brush. Repainting an opaque identical fill is idempotent. The
existing contour, transform, clipping and fill rule remain in use; no second
semantic text layer is emitted. Different colors, partial color alpha and
`Fill=false` keep their separate stroke semantics. Unsupported object opacity
remains rejected by preflight; no paint admission is widened.

`FilledStrokeCoverageTests` reuses the existing pinned font bundle. Its
independent path-only reference uses the original CJK contours and compares
visible interior pixels, not PDF operation names. At test-only commit
`a54aa2fe75f4b0ec592e4a40fbd58cdaccaae8a4`, both regular and bold requests
lose 81 expected interior pixels; five direct/inherited alpha, color and
outline controls pass. The controls compare actual interior RGBA values and
check that semantic text remains single. Existing bold alignment/counter tests
also remain required.

Full regression, API documentation, exact package consumption and fresh actual
page review are required for the repair. PR and merged-main results must bind
their own source, package manifest, products and GUI receipts. The `a05badc2`
functional matrix and attention receipt retain their identities. This repair
does not complete general bidi, R02, the remaining R22 ledger, or R20, and does
not authorize a tag, Release or public package publication.
