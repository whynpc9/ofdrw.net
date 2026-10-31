# D3: reject unverified PDF import reading order

The default PDF importer on merged main
`a05badc212a0b653a0a10afc9648e24c50c3f11a` reproduced a production defect:
the correctly shaped `fpdf2-two-word-cell.pdf` produces transparent OFD text
`abc םוקמ םולש def ` rather than the declared source `abc שלום מקום def`.
The raster preserves its visible phrase. These are separate observations.
PdfPig 0.1.15 raw results are retained; the export-reader decision does not
repair reverse import. This work is independent of PR #78: its importer file
is byte-identical to that main baseline.

The paired unshaped reference returns the declared source string from default
`GetWords()` while its appearance is wrong. Both PDFs have the same embedded
subset SHA-256 `69a75e445ae0446b956a0fd6683bd64ba54689ee032610fb26cf5ccd0bd4e959`
and fourteen `/ToUnicode` entries each. Their Unicode/GID sets agree, while
CID assignments and the actual `/ToUnicode` contents differ. A string comparison or unconditional RTL
reversal cannot establish a valid semantic import. Existing word rectangles
also do not preserve per-character origins or original font/GID identity:
the importer emits name-only OFD font resources and rounds geometry to 0.001 mm.

The repair rejects semantic extraction when parsed characters contain the
existing pinned Unicode 17 R/AL/bidi-control ranges. It runs before writing
semantic objects, and the text fallback performs the same check outside its
best-effort extraction catch. The converter throws `NotSupportedException`
with `PDF_TEXT_READING_ORDER_UNVERIFIED`; it does not silently publish guessed
logical words or turn this rejection into a successful placeholder. It does
not yet consume tagged logical structure or admit a new RTL import profile.

Callers may explicitly use `TextLayerMode=None` with
`RequireRasterization=true` for an image-only conversion. This retains the
input appearance, including a visually incorrect input, and claims no RTL
text semantics. Other complex PDF reading orders remain unverified.

The five focused regressions reuse the two immutable
[controls](../../tests/Ofdrw.Net.Converter.Pdf.Tests/Fixtures/D3/README.md).
On unchanged main, three rejection assertions fail and two raster-only cases
pass. With the repair, all five pass. Additional boundary regressions cover a
later RTL page leaving existing output untouched, selecting only the earlier
LTR page, and supplementary RTL/LTR scalars inside decoded mappings. The
package consumer repeats both immutable controls and emits raster-only
OFD/PDFs with input, output and loaded-DLL hashes. The fallback test invokes that method
directly; it does not simulate a real renderer outage or establish an
end-to-end fallback result. Existing LTR/page-selection regressions remain
required, along with full solution/API checks, exact package consumption,
actual generated-page review, latest-head CI and SES before acceptance.

The baseline read-only probe executed six public conversions (three inputs,
Invisible/None), retained raw PdfPig default/ActualText letters, words and
bounds, independently decoded font mappings, and rendered all nine input or
roundtrip pages. PNG inspection covered every full page and enlarged text
region. Invisible and None roundtrips have identical page pixels for each
input. Input and roundtrip pixels are not identical; the LTR roundtrip is
one pixel wider at 144 dpi. This is not a pixel-fidelity or Preview acceptance
claim. Candidate-specific functional and actual Preview evidence must bind
its own source, package manifest and output hashes.
