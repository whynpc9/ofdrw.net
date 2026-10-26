# R10 restricted image and region profile

This profile is for imported OFD content. The Reader retains the original
`SourceXml`, resource identifiers and encoded bytes. Exporters create temporary
PNG pixels for rendering; the package writer does not replace source media.

## Images

- A JBIG2 resource must have a literal `Format="JBIG2"` declaration matching
  the standalone file signature. The admitted subset is one known-size page,
  sequential segments `PageInformation`, one immediate arithmetic generic
  region covering the page, `EndOfPage`, and `EndOfFile`. No dictionaries,
  globals, references, refinement, MMR, unknown segment lengths or trailing
  payload are admitted. Input is capped at 1 MiB and decoded area at 1,048,576
  pixels. The decoder's dimensions and RGB count must match preflight.
- `ImageMask` is resolved in the same resource scope as the image. Self
  references, unresolved or mismatched declarations, and nonbinary formats
  fail. The admitted mask is one-bit grayscale PNG or the same restricted JBIG2
  profile. Dimensions must equal the source image. White mask pixels retain the
  source alpha; black pixels clear it. Object alpha is then applied by the
  exporter. GB/T 33190-2016 specifies a same-size binary mask resource but
  does not state that polarity in its ImageObject table; independent reader
  visual acceptance is required before treating this mapping as settled.
- PDF, SVG and HTML use the same normalized render projection. Unknown image
  children, attributes and resource references fail before public output.
  Malformed compressed data can still fail during decoding; the result is not
  a claim that every JBIG2 variant is supported.

The managed decoder is `JBig2Decoder.NETStandard` 1.5.2. Its original JPedal
BSD-3-Clause and C# port MIT notices are both retained in
`docs/licenses/JBIG2DECODER-NETSTANDARD-LICENSE.txt` and carried by the Reader
package. The test bitstream is generated from a self-authored pattern with an
independent `jbig2dec` oracle; see the R10 fixture README. The ITU-derived
third-party sample used during early diagnosis is not included.

## Region text

`OfdRegionTextExtractor` accepts a finite, positive-area rectangle in
millimeters relative to the selected page's PhysicalBox. It returns Unicode
scalars in background-template, body and foreground-template paint order,
without inventing spaces or line breaks. A scalar with ink is included only
when its transformed original-font glyph quadrilateral has positive-area
intersection with the rectangle; touching an edge alone does not select it.
For a no-ink scalar, the transformed baseline origin is used with inclusive
left/top and exclusive right/bottom. `OfdKeywordFragment` values expose the
source TextCode offsets, font identity, glyph ID, CTM/Delta-derived baseline
and exact font-box geometry. Templates and their referenced XML are checked
against source ownership. Annotation appearances are excluded and reported as
an `ExcludedAnnotationScope` diagnostic. Unsupported source, transform,
clipping, font geometry or budget exhaustion fails the entire query.

The query requires an explicit `IOfdGlyphGeometryProvider` for the exact
embedded font. It does not substitute host fonts or claim painted-pixel bounds.
