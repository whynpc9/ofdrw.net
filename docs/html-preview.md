# Static HTML previews

`Ofdrw.Net.Converter.Svg` also exports bounded multi-page HTML. No additional package or JavaScript is needed. The Reader selects the first OFD document; page indexes are positions in its sorted pages. The output is a static drawing projection, with no OFD actions, navigation or signature trust UI.

```csharp
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;

var exporter = new OfdToHtmlConverter(new OfdHtmlOptions { Title = "Report" });
await exporter.ConvertAsync(input, output, new[] { 2, 0, 2 }, cancellationToken);
```

The input is loaded once. Null selects all pages; an explicit empty selection fails. Selection order and duplicate occurrences are preserved. Each occurrence has independent clip IDs and embedded font families. The same embedded font is included again on each page, so large fonts and repeated selections can exhaust limits quickly. Successful font parsing is cached only within the current call. The existing single-page SVG API and `ofd-to-svg` command retain their behavior.

```text
ofdrw ofd-to-html input.ofd preview.html --pages 3,1,3 --title "Report"
ofdrw ofd-to-html --input input.ofd --output preview.html --max-pages 200 --max-output-bytes 67108864 --max-page-bytes 16777216
```

CLI pages are one-based; ascending ranges such as `1-3,2` are supported. Empty pieces, descending ranges, zero, overflow, more than 8192 selection characters and over-budget expansions fail before the input is opened. HTML is explicitly selected with `ofd-to-html`; the generic `convert` command does not route `.html` automatically. CLI output uses the existing same-directory atomic staging/commit mechanism; only completed commits report success. Exit codes are 0 for success, 1 for failure, and 130 for cancellation.

All limits are positive: 200 page occurrences, 64 MiB UTF-8 output, 16 MiB per section, 100,000 reserved generated nodes, 32 Mi UTF-16 generation capacity, 32 MiB original embedded resources and 1024 title code units by default. Node/character limits reserve conservative upper bounds before renderer allocations, including glyph expansion, XML projection, path normalization, Base64, CSS and intermediate copies; they may reject an output whose serialized size would fit. Resource and generation costs are charged again for duplicate pages. Reader input and model materialization limits remain independent. These limits are not process-wide memory, CPU or wall-clock limits.

The complete result is serialized without a BOM into a bounded private memory spool, then copied to the caller's stream. This intentionally avoids sensitive plaintext temporary files. Streams are left open and used at their current positions. Generation/selection/policy/limit failures leave public output unchanged. A failure or cancellation during final copying or flushing can leave a prefix in an arbitrary output stream; the caller must discard it. Different stream wrappers around the same underlying storage cannot be reliably detected and must not alias.

The exporter generates its own restricted SVG DOM and validates allowed elements, attributes, numeric values and page-local clip references. It does not sanitize arbitrary SVG. Titles and text are escaped in context. CSS contains only fixed templates and controlled font names/Base64. No scripts, event attributes, external URI loading, arbitrary CSS, embedded SVG images or foreign namespaces are emitted. A defensive CSP is included. Unknown visible drawings, unsupported font/glyph/annotation semantics and unknown drawing extensions are rejected. Unknown document metadata is not inserted into HTML.

Embedded media has a deliberately narrow profile without adding a decoder dependency: non-interlaced grayscale/RGB(A) PNG with bounded dimensions, known static chunks, complete chunk framing/CRC, valid zlib row lengths/filter bytes and Adler checksum; or 8-bit baseline Huffman JPEG with a single scan, checked segment framing, tables and entropy block traversal. Progressive/multiscan/arithmetic JPEG, indexed/interlaced/animated PNG, private or text PNG chunks, trailing payload, BMP/TIFF/GIF and other formats are rejected explicitly. Admission verifies this syntax; it does not certify browser decoders or color-management fidelity. Browsers can still have independent decoding/rendering limits. Image dimensions and the cumulative decoded pixel work (`MaxDecodedImagePixels`) are limited to 100 million pixels by default. Each image occurrence, including duplicate pages, is charged before PNG inflation or JPEG entropy traversal; over-budget work fails before expansion. The inner chunk/scanline/MCU loops check cancellation. Embedded fonts use existing OpenType/cmap admission and finite parse work.

SVG root font inheritance is reset so the preview shell does not replace the standalone SVG fallback font. Unembedded or unavailable font names still depend on the browser/host and are not a font identity guarantee. Inline SVG retains page mm dimensions and viewBox; narrow windows scale pages proportionally. Screen labels sit outside source coordinates. Print CSS gives each occurrence a separate page break; mixed physical paper sizes depend on the browser/print settings. Browser acceptance and PDF/Preview control results are recorded in `plans/029-followup-implementation.md`, and apply only to the inspected artifacts and pages.
