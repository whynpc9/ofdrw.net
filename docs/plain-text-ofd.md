# Plain text to native OFD

`Ofdrw.Net.Layout` provides a thin adapter over `FlowDocument`. Both APIs return a new in-memory package; use `OfdPackageWriter` to save it.

```csharp
var options = new PlainTextOfdOptions { MaxInputBytes = 100_000 };
options.Layout.DefaultFontSizeMillimeters = 4;
options.Layout.MaxPageCount = 100;
var converter = new PlainTextOfdConverter(options);
var package = converter.Render("通知 Alpha iii WWW\r\n\r\n尾行");
// Or read strict UTF-8 from the current position of a borrowed stream:
var fromUtf8 = await converter.RenderUtf8Async(input, cancellationToken);
await new OfdPackageWriter().WriteAsync(fromUtf8, output, cancellationToken);
```

The constructor validates and snapshots all settings. Calls use independent Flow state. Layout defaults are A4, 25.4 mm margins, SimSun, black and 10.5 pt expressed in millimeters. Budgets default to 4,000,000 input bytes (UTF-8 API only), 1,000,000 UTF-16 code units, 10,000 pages and 1,000,000 text elements. Exceeding a budget or admitting text that cannot fit throws `InvalidDataException`. Cancellation propagates as `OperationCanceledException`; I/O failures propagate. Streams remain open, may be non-seekable, and may have advanced on failure.

UTF-8 decoding is strict and chunked. Only one initial byte BOM is removed; subsequent U+FEFF and a string's initial U+FEFF remain content. Invalid UTF-8, unpaired UTF-16 surrogates and XML-invalid scalars are rejected. CRLF, CR and LF are logical line breaks; interior empty lines remain, and terminal breaks do not create empty tail pages. Form feed is a page control with the same deferred tail behavior. Empty text/BOM-only input produces one blank page. Tabs retain their character and use Flow's four-space advance, without tab stops. U+2028/U+2029 remain ordinary text.

Native text objects preserve the admitted string after CRLF/CR normalization and removal of LF/form-feed layout controls, including spaces, tabs, NBSP and combining characters. Reading extraction reconstructs lines from coordinates; automatic wrapping and blank lines prevent byte-for-byte or original-newline recovery. No hidden source-text layer is added.

Font-family names request Flow's existing portable metrics. This adapter does not measure arbitrary installed font files, embed fonts, validate glyph coverage or guarantee identical appearance on different devices. OFD-to-PDF export may resolve a physical fallback font; inspect actual output for the deployment's text and fonts. Markdown, rich styling, tables, images, headers/footers and syntax highlighting are outside this entrypoint.

## CLI

```text
ofdrw text-to-ofd notice.txt notice.ofd --font-size-mm 4
ofdrw text-to-ofd --input notice.txt --output notice.ofd \
  --page-width 210 --page-height 297 \
  --margin-left 15 --margin-right 15 --margin-top 15 --margin-bottom 15 \
  --max-input-bytes 100000 --max-characters 50000 --max-pages 100 --max-text-elements 50000
```

Use two positional paths or both named paths, without mixing. `--font-family` changes the name request. Page width/height must be paired; dimensions and font size must be finite positive numbers, margins finite nonnegative numbers. Numbers use invariant decimal notation. Duplicate/unknown options are rejected before opening the input. The output must end in `.ofd`; the input extension is unrestricted but content must be UTF-8. `convert` inference is unchanged. stdin/stdout, `--pages`, DOCX options and `--font-directory` are unsupported.

The command validates and renders before staging, writes beside the destination, and atomically replaces it only after completion. Failures or cancellation before commit retain an existing destination and remove staging. Cancellation after replacement cannot undo the commit. Existing conversion path identity protection rejects the same path and symbolic aliases. Exit codes are 0 (success), 1 (error), and 130 (canceled).

Implementation evidence and inspected sample scope: [Plan 028 follow-up](../plans/028-followup-implementation.md).
