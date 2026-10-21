# Real-font shaping and prelayout query (R21)

This source branch adds a **new** query path. Existing `FlowDocument.Render` and
`OfdGraphics.DrawString` retain their approximate or caller-positioned behavior.
The query does not locate text in an imported OFD.

`Ofdrw.Net.Layout.Text` contains immutable face/request/result types and the
replaceable `ITextShaper` interface. The optional `Ofdrw.Net.Text.HarfBuzz`
project supplies `HarfBuzzTextShaper` and `HarfBuzzTextPrelayout.Create()`.
It pins HarfBuzzSharp and macOS, Win32 and Linux native asset packages to
`14.2.1.301`. Native loading is a runtime requirement; an absent or incompatible
library produces `TextLayoutError.NativeUnavailable`, never approximate metrics.

```csharp
var face = new TextFontFace("Noto Sans", File.ReadAllBytes(fontPath));
var request = new TextPrelayoutRequest("AV A\nTo", new[] { face },
    fontSizeMillimeters: 4, widthMillimeters: 50);
var result = HarfBuzzTextPrelayout.Create().Layout(request);
foreach (var line in result.Lines)
    Console.WriteLine($"{line.SourceStart}:{line.SourceLength} " +
        $"baseline={line.BaselineMillimeters} advance={line.AdvanceMillimeters}");
```

The caller supplies licensed font bytes. `TextFontFace` copies them, selects a
TTC face when requested, validates a bounded static monochrome TrueType profile,
and records both source and selected-face SHA-256. No system-font discovery or
synthetic weight/style change occurs. An explicit face list may choose a
whole-run fallback; `TEXT_FACE_FALLBACK` records the selected face. If no face
covers the whole run, the query fails. Mixed-face cluster fallback is outside
this first profile.

`TextShapeRequest` accepts one logical Unicode run, absolute .NET UTF-16 source
offset, face, direction, script, language and feature intervals. The result
holds actual GIDs, X/Y advances and offsets in font design units, cluster source
intervals, ink bounds and the engine package version. Clusters can be many-to-one
and their order can decrease in RTL. They are not character, grapheme or caret
indexes. The prelayout result reports total advance width/height, final line
positions and baselines, source intervals, ink rectangles, overflow, diagnostics,
and the exact shaped run used for each final line. Y is converted to downward
page millimeters once. Explicit LF/CRLF and bounded Latin/Han line opportunities
are supported; final lines are reshaped after break selection. Prelayout rejects
other script tags; direct single-run shaping remains caller-analyzed and has no
paragraph bidi or word-breaking contract. Format controls, variation selectors,
and variable/color/CFF fonts fail explicitly.

`TextPrelayoutOfdWriter.CreateSinglePage` consumes a completed result only for
regular ASCII Latin where each source character maps to one nominal glyph with
no offset or vertical advance. It preflights the whole page, uses the measured
Delta and baseline, and embeds the exact selected full face. A ligature,
combining sequence, unsupported style or overflowing page is rejected before
creating a package; page preflight checks both final advance and translated ink
extents, including negative side bearings. Arbitrary GID/cluster output,
glyph-aware subsetting, mixed-direction paragraphs and PDF/SVG semantic
consumption remain the separate
R22 stages; a successful R21 query does not imply those output profiles work.

The API enforces 32 MiB per font, 4096 UTF-16 units per run/request, 256 lines,
8192 maximum caller-selected shape calls, 8 million estimated cmap lookup steps
during whole-run face selection, and bounded native glyph expansion.
Cancellation is checked before and after native shaping and during layout;
the synchronous native call is not interruptible mid-instruction. Query failure
returns no partial result. File/stream writing retains the repository's separate
atomicity contract in [capability-contract-v12.md](capability-contract-v12.md).

The current checked evidence for this branch is recorded with its exact source
SHA and generated outputs in the R21 PR; this document defines the profile and
does not by itself establish cross-platform native loading or desktop-reader
acceptance.
