# Plan 027: DOCX native layout survives a first-line indent that does not fit the text column

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Layout/Internal/Flow/FlowFinalTextLayout.cs src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs tests/Ofdrw.Net.Converter.Docx.Tests/DocxNativeParagraphLayoutTests.cs`
> If any of these changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat
> it as a STOP condition.

## Status

- **Priority**: P2
- **Effort**: S
- **Risk**: LOW
- **Depends on**: none
- **Category**: bug
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

DOCX → native OFD has an opt-in paragraph layout engine
(`DocxConversionOptions.NativeParagraphLayoutEngine`). When a paragraph's
first-line indent (or hanging indent) is as large as its text column, the whole
conversion fails with a raw `TextLayoutException: Invalid frame, spacing or
indent geometry.` The same document converts without error when the engine is
not set.

Reproduced at `427433a` (Linux x64, Release) with two one-paragraph documents:

| Paragraph properties | Engine off | Engine on |
|---|---|---|
| `<w:ind w:left="10000" w:hanging="720"/>` | converts | `Ofdrw.Net.Layout.Text.TextLayoutException: Invalid frame, spacing or indent geometry.` |
| `<w:ind w:left="4000" w:right="4000" w:firstLine="2000"/>` | converts | same exception |

Two things are wrong. The indent is never brought into the range the engine
accepts, and the exception type is not the `InvalidDataException("DOCX native
paragraph layout failed: ...")` that every other engine failure is turned into.
Such documents are unusual (large indents in a narrow column), but one such
paragraph loses the whole document.

After this plan those paragraphs are laid out with the indent limited to what
fits, the conversion reports a diagnostic once, and no layout option error can
escape unwrapped.

## Current state

### The engine's option check

`src/Ofdrw.Net.Layout/Text/Paragraphs/ParagraphLayoutRequest.cs:365-370`:

```csharp
        if (!Finite(widthMillimeters) || widthMillimeters <= 0 || widthMillimeters > 1_000_000 ||
            pageHeightMillimeters is { } height && (!Finite(height) || height <= 0 || height > 1_000_000) ||
            !Finite(paragraphSpacingMillimeters) || paragraphSpacingMillimeters < 0 || paragraphSpacingMillimeters > 1000 ||
            !Finite(firstLineIndentMillimeters) || firstLineIndentMillimeters <= -widthMillimeters ||
            firstLineIndentMillimeters >= widthMillimeters)
            throw new TextLayoutException(TextLayoutError.InvalidRequest, "Invalid frame, spacing or indent geometry.");
```

So the indent must satisfy `-width < indent < width`.

### Where the options are built — outside the `try`

`src/Ofdrw.Net.Layout/Internal/Flow/FlowFinalTextLayout.cs` (`internal sealed
class FlowFinalTextLayout`, namespace `Ofdrw.Net.Layout.Internal.Flow`), lines
166–181:

```csharp
        // A style change inside an extended grapheme cluster: that cluster takes its leading style, as in
        // the compatibility layout; the following text keeps its own style.
        var layout = Run(text, ClusterAlignedSpans(text, spans),
            Options(format, width, first ? format.FirstLineIndentMillimeters : 0, text.Length), token);
        foreach (var line in layout.Lines)
            result.Add(new FlowLine(new List<FlowGlyph>(), line.HeightMillimeters, 0, FlowAlignment.Left) { Final = line });
    }

    private ParagraphLayoutResult Run(string text, IReadOnlyList<TextStyleSpan> spans, ParagraphLayoutOptions options,
        CancellationToken token)
    {
        try { return _engine.Layout(new ParagraphLayoutRequest(text, spans, options), token); }
        catch (TextLayoutException exception)
        {
            throw new InvalidOperationException("Final text layout failed (" + exception.Error + "): " + exception.Message, exception);
        }
    }
```

`Options(...)` (lines 133–151 of the same file) calls
`new ParagraphLayoutOptions(width, null, ..., 0, indent, TextOverflowPolicy.BreakGraphemes, ...)`.
It runs while the arguments of `Run` are evaluated, so its exception is not
caught by `Run`'s `try`.

### Where DOCX computes the indent

`src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs:226-229`:

```csharp
        var leftIndent = PointsToMillimeters(paragraph.Format.LeftIndentPoints ?? 0);
        var rightIndent = PointsToMillimeters(paragraph.Format.RightIndentPoints ?? 0);
        var firstIndent = PointsToMillimeters(paragraph.Format.FirstLineIndentPoints ?? 0);
        var width = Math.Max(availableWidth - leftIndent - rightIndent, 5d);
```

`width` has a 5 mm floor; `firstIndent` (negative for a hanging indent) has no
limit. Both go to `LayoutParagraph(paragraph, width, firstIndent, ...)` (line
556), which stores the indent (lines 611–616):

```csharp
        var format = new FlowParagraphFormat
        {
            Alignment = alignment,
            FirstLineIndentMillimeters = firstIndent,
            MinimumLineHeightMillimeters = PointsToMillimeters(paragraph.Format.LineSpacingPoints ?? 0)
        };
```

and ends (lines 637–645):

```csharp
        try
        {
            return FlowParagraphLayout.Layout(inlines, format, width, this, fallback, _cancellationToken, allowOversizeGlyph: true,
                final: _final);
        }
        catch (InvalidOperationException exception) when (_final is not null && exception.InnerException is TextLayoutException)
        {
            throw new InvalidDataException("DOCX native paragraph layout failed: " + exception.InnerException.Message, exception);
        }
```

`LayoutParagraph` has four call sites (lines 233, 263, 376, 999); line 999 is
the table-cell path and passes the cell's own width.

The existing once-only diagnostic pattern in the same method (field at line 47,
use at lines 605–610):

```csharp
        if (_final is not null && !_imageFallbackReported && inlines.Any(inline => inline.Image is not null))
        {
            _imageFallbackReported = true;
            _diagnostics.Add(new DocxConversionDiagnostic("DOCX_NATIVE_TEXT_ENGINE_FALLBACK",
                "Paragraphs with inline images keep the compatibility text layout; other paragraphs use the paragraph layout engine."));
        }
```

### Who else reaches `FlowFinalTextLayout`

The public Flow API validates indents first
(`src/Ofdrw.Net.Layout/FlowDocument.cs:381-389`: negative values and
`FirstLineIndentMillimeters >= width` throw), so only DOCX can deliver an
out-of-range indent. In-range indents must keep producing exactly today's
output.

### Tests and docs

`tests/Ofdrw.Net.Converter.Docx.Tests/DocxNativeParagraphLayoutTests.cs` has the
helpers this plan reuses:

```csharp
    private static DocxConversionOptions Options(PinnedCjkFontDirectory fonts, bool engine)
    private static async Task<(OfdDocumentPackage Package, DocxToOfdConversionResult Result)> Convert(Stream source, DocxConversionOptions options)
    private static string Text(OfdDocumentPackage package)   // all text, spaces removed
```

and builds documents from body XML, e.g. (line 62):

```csharp
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var (package, _) = await Convert(source, Options(fonts, engine: true));
```

`docs/content/guides/final-text-layout.md:59` lists the engine's DOCX
diagnostics (`DOCX_NATIVE_TEXT_ENGINE_FALLBACK`,
`DOCX_FONT_TEXT_ENGINE_UNSUPPORTED`). The guide says the engine is available
since 0.3.0 (line 9). This repository marks documentation for behaviour that is
on `main` but not in a published version with `（main，尚未发布）`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| DOCX tests | `dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Layout tests | `dotnet test tests/Ofdrw.Net.Layout.Tests/Ofdrw.Net.Layout.Tests.csproj -c Release --no-build --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| OFD → PDF | `dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll ofd-to-pdf --input <in.ofd> --output <out.pdf>` | exit 0 |
| Docs check | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | all exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Layout/Internal/Flow/FlowFinalTextLayout.cs`
- `src/Ofdrw.Net.Converter.Docx/Internal/BuiltIn/BuiltInOfdRenderer.cs`
- `tests/Ofdrw.Net.Converter.Docx.Tests/DocxNativeParagraphLayoutTests.cs`
- `docs/content/guides/final-text-layout.md` (one sentence)
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- `ParagraphLayoutOptions` validation — the engine's contract stays as it is.
- `FlowParagraphLayout.cs` (the compatibility layout) and the public Flow API
  validation in `FlowDocument.cs` / `FlowTableRenderer.cs`.
- The 5 mm width floor in `BuiltInOfdRenderer.cs:229`.

## Git workflow

- Branch: `codex/docx-native-indent-geometry`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Write the failing tests

Add to `DocxNativeParagraphLayoutTests`:

```csharp
    [Theory]
    [InlineData("<w:ind w:left=\"10000\" w:hanging=\"720\"/>")]
    [InlineData("<w:ind w:left=\"4000\" w:right=\"4000\" w:firstLine=\"2000\"/>")]
    public async Task IndentThatDoesNotFitTheColumnIsLimitedNotFatal(string indent)
    {
        var body = "<w:p><w:pPr>" + indent + "</w:pPr><w:r><w:t>Indented 缩进 text</w:t></w:r></w:p>" +
                   "<w:p><w:r><w:t>Next 下一段</w:t></w:r></w:p>";
        using var fonts = new PinnedCjkFontDirectory();
        using var source = DocxReleaseFixtures.CreateMinimalDocx(body);
        var (package, result) = await Convert(source, Options(fonts, engine: true));
        Assert.Equal("Indented缩进textNext下一段", Text(package));
        Assert.Single(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_NATIVE_INDENT_LIMITED");
        var page = package.Pages.First();
        Assert.All(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>(),
            text => Assert.InRange(text.XMillimeters, 0, page.WidthMillimeters));
        if (Environment.GetEnvironmentVariable("OFDRW_PARAGRAPH_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            using var again = DocxReleaseFixtures.CreateMinimalDocx(body);
            using var output = new MemoryStream();
            await new DocxToOfdConverter(Options(fonts, engine: true)).ConvertWithResultAsync(again, output);
            await File.WriteAllBytesAsync(Path.Combine(directory,
                "indent-limited-" + (indent.Contains("hanging") ? "hanging" : "first-line") + ".ofd"), output.ToArray());
        }
    }
```

`WidthMillimeters` is declared on the page model at
`src/Ofdrw.Net.Core/Models/OfdPage.cs:25`.

**Verify**: run the DOCX tests with
`--filter "FullyQualifiedName~IndentThatDoesNotFitTheColumnIsLimitedNotFatal"` →
both cases FAIL with `TextLayoutException: Invalid frame, spacing or indent geometry.`

### Step 2: Limit the indent where the engine is called

In `FlowFinalTextLayout.cs` add:

```csharp
    // The engine accepts -width < indent < width. DOCX indents arrive unvalidated; an indent outside
    // that range is pulled in so the first line keeps 5 mm (or half of a narrower frame). In-range
    // indents are returned unchanged.
    internal static double LimitFirstLineIndent(double indent, double width)
    {
        if (double.IsNaN(indent) || double.IsInfinity(indent)) return 0;
        if (indent > -width && indent < width) return indent;
        return indent > 0 ? width - Math.Min(5, width / 2) : -width * 0.999;
    }
```

In `Lines`, build the arguments inside a `try` so no option error escapes
unwrapped, and use the limited indent:

```csharp
        ParagraphLayoutOptions options;
        try
        {
            options = Options(format, width,
                first ? LimitFirstLineIndent(format.FirstLineIndentMillimeters, width) : 0, text.Length);
        }
        catch (TextLayoutException exception)
        {
            throw new InvalidOperationException("Final text layout failed (" + exception.Error + "): " + exception.Message, exception);
        }
        var layout = Run(text, ClusterAlignedSpans(text, spans), options, token);
```

Keep the existing comment above the `Run` call.

**Verify**: build → `0 Error(s)`; the Step 1 tests now fail only on the missing
`DOCX_NATIVE_INDENT_LIMITED` diagnostic.

### Step 3: Report the adjustment once per conversion

In `BuiltInOfdRenderer.cs` add a field next to `_imageFallbackReported`:
`private bool _indentLimitReported;`. In `LayoutParagraph`, directly after the
`DOCX_NATIVE_TEXT_ENGINE_FALLBACK` block:

```csharp
        if (_final is not null && !_indentLimitReported &&
            FlowFinalTextLayout.LimitFirstLineIndent(firstIndent, width) != firstIndent)
        {
            _indentLimitReported = true;
            _diagnostics.Add(new DocxConversionDiagnostic("DOCX_NATIVE_INDENT_LIMITED",
                "A first-line or hanging indent did not fit its text column and was limited to the column."));
        }
```

The renderer runs several layout passes (`FinalizePagination`), so the flag must
stay set across passes: do not reset it anywhere.

**Verify**: Step 1 tests → `Passed: 2`.

### Step 4: Document the diagnostic

In `docs/content/guides/final-text-layout.md`, in the list item at line 59 that
names the DOCX diagnostics, add one sentence in Chinese, for example:
`首行缩进或悬挂缩进超出文本栏宽度时，缩进被限制在栏内，并报告一次 DOCX_NATIVE_INDENT_LIMITED（main，尚未发布）。`
(put the diagnostic code in backticks like its neighbours). Do not add a new page.

**Verify**: the docs check commands all exit 0.

### Step 5: Regression and visual check

1. Run the full DOCX and Layout test projects → `Failed: 0`. These include
   `GeneratedLayoutSampleKeepsItsTextAndUsesClusterMappedFinalLayout` and
   `TabsStrikeHighlightAndHangingIndentReachTheEngine`, which prove in-range
   indents are unchanged.
2. Render the two new samples:

```sh
OUT="$(mktemp -d)"
OFDRW_PARAGRAPH_ARTIFACTS="$OUT" dotnet test tests/Ofdrw.Net.Converter.Docx.Tests/Ofdrw.Net.Converter.Docx.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~IndentThatDoesNotFitTheColumnIsLimitedNotFatal" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
for f in "$OUT"/indent-limited-*.ofd; do
  dotnet src/Ofdrw.Net.Cli/bin/Release/net10.0/Ofdrw.Net.Cli.dll ofd-to-pdf --input "$f" --output "${f%.ofd}.pdf"
  pdftoppm -png -r 144 "${f%.ofd}.pdf" "${f%.ofd}"
done
```

Open each PNG: both paragraphs are fully inside the page, no glyphs overlap,
Chinese and Latin text are legible, the second paragraph starts at the left
margin. The viewing chain is DOCX → native OFD → PDF → PNG. Record the pages
viewed; macOS Preview acceptance is not possible on Linux: report "not
performed".

**Verify**: two `.png` files exist and the review found no defect.

## Test plan

- `IndentThatDoesNotFitTheColumnIsLimitedNotFatal` (2 cases) in
  `DocxNativeParagraphLayoutTests.cs`, modelled on
  `TabsStrikeHighlightAndHangingIndentReachTheEngine` in the same file.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] New theory → 2 passed; full DOCX and Layout test projects → `Failed: 0`
- [ ] `grep -n "DOCX_NATIVE_INDENT_LIMITED" src docs/content -r` → one hit in `BuiltInOfdRenderer.cs`, one in `final-text-layout.md`
- [ ] Docs check commands exit 0
- [ ] `git status --short` lists only in-scope files; storage gate exits 0
- [ ] `plans/plan-index.md` status row updated; report lists functional tests, PNG review and "macOS Preview: not performed" separately

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- Step 1 tests do not fail with the stated exception before the change (the
  defect may already be fixed).
- After Step 2 a case fails with a different engine error (for example
  `A single extended grapheme cluster exceeds the line width.`). Report it; do
  not change the 5 mm floor or the overflow policy.
- Any existing DOCX or Layout test changes its result.

## Maintenance notes

- `LimitFirstLineIndent` is deliberately a no-op for indents the engine already
  accepts; do not "tidy" it into an unconditional clamp, which would move
  existing output.
- If the renderer's 5 mm width floor changes, revisit the `Math.Min(5, width / 2)`
  term so the two stay consistent.
- Reviewers: check that the diagnostic appears once even though pagination
  re-runs the layout.
