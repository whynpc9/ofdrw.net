# Plan 034: OFD → image export can keep transparency or use a chosen background (SDK)

> **Executor instructions**: Follow this plan step by step. Run every
> verification command and confirm the expected result before moving to the
> next step. If anything in the "STOP conditions" section occurs, stop and
> report — do not improvise. When done, update the status row for this plan
> in `plans/plan-index.md` — unless a reviewer dispatched you and told you they
> maintain the index.
>
> **Drift check (run first)**: `git diff --stat 427433a..HEAD -- src/Ofdrw.Net.Converter.Pdf/ImageIO tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs docs/content/guides/images.md`
> If any of these changed since this plan was written, compare the "Current
> state" excerpts against the live code before proceeding; on a mismatch, treat
> it as a STOP condition.

## Status

- **Priority**: P3 (feature; the maintainer listed it as the next candidate after the 2026-11 upstream review)
- **Effort**: M
- **Risk**: LOW (default output must not change)
- **Depends on**: none
- **Category**: direction
- **Planned at**: commit `427433a`, 2026-11-13

## Why this matters

`OfdToImageConverter` always flattens the page onto white and writes an RGB
PNG (or JPEG). Callers who place page images on a coloured surface, or overlay
stamps and pages, need the page's own transparency or a background of their
choice. The maintainer recorded this as the preferred next feature in
`plans/019-upstream-2026-11-validation.md` ("后续功能候选：图片背景与透明 PNG")
with one condition: first prove that the existing PDF/PDFium chain delivers
correct RGBA — the feature must not be faked by turning white pixels
transparent.

That proof was done during planning (Linux x64, Release build at `427433a`): a
40 × 40 mm page with filled rectangles was written as OFD, converted with
`OfdToPdfConverter`, and rasterised with the same Docnet/PDFium call the
exporter uses. Raw pixels (B, G, R, A):

| Region | Raw pixel |
|---|---|
| Nothing painted | 0, 0, 0, **0** |
| Opaque red fill (`OfdColor(255, 0, 0)`) | 0, 0, 255, **255** |
| Blue fill with alpha 128 over nothing | 255, 0, 0, **127** |
| Same blue over the opaque red | 127, 0, 128, **255** |

So unpainted areas are fully transparent, a half-transparent fill keeps its
alpha with **straight (not premultiplied)** colour, and overlaps are blended
correctly. A text-only page gave 2,001,632 transparent pixels, 1,110 partially
transparent (glyph edges) and 28 opaque. The alpha is real; today's exporter
discards it in one line. **Not** checked during planning: images with their own
alpha channel, signature stamps, and Windows/macOS rasters — the tests in this
plan cover the first, and STOP conditions cover surprises.

Scope follows the maintainer's note: SDK only, default (white) unchanged,
explicit RGBA for PNG, JPEG always on an opaque background, unsupported
combinations rejected, no CLI option, no rotation/BMP/batch naming.

## Current state

### Options

`src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageOptions.cs`:

```csharp
/// <summary>Resolution, encoding and budgets for one OFD page export.</summary>
public sealed class OfdToImageOptions
{
    /// <summary>Pixels per millimeter; default 144 DPI. Must be finite and positive.</summary>
    public double PixelsPerMillimeter { get; set; } = 144d / 25.4d;
    /// <summary>Output encoding, independent of the destination filename.</summary>
    public OfdImageFormat Format { get; set; } = OfdImageFormat.Png;
    /// <summary>JPEG quality from 1 to 100, default 90.</summary>
    public int JpegQuality { get; set; } = 90;
    ...
```

### The converter

`src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageConverter.cs`. The constructor
validates and **copies** the options into a private instance (lines 26–52):

```csharp
        if (!Enum.IsDefined(typeof(OfdImageFormat), options.Format) || options.JpegQuality < 1 || options.JpegQuality > 100 ||
            options.MaxSignatureAppearanceCount <= 0 || options.MaxPixels <= 0 || options.MaxRasterWorkingBytes <= 0 || options.MaxIntermediatePdfBytes <= 0 || options.MaxOutputBytes <= 0)
            throw new ArgumentException("Encoding, quality and positive image budgets are required.", nameof(options));
        ...
        _options = new OfdToImageOptions
        {
            PixelsPerMillimeter = options.PixelsPerMillimeter, Format = options.Format, JpegQuality = options.JpegQuality,
            MaxSignatureAppearanceCount = options.MaxSignatureAppearanceCount,
            ...
```

The end of `ConvertAsync` (lines 93–106):

```csharp
        var raw = raster.GetImage();
        cancellationToken.ThrowIfCancellationRequested();
        using var source = Image.LoadPixelData<Bgra32>(raw, actualWidth, actualHeight);
        using var canvas = new Image<Rgb24>(actualWidth, actualHeight, Color.White);
        canvas.Mutate(context => context.DrawImage(source, 1f));
        // PDFium rounds physical sizes to integer pixels. Normalize the at-most-one-pixel difference to the documented ceil grid.
        if (actualWidth != width || actualHeight != height) canvas.Mutate(context => context.Resize(width, height));
        using var encoded = new ImageIoStagingStream(_options.MaxOutputBytes);
        if (_options.Format == OfdImageFormat.Png)
            await canvas.SaveAsync(encoded, new PngEncoder { ColorType = PngColorType.Rgb }, cancellationToken).ConfigureAwait(false);
        else
            await canvas.SaveAsync(encoded, new JpegEncoder { Quality = _options.JpegQuality }, cancellationToken).ConfigureAwait(false);
        await encoded.PublishAsync(imageOutput, cancellationToken).ConfigureAwait(false);
```

The class summary (line 19) says "composited onto white".

### Budget

`src/Ofdrw.Net.Converter.Pdf/ImageIO/ImageIoBudget.cs:16-23` already reserves 16
bytes of working memory per pixel
(`(long)width * height > workingBytes / 16`). The raw raster (4), the loaded
source (4) and an RGBA canvas (4) fit in that; no budget change is needed.

### Tests and docs

`tests/Ofdrw.Net.Converter.Pdf.Tests/ImageIoTests.cs` holds the existing image
tests (ImageSharp is already referenced there:
`using SixLabors.ImageSharp; using SixLabors.ImageSharp.PixelFormats;`). A
typical call:

```csharp
        await new OfdToImageConverter(new OfdToImageOptions { Format = format, PixelsPerMillimeter = ppm }).ConvertAsync(input, output, page);
```

The same test project keeps review artifacts when an environment variable is
set (`OFDRW_TEST_ARTIFACTS`, see
`tests/Ofdrw.Net.Converter.Pdf.Tests/TypedGraphicRenderingTests.cs:75`).

A page with filled rectangles is built from the public model like this (the
fixture used for the measurements above):

```csharp
    private static OfdPathElement Rect(double x, double y, double w, double h, OfdColor fill) => new()
    {
        XMillimeters = x, YMillimeters = y, WidthMillimeters = w, HeightMillimeters = h,
        AbbreviatedData = FormattableString.Invariant($"M 0 0 L {w} 0 L {w} {h} L 0 {h} C"),
        Stroke = false, Fill = true, FillColor = fill,
    };

    private static async Task<MemoryStream> Fixture(double size = 40)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { Index = 0, WidthMillimeters = size, HeightMillimeters = size };
        page.Elements.Add(Rect(0, 0, 20, 20, new OfdColor(255, 0, 0)));         // opaque red
        page.Elements.Add(Rect(20, 0, 20, 20, new OfdColor(0, 0, 255, 128)));   // half blue over nothing
        page.Elements.Add(Rect(0, 25, 20, 10, new OfdColor(255, 0, 0)));        // opaque red ...
        page.Elements.Add(Rect(10, 25, 20, 10, new OfdColor(0, 0, 255, 128)));  // ... under half blue
        package.Pages.Add(page);
        var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        return stream;
    }
```

At 10 pixels per millimetre the sample points are (x, y in pixels): opaque red
(100, 100); half blue over nothing (300, 100); blank (350, 380); half blue over
red (150, 300).

User docs: `docs/content/guides/images.md:29` says
"每次调用导出所选的一个页面，背景合成到白色。" Unreleased behaviour is marked
`（main，尚未发布）`.

## Commands you will need

```sh
export DOTNET_CLI_HOME="$(mktemp -d)" DOTNET_SKIP_FIRST_TIME_EXPERIENCE=1 DOTNET_CLI_TELEMETRY_OPTOUT=1
```

| Purpose | Command | Expected on success |
|---|---|---|
| Fetch fixtures | `git lfs pull --include="e2e/**,tests/**" --exclude=""` | exit 0 |
| Build | `dotnet build Ofdrw.Net.sln -c Release --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `0 Error(s)` |
| Image tests | `dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~Image" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| All PDF converter tests | same without `--filter` | `Failed: 0` |
| CLI tests | `dotnet test tests/Ofdrw.Net.Cli.Tests/Ofdrw.Net.Cli.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~ImageCommandTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false` | `Failed: 0` |
| Docs check | `cd docs/site && npm ci && npm run build && npm run typecheck && npm run check` | all exit 0 |
| Storage gate | `git add -A && python3 scripts/check-repository-storage.py` | exit 0 |

## Scope

**In scope** (the only files you should modify):

- `src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdImageBackground.cs` (create)
- `src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageOptions.cs`
- `src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdToImageConverter.cs`
- `tests/Ofdrw.Net.Converter.Pdf.Tests/ImageBackgroundTests.cs` (create)
- `docs/content/guides/images.md`
- `plans/plan-index.md` (status row only)

**Out of scope** (do NOT touch):

- The CLI (`ofd-to-image` keeps its white background and its help text).
- `ImageIoBudget.cs`, the OFD → PDF renderer, PDFium/Docnet settings.
- Image import (`ImageToOfdConverter`), rotation, BMP, multi-page naming.
- `docs/validation/` records and `plans/019-upstream-2026-11-validation.md`.

## Git workflow

- Branch: `codex/image-export-background`.
- Commit messages: one imperative sentence, capitalized, no prefix. Example
  from `git log`: `Prevent color bleed at nested OFD riding seal seams`.
- Do NOT push or open a PR unless the operator instructed it.

## Steps

### Step 1: Pin today's default output (before any source change)

Create `tests/Ofdrw.Net.Converter.Pdf.Tests/ImageBackgroundTests.cs` (namespace
`Ofdrw.Net.Converter.Pdf.Tests`) with the `Rect`/`Fixture` helpers above and a
helper that exports at 10 px/mm and decodes the result:

```csharp
    private static async Task<(Image<Rgba32> Image, byte[] Bytes)> Export(OfdToImageOptions options, double size = 40)
    {
        options.PixelsPerMillimeter = 10;
        using var input = await Fixture(size);
        using var output = new MemoryStream();
        await new OfdToImageConverter(options).ConvertAsync(input, output);
        return (Image.Load<Rgba32>(output.ToArray()), output.ToArray());
    }
```

Add `DefaultStaysOpaqueWhiteRgb`: export with `new OfdToImageOptions()`. Assert
the image is 400 × 400; blank pixel is exactly (255, 255, 255, 255); opaque red
is (255, 0, 0, 255); half blue over nothing is (R 126–130, G 126–130, B 255,
A 255); and the PNG has no alpha channel
(`Image.Identify(bytes).Metadata.GetPngMetadata().ColorType == PngColorType.Rgb`).

**Verify**: Image tests command with
`--filter "FullyQualifiedName~ImageBackgroundTests"` → `Passed: 1`. Commit. This
test must keep passing unchanged through every later step.

### Step 2: The background option

Create `src/Ofdrw.Net.Converter.Pdf/ImageIO/OfdImageBackground.cs` in namespace
`Ofdrw.Net.Converter.Pdf`:

```csharp
/// <summary>What an exported page image shows where the page paints nothing. Immutable.</summary>
public sealed class OfdImageBackground
{
    private OfdImageBackground(byte red, byte green, byte blue, bool transparent)
    { Red = red; Green = green; Blue = blue; IsTransparent = transparent; }

    /// <summary>Opaque white; the default.</summary>
    public static OfdImageBackground White { get; } = new(255, 255, 255, false);
    /// <summary>Keeps the page's own transparency. PNG only; the output has an alpha channel.</summary>
    public static OfdImageBackground Transparent { get; } = new(0, 0, 0, true);
    /// <summary>An opaque background colour.</summary>
    public static OfdImageBackground Opaque(byte red, byte green, byte blue) => new(red, green, blue, false);

    /// <summary>True for <see cref="Transparent"/>.</summary>
    public bool IsTransparent { get; }
    /// <summary>Red component of an opaque background.</summary>
    public byte Red { get; }
    /// <summary>Green component of an opaque background.</summary>
    public byte Green { get; }
    /// <summary>Blue component of an opaque background.</summary>
    public byte Blue { get; }
}
```

In `OfdToImageOptions` add, after `JpegQuality`:

```csharp
    /// <summary>Page background; default opaque white. <see cref="OfdImageBackground.Transparent"/> requires PNG.</summary>
    public OfdImageBackground Background { get; set; } = OfdImageBackground.White;
```

In the converter constructor, after the existing validation block, add:

```csharp
        if (options.Background is null) throw new ArgumentException("An image background is required.", nameof(options));
        if (options.Background.IsTransparent && options.Format != OfdImageFormat.Png)
            throw new ArgumentException("JPEG cannot store transparency; use PNG or an opaque background.", nameof(options));
```

and copy the value in the `_options = new OfdToImageOptions { ... }` initializer
(`Background = options.Background,`).

**Verify**: build → `0 Error(s)`; Step 1 test still passes.

### Step 3: Tests for the new behaviour (failing first)

Add to `ImageBackgroundTests` (tolerances are ±2 per channel unless stated):

- `TransparentPngKeepsThePagesAlpha`: `Background = OfdImageBackground.Transparent`.
  Blank → alpha 0. Opaque red → (255, 0, 0, 255). Half blue over nothing →
  B 255, R 0, G 0, alpha 125–130. Half blue over red → (128, 0, 127, 255).
  PNG `ColorType == PngColorType.RgbWithAlpha`.
- `OpaqueColourIsUsedAsBackground`: `Opaque(0, 255, 0)`. Blank → (0, 255, 0,
  255). Half blue over nothing → (0, 128, 127, 255). Opaque red unchanged.
  `ColorType == PngColorType.Rgb`.
- `JpegUsesTheOpaqueBackground`: `Format = Jpeg`, `Opaque(0, 0, 0)`. Blank pixel
  has every colour channel below 16.
- `JpegWithTransparentBackgroundIsRejected`: constructing the converter throws
  `ArgumentException`.
- `NullBackgroundIsRejected`: `Background = null!` → `ArgumentException`.
- `OptionsAreCopiedAtConstruction`: create options with `Transparent`, construct
  the converter, then set `options.Background = OfdImageBackground.White`, then
  convert with the already-constructed converter → blank pixel alpha is 0.
- `ResizedTransparentPageKeepsItsEdges`: `Fixture(size: 40.03)` and
  `Transparent` → image is 401 × 401 (the documented ceil grid); pixel (395,
  395) has alpha 0; pixel (100, 100) is red with alpha 255 (±2 on colour).
- `BudgetAppliesToTransparentExport`: `Transparent` with `MaxPixels = 100` →
  `InvalidDataException`, and the output stream is still empty.
- If `OFDRW_TEST_ARTIFACTS` is set, write the default, transparent and green
  PNGs into that directory (names `background-default.png`,
  `background-transparent.png`, `background-green.png`).

**Verify**: build; the tests that need the new pipeline
(`TransparentPngKeepsThePagesAlpha`, `OpaqueColourIsUsedAsBackground`,
`JpegUsesTheOpaqueBackground`, `OptionsAreCopiedAtConstruction`,
`ResizedTransparentPageKeepsItsEdges`) FAIL; the three rejection/budget tests
and the Step 1 test pass.

### Step 4: Use the background in `ConvertAsync`

Replace the block from `using var canvas = ...` to the two `SaveAsync` calls
with code that keeps the opaque path exactly as it is today apart from the
colour, and adds the transparent path:

```csharp
        using var encoded = new ImageIoStagingStream(_options.MaxOutputBytes);
        var background = _options.Background;
        if (background.IsTransparent)
        {
            // PDFium delivers straight (non-premultiplied) BGRA; unpainted pixels have alpha 0.
            using var canvas = source.CloneAs<Rgba32>();
            if (actualWidth != width || actualHeight != height) canvas.Mutate(context => context.Resize(width, height));
            await canvas.SaveAsync(encoded, new PngEncoder { ColorType = PngColorType.RgbWithAlpha }, cancellationToken).ConfigureAwait(false);
        }
        else
        {
            using var canvas = new Image<Rgb24>(actualWidth, actualHeight, Color.FromRgb(background.Red, background.Green, background.Blue));
            canvas.Mutate(context => context.DrawImage(source, 1f));
            // PDFium rounds physical sizes to integer pixels. Normalize the at-most-one-pixel difference to the documented ceil grid.
            if (actualWidth != width || actualHeight != height) canvas.Mutate(context => context.Resize(width, height));
            if (_options.Format == OfdImageFormat.Png)
                await canvas.SaveAsync(encoded, new PngEncoder { ColorType = PngColorType.Rgb }, cancellationToken).ConfigureAwait(false);
            else
                await canvas.SaveAsync(encoded, new JpegEncoder { Quality = _options.JpegQuality }, cancellationToken).ConfigureAwait(false);
        }
        await encoded.PublishAsync(imageOutput, cancellationToken).ConfigureAwait(false);
```

(Remove the original `using var encoded` line so it is declared once.) Update
the class summary to say the page is composited onto the configured background
(white by default) or exported with its transparency.

**Verify**: build → `0 Error(s)`; `ImageBackgroundTests` → all pass; Image tests
command → `Failed: 0`; all PDF converter tests → `Failed: 0`; CLI
`ImageCommandTests` → `Failed: 0`.

### Step 5: Visual check

```sh
OUT="$(mktemp -d)"
OFDRW_TEST_ARTIFACTS="$OUT" dotnet test tests/Ofdrw.Net.Converter.Pdf.Tests/Ofdrw.Net.Converter.Pdf.Tests.csproj -c Release --no-build --filter "FullyQualifiedName~ImageBackgroundTests" --disable-build-servers -m:1 /nodeReuse:false /p:UseSharedCompilation=false
ls -la "$OUT"
```

Open the three PNGs. Expected: default — white page, solid red squares, pale
blue square top right, purple overlap; transparent — the same shapes with the
page area see-through (a viewer shows its checkerboard or window colour) and no
white halo around shapes; green — as default but on green, with the half-blue
square visibly blue-green. Record what you viewed. macOS Preview acceptance is
not possible on Linux: report "not performed".

**Verify**: three files exist; no defect seen.

### Step 6: Document

In `docs/content/guides/images.md`, after the bullet at line 29, add a bullet in
Chinese, for example:

> SDK 可通过 `OfdToImageOptions.Background` 选择背景：`OfdImageBackground.White`（默认）、`OfdImageBackground.Opaque(r, g, b)` 或 `OfdImageBackground.Transparent`。透明背景只支持 PNG，输出带 alpha 通道；JPEG 必须使用不透明背景，否则构造时抛出 `ArgumentException`。CLI 的 `ofd-to-image` 仍固定为白底（main，尚未发布）。

**Verify**: docs check commands all exit 0.

## Test plan

- New `tests/Ofdrw.Net.Converter.Pdf.Tests/ImageBackgroundTests.cs`: 1
  characterization test (Step 1) and 8 behaviour tests (Step 3). Pattern:
  `ImageIoTests.cs` in the same project.

## Done criteria

- [ ] Build prints `0 Error(s)`
- [ ] `ImageBackgroundTests` → 9 passed; `DefaultStaysOpaqueWhiteRgb` is byte-for-byte the test committed in Step 1
- [ ] All PDF converter tests and CLI `ImageCommandTests` → `Failed: 0`
- [ ] Docs check commands exit 0; storage gate exits 0 (no PNG committed)
- [ ] `plans/plan-index.md` status row updated; report lists functional tests, PNG review and "macOS Preview: not performed" separately

## STOP conditions

Stop and report back (do not improvise) if:

- The excerpts in "Current state" do not match the live code.
- `TransparentPngKeepsThePagesAlpha` shows the half-blue pixel with a blue
  channel near 127 instead of 255 (the raster would be premultiplied on your
  platform, contradicting the planning measurement). Report the raw values; do
  not add an un-premultiply step on your own.
- A transparent export shows a white or dark fringe around shapes after the
  resize in `ResizedTransparentPageKeepsItsEdges`.
- Any existing image test changes its result.
- You are tempted to detect "white" pixels and make them transparent. That is
  explicitly ruled out by the maintainer.

## Maintenance notes

- Follow-ups the maintainer listed but did not bundle: a CLI option, rotation,
  BMP, batch file naming. Images with their own alpha channel and signature
  stamps were not measured in planning; add a fixture for each when the CLI
  option is designed.
- When this ships, the release note must say the default is unchanged.
- Windows and macOS rasters come from the same PDFium build via Docnet, but
  only Linux was measured. CI runs these tests on all three; a platform
  difference will show in `TransparentPngKeepsThePagesAlpha`.
