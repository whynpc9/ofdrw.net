using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Numerics;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using PdfSharpCore.Fonts;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace Ofdrw.Net.Converter.Pdf.Converters;

public sealed class OfdToPdfConverter : IOfdToPdfConverter
{
    private readonly OfdToPdfOptions _options;

    public OfdToPdfConverter() : this(new OfdToPdfOptions()) { }

    public OfdToPdfConverter(OfdToPdfOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        if (options.PackageLoadOptions is null || options.MaxDecodedImagePixels <= 0)
            throw new ArgumentException("OFD load options and a positive image pixel budget are required.", nameof(options));
        PdfFontRegistry.EnsureInstalled();
    }

    public async Task ConvertAsync(Stream ofdInput, Stream pdfOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default)
    {
        if (ofdInput is null)
        {
            throw new ArgumentNullException(nameof(ofdInput));
        }

        if (pdfOutput is null)
        {
            throw new ArgumentNullException(nameof(pdfOutput));
        }

        var reader = new OfdReader();
        var package = await reader.ReadAsync(ofdInput, _options.PackageLoadOptions, cancellationToken).ConfigureAwait(false);
        await ConvertPackageAsync(package, pdfOutput, pages, cancellationToken).ConfigureAwait(false);
    }

    // Shared rendering implementation; image export has already loaded and budgeted its package.
    internal async Task ConvertPackageAsync(OfdDocumentPackage package, Stream pdfOutput,
        IReadOnlyList<int>? pages, CancellationToken cancellationToken, bool strictAppearanceBudgets = false, int? maximumSignatureAppearances = null)
    {
        var orderedPages = package.Pages.OrderBy(x => x.Index).ToList();
        var selected = OfdPageSelection.Normalize(orderedPages.Count, pages);
        var selectedIds = strictAppearanceBudgets
            ? new HashSet<string>(selected.Select(index => orderedPages[index].Id ?? string.Empty), StringComparer.OrdinalIgnoreCase)
            : null;
        var fonts = new DocumentFontContext(package.Fonts);
        var signatureAppearances = await PrepareSignatureAppearancesAsync(
                package, selectedIds, strictAppearanceBudgets, maximumSignatureAppearances,
                cancellationToken)
            .ConfigureAwait(false);

        using var document = new PdfDocument();
        foreach (var selectedIndex in selected)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var pageModel = orderedPages[selectedIndex];
            if (!(pageModel.WidthMillimeters > 0) || !(pageModel.HeightMillimeters > 0) ||
                double.IsInfinity(pageModel.WidthMillimeters) || double.IsInfinity(pageModel.HeightMillimeters))
                throw new InvalidDataException("OFD page has invalid physical dimensions.");
            var pdfPage = document.AddPage();
            pdfPage.Width = MillimetersToPoints(pageModel.WidthMillimeters);
            pdfPage.Height = MillimetersToPoints(pageModel.HeightMillimeters);

            using var graphics = XGraphics.FromPdfPage(pdfPage);
            DrawPage(graphics, pageModel, fonts, _options.MaxDecodedImagePixels, cancellationToken);

            DrawSignatureAppearances(
                document,
                graphics,
                pageModel,
                signatureAppearances, _options.MaxDecodedImagePixels, cancellationToken, strictAppearanceBudgets);
        }

        cancellationToken.ThrowIfCancellationRequested();
        document.Save(pdfOutput, false);
        await pdfOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private async Task<IReadOnlyList<PreparedSignatureAppearance>> PrepareSignatureAppearancesAsync(
        OfdDocumentPackage package, HashSet<string>? selectedIds, bool strictAppearanceBudgets,
        int? maximumSignatureAppearances, CancellationToken cancellationToken)
    {
        var result = new List<PreparedSignatureAppearance>();
        var nestedCache = new Dictionary<string, (OfdDocumentPackage? Package, DocumentFontContext? Fonts)>();
        var payloadIdentities = new Dictionary<byte[], string>();
        var usedBytes = strictAppearanceBudgets ? package.PreservedEntries.Values.Sum(data => (long)data.Length) : 0;
        var usedEntries = strictAppearanceBudgets ? package.PreservedEntries.Count : 0;
        var usedPages = strictAppearanceBudgets ? package.Pages.Count : 0;
        foreach (var appearance in OfdSignatureAppearanceReader.Read(package, selectedIds, maximumSignatureAppearances, cancellationToken))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsZip(appearance.Data))
            {
                // Cache only on the image path; preserve legacy PDF appearance tolerance and accounting.
                var key = string.Empty;
                if (strictAppearanceBudgets && !payloadIdentities.TryGetValue(appearance.Data, out key))
                {
                    key = BinaryIdentity.Hash(appearance.Data);
                    payloadIdentities.Add(appearance.Data, key);
                }
                if (!strictAppearanceBudgets || !nestedCache.TryGetValue(key, out var cached))
                {
                    var remaining = strictAppearanceBudgets
                        ? RemainingAppearanceBudget(usedBytes, usedEntries, usedPages)
                        : null;
                    var nested = await ReadAppearanceOfdPackageAsync(appearance.Data, remaining, cancellationToken).ConfigureAwait(false);
                    cached = (nested, nested is null ? null : new DocumentFontContext(nested.Fonts));
                    if (strictAppearanceBudgets)
                    {
                        nestedCache.Add(key, cached);
                        if (nested is not null)
                        {
                            usedBytes += nested.PreservedEntries.Values.Sum(data => (long)data.Length);
                            usedEntries += nested.PreservedEntries.Count;
                            usedPages += nested.Pages.Count;
                        }
                    }
                }
                if (cached.Package is not null) result.Add(new PreparedSignatureAppearance(appearance, cached.Package, cached.Fonts!));
                continue;
            }
            result.Add(new PreparedSignatureAppearance(appearance, appearance.Data));
        }
        return result;
    }

    private OfdPackageLoadOptions RemainingAppearanceBudget(long usedBytes, int usedEntries, int usedPages)
    {
        var load = _options.PackageLoadOptions;
        if (usedBytes >= load.MaxTotalUncompressedBytes || usedEntries >= load.MaxEntryCount || usedPages >= load.MaxPageCount)
            throw new InvalidDataException("Cumulative OFD signature appearance expansion budget exceeded.");
        return new OfdPackageLoadOptions
        {
            MaxInputBytes = load.MaxInputBytes, MaxEntryUncompressedBytes = load.MaxEntryUncompressedBytes,
            MaxCompressionRatio = load.MaxCompressionRatio,
            MaxTotalUncompressedBytes = load.MaxTotalUncompressedBytes - usedBytes,
            MaxEntryCount = load.MaxEntryCount - usedEntries,
            MaxPageCount = load.MaxPageCount - usedPages
        };
    }

    private async Task<OfdDocumentPackage?> ReadAppearanceOfdPackageAsync(
        byte[] ofdData, OfdPackageLoadOptions? appearanceBudget,
        CancellationToken cancellationToken)
    {
        try
        {
            var reader = new OfdReader();
            using var input = new MemoryStream(ofdData, writable: false);
            var package = await reader
                .ReadAsync(input, appearanceBudget ?? new OfdPackageLoadOptions(), cancellationToken)
                .ConfigureAwait(false);
            var pageModel = package.Pages
                .OrderBy(page => page.Index)
                .FirstOrDefault();
            if (pageModel is null ||
                pageModel.WidthMillimeters <= 0 ||
                pageModel.HeightMillimeters <= 0)
            {
                return null;
            }

            return package;
        }
        catch (OperationCanceledException) { throw; }
        catch (InvalidDataException) when (appearanceBudget is not null) { throw; }
        catch
        {
            return null;
        }
    }

    private static void DrawSignatureAppearances(PdfDocument document, XGraphics graphics, OfdPage page,
        IReadOnlyList<PreparedSignatureAppearance> appearances, long maximumPixels,
        CancellationToken cancellationToken, bool reusePayloads = false)
    {
        var forms = new Dictionary<OfdPage, XForm>();
        using var bitmap = new SinglePayloadResource<XImage>(data =>
        {
            ValidateImage(data, maximumPixels);
            return XImage.FromImageSource(new EncodedBitmapImageSource(data, maximumPixels, cancellationToken));
        });
        try
        {
            foreach (var appearance in appearances.Where(item => string.Equals(item.PageId, page.Id, StringComparison.OrdinalIgnoreCase)))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var target = new XRect(MillimetersToPoints(appearance.XMillimeters - page.XMillimeters),
                    MillimetersToPoints(appearance.YMillimeters - page.YMillimeters),
                    MillimetersToPoints(appearance.WidthMillimeters), MillimetersToPoints(appearance.HeightMillimeters));
                try
                {
                    if (appearance.OfdPage is not null)
                    {
                        if (!reusePayloads)
                        {
                            using var form = CreateAppearanceForm(document, appearance, maximumPixels, cancellationToken);
                            graphics.DrawImage(form, target);
                        }
                        else
                        {
                            if (!forms.TryGetValue(appearance.OfdPage, out var form))
                            {
                                form = CreateAppearanceForm(document, appearance, maximumPixels, cancellationToken);
                                forms.Add(appearance.OfdPage, form);
                            }
                            graphics.DrawImage(form, target);
                        }
                    }
                    else if (!reusePayloads)
                    {
                        ValidateImage(appearance.Data, maximumPixels);
                        using var image = XImage.FromStream(() => new MemoryStream(appearance.Data, writable: false));
                        graphics.DrawImage(image, target);
                    }
                    else graphics.DrawImage(bitmap.Get(appearance.Data), target);
                }
                catch (OperationCanceledException) { throw; }
                catch (InvalidDataException) { throw; }
                catch { /* Unsupported vendor payloads retain the existing preview fallback. */ }
            }
        }
        finally
        {
            foreach (var form in forms.Values) form.Dispose();
        }
    }

    private static XForm CreateAppearanceForm(PdfDocument document, PreparedSignatureAppearance appearance,
        long maximumPixels, CancellationToken token)
    {
        var form = new XForm(document, XUnit.FromMillimeter(appearance.OfdPage!.WidthMillimeters),
            XUnit.FromMillimeter(appearance.OfdPage.HeightMillimeters));
        try
        {
            using var graphics = XGraphics.FromForm(form);
            DrawPage(graphics, appearance.OfdPage, appearance.Fonts!, maximumPixels, token);
            return form;
        }
        catch { form.Dispose(); throw; }
    }

    private static bool IsZip(byte[] data)
    {
        return data.Length >= 4 &&
            data[0] == 0x50 &&
            data[1] == 0x4b &&
            data[2] == 0x03 &&
            data[3] == 0x04;
    }

    private static void DrawPage(XGraphics graphics, OfdPage page, DocumentFontContext fonts,
        long maximumPixels, CancellationToken cancellationToken)
    {
        try
        {
            var outlineFonts = fonts.OutlineFonts;
            foreach (var element in EnumerateRenderableElements(page))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (element is OfdTextElement text)
                {
                    var fontSize = Math.Max(0.1, MillimetersToPoints(text.FontSizeMillimeters));
                    var familyName = fonts.Resolve(text, out var resource);
                    // CT_Text Weight/Italic is the per-object style viewers apply;
                    // the resource flags describe the bound font file.
                    var bold = resource?.Bold == true || text.Weight >= 600;
                    var italic = resource?.Italic == true || text.Italic;
                    var style = (bold ? XFontStyle.Bold : XFontStyle.Regular) |
                        (italic ? XFontStyle.Italic : XFontStyle.Regular);
                    XFont font;
                    FontResolverInfo face;
                    try
                    {
                        font = new XFont(familyName, fontSize, style);
                        face = GlobalFontSettings.FontResolver.ResolveTypeface(familyName, bold, italic);
                    }
                    catch (Exception exception) when (resource?.Data.Length > 0)
                    {
                        throw new InvalidDataException($"Embedded font '{resource.FontName}' could not be initialized.", exception);
                    }
                    catch (Exception exception) when (exception is not OutOfMemoryException &&
                                                       exception is not OperationCanceledException)
                    {
                        font = new XFont("Arial", fontSize);
                        // Use the face actually drawn. Re-querying the rejected
                        // name here would repeat the host failure after fallback.
                        face = GlobalFontSettings.FontResolver.ResolveTypeface("Arial", false, false);
                    }
                    var simulateBold = face.MustSimulateBold;
                    var simulateItalic = face.MustSimulateItalic;
                    SixLabors.Fonts.Font? outlineFont = null;
                    if (simulateBold)
                    {
                        if (!outlineFonts.TryGetValue(face.FaceName, out var family))
                        {
                            using var fontStream = new MemoryStream(GlobalFontSettings.FontResolver.GetFont(face.FaceName));
                            family = new SixLabors.Fonts.FontCollection().Add(fontStream);
                            outlineFonts.Add(face.FaceName, family);
                        }
                        outlineFont = family.CreateFont((float)fontSize);
                    }
                    var brush = new XSolidBrush(XColor.FromArgb(
                        text.FillColor.Alpha,
                        text.FillColor.Red,
                        text.FillColor.Green,
                        text.FillColor.Blue));
                    if (text.Runs.Count > 0)
                    {
                        DrawTextRuns(
                            graphics,
                            text,
                            font,
                            brush,
                            page.XMillimeters,
                            page.YMillimeters,
                            outlineFont,
                            simulateItalic);
                    }
                    else
                    {
                        DrawStyledString(graphics, text.Text, font, brush,
                            new XPoint(MillimetersToPoints(text.XMillimeters - page.XMillimeters),
                                MillimetersToPoints(text.YMillimeters - page.YMillimeters)),
                            outlineFont, simulateItalic, XStringFormats.TopLeft);
                    }
                }

                if (element is OfdImageElement image && image.Data.Length > 0)
                {
                    DrawImage(graphics, page, image, maximumPixels);
                }

                if (element is OfdPathElement path)
                {
                    OfdPathRenderer.TryDraw(
                        graphics,
                        path,
                        page.XMillimeters,
                        page.YMillimeters);
                }
            }

        }
        catch (OperationCanceledException) { throw; }
        catch (InvalidDataException) { throw; }
        catch (NotSupportedException) { throw; }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            throw new InvalidDataException($"OFD page {page.Index + 1} could not be rendered without content loss.", exception);
        }
    }

    private static void DrawImage(XGraphics graphics, OfdPage page, OfdImageElement image, long maximumPixels)
    {
        ValidateImage(image.Data, maximumPixels);
        var state = graphics.Save();
        try
        {
            graphics.TranslateTransform(MillimetersToPoints(image.XMillimeters - page.XMillimeters),
                MillimetersToPoints(image.YMillimeters - page.YMillimeters));
            graphics.IntersectClip(new XRect(0, 0, MillimetersToPoints(image.WidthMillimeters), MillimetersToPoints(image.HeightMillimeters)));
            foreach (var clip in OfdClipGeometry.Read(image.ClipsXml))
                graphics.IntersectClip(OfdPathRenderer.BuildClip(clip));
            var matrix = image.Transform ?? new double[] { image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0 };
            if (matrix.Length != 6 || matrix.Any(value => double.IsNaN(value) || double.IsInfinity(value)))
                throw new InvalidDataException("Invalid OFD image transform.");
            graphics.MultiplyTransform(new XMatrix(MillimetersToPoints(matrix[0]), MillimetersToPoints(matrix[1]),
                MillimetersToPoints(matrix[2]), MillimetersToPoints(matrix[3]), MillimetersToPoints(matrix[4]), MillimetersToPoints(matrix[5])));
            var data = image.Data;
            if (image.Alpha < 255)
            {
                using var pixels = Image.Load<Rgba32>(data);
                pixels.Mutate(context => context.Opacity(Math.Max(0, image.Alpha) / 255f));
                using var png = new MemoryStream();
                pixels.SaveAsPng(png);
                data = png.ToArray();
            }
            using var xImage = XImage.FromStream(() => new MemoryStream(data, writable: false));
            graphics.DrawImage(xImage, 0, 0, 1, 1);
        }
        finally
        {
            graphics.Restore(state);
        }
    }

    private static void ValidateImage(byte[] data, long maximumPixels)
    {
        var info = Image.Identify(data);
        if (info is null || info.Width <= 0 || info.Height <= 0 || (long)info.Width * info.Height > maximumPixels)
            throw new InvalidDataException("OFD image exceeds the configured decoded pixel limit or has invalid dimensions.");
    }

    private static double MillimetersToPoints(double millimeters)
    {
        return millimeters * 72d / 25.4d;
    }

    private static void DrawStyledString(XGraphics graphics, string text, XFont font, XBrush brush,
        XPoint point, SixLabors.Fonts.Font? outlineFont, bool italic, XStringFormat? format = null)
    {
        // PDFsharp Core 1.3.67 drops resolver style simulations when creating
        // XGlyphTypeface. Apply the missing fallback appearance at draw time.
        var state = graphics.Save();
        try
        {
            // PDFsharp Core can omit the zero-alpha graphics state for text
            // after an image. Keep semantic text extractable, but guarantee
            // that a fully transparent fill cannot paint any pixels.
            if (brush is XSolidBrush transparent && transparent.Color.A == 0)
                graphics.IntersectClip(new XRect(0, 0, 0, 0));
            graphics.TranslateTransform(point.X, point.Y);
            if (italic)
            {
                graphics.MultiplyTransform(new XMatrix(1, 0, -0.21256, 1, 0, 0));
            }
            graphics.DrawString(text, font, brush, new XRect(0, 0, 0, 0), format ?? XStringFormats.Default);
            if (outlineFont is not null && brush is XSolidBrush solid)
            {
                var topBearing = 0;
                for (var index = 0; index < text.Length; index++)
                {
                    var codePoint = char.ConvertToUtf32(text, index);
                    if (char.IsHighSurrogate(text[index])) index++;
                    if (outlineFont.TryGetGlyphs(new SixLabors.Fonts.Unicode.CodePoint(codePoint),
                            SixLabors.Fonts.ColorFontSupport.None, out var glyphs))
                    {
                        foreach (var glyph in glyphs)
                            topBearing = Math.Min(topBearing, glyph.GlyphMetrics.TopSideBearing);
                    }
                }
                // Fonts 1.0.1 centers the em box within the horizontal line metrics,
                // then shifts the ascender for negative top bearings. Undo both
                // offsets; otherwise the outline and PDF text have different baselines.
                var metrics = outlineFont.FontMetrics;
                var layoutAscender = metrics.HorizontalMetrics.Ascender - topBearing -
                    (metrics.HorizontalMetrics.LineHeight - metrics.UnitsPerEm) * 0.5f;
                var baselineOffset = -layoutAscender * outlineFont.Size / metrics.UnitsPerEm;
                if (format == XStringFormats.TopLeft)
                {
                    // Use the same top-to-baseline distance as PDFsharp DrawString.
                    baselineOffset += (float)(font.GetHeight() * font.CellAscent / font.CellSpace);
                }
                SixLabors.Fonts.TextRenderer.RenderTextTo(
                    new PdfGlyphOutlineRenderer(graphics, solid.Color, font.Size * 0.025), text,
                    new SixLabors.Fonts.TextOptions(outlineFont)
                    {
                        Dpi = 72,
                        Origin = new Vector2(0, baselineOffset),
                        KerningMode = SixLabors.Fonts.KerningMode.None
                    });
            }
        }
        finally
        {
            graphics.Restore(state);
        }
    }

    private static void DrawTextRuns(
        XGraphics graphics,
        OfdTextElement text,
        XFont font,
        XBrush brush,
        double pageOriginX,
        double pageOriginY,
        SixLabors.Fonts.Font? outlineFont,
        bool simulateItalic)
    {
        foreach (var run in text.Runs)
        {
            var glyphs = OfdTextGeometry.Glyphs(run.Text);
            var deltaX = OfdTextGeometry.ExpandDeltas(run.DeltaX, Math.Max(0, glyphs.Count - 1));
            var deltaY = OfdTextGeometry.ExpandDeltas(run.DeltaY, Math.Max(0, glyphs.Count - 1));
            if (deltaX.Count == 0 && deltaY.Count == 0)
            {
                var point = TransformTextPoint(
                    text,
                    run.XMillimeters,
                    run.YMillimeters,
                    pageOriginX,
                    pageOriginY);
                DrawStyledString(graphics, run.Text, font, brush, point, outlineFont, simulateItalic);
                continue;
            }

            var x = run.XMillimeters;
            var y = run.YMillimeters;
            for (var i = 0; i < glyphs.Count; i++)
            {
                var point = TransformTextPoint(
                    text,
                    x,
                    y,
                    pageOriginX,
                    pageOriginY);
                DrawStyledString(graphics, glyphs[i], font, brush, point, outlineFont, simulateItalic);
                if (i < deltaX.Count)
                {
                    x += deltaX[i];
                }

                if (i < deltaY.Count)
                {
                    y += deltaY[i];
                }
            }
        }
    }

    private static XPoint TransformTextPoint(
        OfdTextElement text,
        double x,
        double y,
        double pageOriginX,
        double pageOriginY)
    {
        if (text.Transform is { Length: 6 } matrix)
        {
            var transformedX = (matrix[0] * x) + (matrix[2] * y) + matrix[4];
            var transformedY = (matrix[1] * x) + (matrix[3] * y) + matrix[5];
            x = transformedX;
            y = transformedY;
        }

        return new XPoint(
            MillimetersToPoints(text.XMillimeters + x - pageOriginX),
            MillimetersToPoints(text.YMillimeters + y - pageOriginY));
    }

    private static IEnumerable<OfdElement> EnumerateRenderableElements(OfdPage page)
    {
        foreach (var template in page.Templates.Where(x =>
            string.Equals(x.ZOrder, "Background", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var element in template.Elements)
            {
                yield return element;
            }
        }

        foreach (var element in page.Elements)
        {
            yield return element;
        }

        foreach (var template in page.Templates.Where(x =>
            !string.Equals(x.ZOrder, "Background", StringComparison.OrdinalIgnoreCase)))
        {
            foreach (var element in template.Elements)
            {
                yield return element;
            }
        }
    }

    private sealed class PreparedSignatureAppearance
    {
        public PreparedSignatureAppearance(
            OfdSignatureAppearance source,
            OfdDocumentPackage package, DocumentFontContext fonts)
        {
            PageId = source.PageId;
            XMillimeters = source.XMillimeters;
            YMillimeters = source.YMillimeters;
            WidthMillimeters = source.WidthMillimeters;
            HeightMillimeters = source.HeightMillimeters;
            Data = Array.Empty<byte>();
            OfdPage = package.Pages.OrderBy(page => page.Index).First();
            Fonts = fonts;
        }

        public PreparedSignatureAppearance(
            OfdSignatureAppearance source,
            byte[] data)
        {
            PageId = source.PageId;
            XMillimeters = source.XMillimeters;
            YMillimeters = source.YMillimeters;
            WidthMillimeters = source.WidthMillimeters;
            HeightMillimeters = source.HeightMillimeters;
            Data = data;
        }

        public string PageId { get; }

        public double XMillimeters { get; }

        public double YMillimeters { get; }

        public double WidthMillimeters { get; }

        public double HeightMillimeters { get; }

        public byte[] Data { get; }

        public OfdPage? OfdPage { get; }
        internal DocumentFontContext? Fonts { get; }
    }
}
