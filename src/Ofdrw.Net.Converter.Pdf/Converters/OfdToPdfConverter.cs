using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace Ofdrw.Net.Converter.Pdf.Converters;

public sealed class OfdToPdfConverter : IOfdToPdfConverter
{
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
        var package = await reader.ReadAsync(ofdInput, cancellationToken).ConfigureAwait(false);
        RegisterEmbeddedFonts(package);
        var signatureAppearances = await PrepareSignatureAppearancesAsync(
                package,
                cancellationToken)
            .ConfigureAwait(false);

        var orderedPages = package.Pages.OrderBy(x => x.Index).ToList();
        var selected = PageSelection.Normalize(orderedPages.Count, pages);

        using var document = new PdfDocument();
        foreach (var selectedIndex in selected)
        {
            var pageModel = orderedPages[selectedIndex];
            var pdfPage = document.AddPage();
            pdfPage.Width = MillimetersToPoints(pageModel.WidthMillimeters);
            pdfPage.Height = MillimetersToPoints(pageModel.HeightMillimeters);

            using var graphics = XGraphics.FromPdfPage(pdfPage);
            var drawn = TryDrawVectorPage(graphics, pageModel);
            if (!drawn)
            {
                DrawRasterFallback(graphics, pageModel);
            }

            DrawSignatureAppearances(
                document,
                graphics,
                pageModel,
                signatureAppearances);
        }

        document.Save(pdfOutput, false);
        await pdfOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static void RegisterEmbeddedFonts(OfdDocumentPackage package)
    {
        OfdEmbeddedFontResolver.EnsureInstalled();
        foreach (var font in package.Fonts.Where(font => font.Data.Length > 0))
        {
            OfdEmbeddedFontResolver.Register(
                font.FontName,
                font.Data,
                font.Bold,
                font.Italic);
            if (!string.IsNullOrWhiteSpace(font.FamilyName) &&
                !string.Equals(
                    font.FamilyName,
                    font.FontName,
                    StringComparison.OrdinalIgnoreCase))
            {
                OfdEmbeddedFontResolver.Register(
                    font.FamilyName!,
                    font.Data,
                    font.Bold,
                    font.Italic);
            }
        }
    }

    private static async Task<IReadOnlyList<PreparedSignatureAppearance>>
        PrepareSignatureAppearancesAsync(
            OfdDocumentPackage package,
            CancellationToken cancellationToken)
    {
        var result = new List<PreparedSignatureAppearance>();
        foreach (var appearance in OfdSignatureAppearanceReader.Read(package))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (IsZip(appearance.Data))
            {
                var appearancePage = await ReadAppearanceOfdPageAsync(
                        appearance.Data,
                        cancellationToken)
                    .ConfigureAwait(false);
                if (appearancePage is not null)
                {
                    result.Add(
                        new PreparedSignatureAppearance(
                            appearance,
                            appearancePage));
                }

                continue;
            }

            result.Add(
                new PreparedSignatureAppearance(
                    appearance,
                    appearance.Data));
        }

        return result;
    }

    private static async Task<OfdPage?> ReadAppearanceOfdPageAsync(
        byte[] ofdData,
        CancellationToken cancellationToken)
    {
        try
        {
            var reader = new OfdReader();
            using var input = new MemoryStream(ofdData, writable: false);
            var package = await reader
                .ReadAsync(input, cancellationToken)
                .ConfigureAwait(false);
            RegisterEmbeddedFonts(package);
            var pageModel = package.Pages
                .OrderBy(page => page.Index)
                .FirstOrDefault();
            if (pageModel is null ||
                pageModel.WidthMillimeters <= 0 ||
                pageModel.HeightMillimeters <= 0)
            {
                return null;
            }

            return pageModel;
        }
        catch
        {
            return null;
        }
    }

    private static void DrawSignatureAppearances(
        PdfDocument document,
        XGraphics graphics,
        OfdPage page,
        IReadOnlyList<PreparedSignatureAppearance> appearances)
    {
        foreach (var appearance in appearances.Where(item =>
            string.Equals(
                item.PageId,
                page.Id,
                StringComparison.OrdinalIgnoreCase)))
        {
            var target = new XRect(
                MillimetersToPoints(
                    appearance.XMillimeters - page.XMillimeters),
                MillimetersToPoints(
                    appearance.YMillimeters - page.YMillimeters),
                MillimetersToPoints(appearance.WidthMillimeters),
                MillimetersToPoints(appearance.HeightMillimeters));
            try
            {
                if (appearance.OfdPage is not null)
                {
                    using var form = new XForm(
                        document,
                        XUnit.FromMillimeter(
                            appearance.OfdPage.WidthMillimeters),
                        XUnit.FromMillimeter(
                            appearance.OfdPage.HeightMillimeters));
                    using (var formGraphics = XGraphics.FromForm(form))
                    {
                        var drawn = TryDrawVectorPage(
                            formGraphics,
                            appearance.OfdPage);
                        if (!drawn)
                        {
                            DrawRasterFallback(
                                formGraphics,
                                appearance.OfdPage);
                        }
                    }

                    graphics.DrawImage(form, target);
                }
                else
                {
                    using var image = XImage.FromStream(
                        () => new MemoryStream(
                            appearance.Data,
                            writable: false));
                    graphics.DrawImage(image, target);
                }
            }
            catch
            {
                // Unsupported vendor seal payloads do not prevent conversion of
                // the signed document body.
            }
        }
    }

    private static bool IsZip(byte[] data)
    {
        return data.Length >= 4 &&
            data[0] == 0x50 &&
            data[1] == 0x4b &&
            data[2] == 0x03 &&
            data[3] == 0x04;
    }

    private static bool TryDrawVectorPage(XGraphics graphics, OfdPage page)
    {
        try
        {
            foreach (var element in EnumerateRenderableElements(page))
            {
                if (element is OfdTextElement text)
                {
                    var fontSize = Math.Max(6, MillimetersToPoints(text.FontSizeMillimeters));
                    XFont font;
                    try
                    {
                        font = new XFont(string.IsNullOrWhiteSpace(text.FontName) ? "Arial" : text.FontName, fontSize);
                    }
                    catch
                    {
                        font = new XFont("Arial", fontSize);
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
                            page.YMillimeters);
                    }
                    else
                    {
                        graphics.DrawString(
                            text.Text,
                            font,
                            brush,
                            new XRect(
                                MillimetersToPoints(text.XMillimeters - page.XMillimeters),
                                MillimetersToPoints(text.YMillimeters - page.YMillimeters),
                                MillimetersToPoints(Math.Max(text.WidthMillimeters, 1)),
                                MillimetersToPoints(Math.Max(text.HeightMillimeters, text.FontSizeMillimeters * 1.6))),
                            XStringFormats.TopLeft);
                    }
                }

                if (element is OfdImageElement image && image.Data.Length > 0)
                {
                    using var xImage = XImage.FromStream(() => new MemoryStream(image.Data));
                    graphics.DrawImage(
                        xImage,
                        MillimetersToPoints(image.XMillimeters - page.XMillimeters),
                        MillimetersToPoints(image.YMillimeters - page.YMillimeters),
                        MillimetersToPoints(image.WidthMillimeters),
                        MillimetersToPoints(image.HeightMillimeters));
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

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void DrawRasterFallback(XGraphics graphics, OfdPage page)
    {
        var widthPx = Math.Max((int)Math.Ceiling(page.WidthMillimeters * 6d), 64);
        var heightPx = Math.Max((int)Math.Ceiling(page.HeightMillimeters * 6d), 64);

        using var canvas = new Image<Rgba32>(widthPx, heightPx, Color.White);

        foreach (var element in EnumerateRenderableElements(page))
        {
            if (element is OfdTextElement text)
            {
                // Text fallback keeps page blank if text drawing is unavailable.
                _ = text;
            }

            if (element is OfdImageElement imageElement && imageElement.Data.Length > 0)
            {
                try
                {
                    using var overlay = Image.Load<Rgba32>(imageElement.Data);
                    var x = ClampInt((int)Math.Round(
                        (imageElement.XMillimeters - page.XMillimeters) * 6d), 0, widthPx - 1);
                    var y = ClampInt((int)Math.Round(
                        (imageElement.YMillimeters - page.YMillimeters) * 6d), 0, heightPx - 1);
                    var w = ClampInt((int)Math.Round(Math.Max(imageElement.WidthMillimeters, 2) * 6d), 1, widthPx - x);
                    var h = ClampInt((int)Math.Round(Math.Max(imageElement.HeightMillimeters, 2) * 6d), 1, heightPx - y);

                    overlay.Mutate(ctx => ctx.Resize(w, h));
                    canvas.Mutate(ctx => ctx.DrawImage(overlay, new Point(x, y), 1f));
                }
                catch
                {
                    // ignored
                }
            }
        }

        using var ms = new MemoryStream();
        canvas.SaveAsPng(ms);
        var imageBytes = ms.ToArray();

        using var xImage = XImage.FromStream(() => new MemoryStream(imageBytes));
        graphics.DrawImage(
            xImage,
            0,
            0,
            MillimetersToPoints(page.WidthMillimeters),
            MillimetersToPoints(page.HeightMillimeters));
    }

    private static double MillimetersToPoints(double millimeters)
    {
        return millimeters * 72d / 25.4d;
    }

    private static void DrawTextRuns(
        XGraphics graphics,
        OfdTextElement text,
        XFont font,
        XBrush brush,
        double pageOriginX,
        double pageOriginY)
    {
        foreach (var run in text.Runs)
        {
            var deltaX = ExpandTextDeltas(run.DeltaX);
            var deltaY = ExpandTextDeltas(run.DeltaY);
            if (deltaX.Count == 0 && deltaY.Count == 0)
            {
                var point = TransformTextPoint(
                    text,
                    run.XMillimeters,
                    run.YMillimeters,
                    pageOriginX,
                    pageOriginY);
                graphics.DrawString(run.Text, font, brush, point);
                continue;
            }

            var x = run.XMillimeters;
            var y = run.YMillimeters;
            var glyphs = new List<string>();
            var enumerator = StringInfo.GetTextElementEnumerator(run.Text);
            while (enumerator.MoveNext())
            {
                glyphs.Add(enumerator.GetTextElement());
            }

            for (var i = 0; i < glyphs.Count; i++)
            {
                var point = TransformTextPoint(
                    text,
                    x,
                    y,
                    pageOriginX,
                    pageOriginY);
                graphics.DrawString(glyphs[i], font, brush, point);
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

    private static IReadOnlyList<double> ExpandTextDeltas(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return Array.Empty<double>();
        }

        var tokens = value!.Split(
            new[] { ' ', '\t', '\r', '\n' },
            StringSplitOptions.RemoveEmptyEntries);
        var result = new List<double>();
        for (var i = 0; i < tokens.Length;)
        {
            if (string.Equals(tokens[i], "g", StringComparison.OrdinalIgnoreCase) &&
                i + 2 < tokens.Length &&
                int.TryParse(tokens[i + 1], NumberStyles.Integer, CultureInfo.InvariantCulture, out var count) &&
                double.TryParse(tokens[i + 2], NumberStyles.Float, CultureInfo.InvariantCulture, out var repeated))
            {
                for (var repeat = 0; repeat < Math.Max(count, 0); repeat++)
                {
                    result.Add(repeated);
                }

                i += 3;
                continue;
            }

            if (double.TryParse(tokens[i], NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed))
            {
                result.Add(parsed);
            }

            i++;
        }

        return result;
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

    private static int ClampInt(int value, int min, int max)
    {
        if (value < min)
        {
            return min;
        }

        if (value > max)
        {
            return max;
        }

        return value;
    }

    private sealed class PreparedSignatureAppearance
    {
        public PreparedSignatureAppearance(
            OfdSignatureAppearance source,
            OfdPage ofdPage)
        {
            PageId = source.PageId;
            XMillimeters = source.XMillimeters;
            YMillimeters = source.YMillimeters;
            WidthMillimeters = source.WidthMillimeters;
            HeightMillimeters = source.HeightMillimeters;
            Data = Array.Empty<byte>();
            OfdPage = ofdPage;
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
    }
}
