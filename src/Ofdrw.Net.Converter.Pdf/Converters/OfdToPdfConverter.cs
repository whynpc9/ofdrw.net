using System;
using System.Collections.Generic;
using System.IO;
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
        }

        document.Save(pdfOutput, false);
        await pdfOutput.FlushAsync(cancellationToken).ConfigureAwait(false);
    }

    private static bool TryDrawVectorPage(XGraphics graphics, OfdPage page)
    {
        try
        {
            foreach (var element in page.Elements)
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

                    graphics.DrawString(
                        text.Text,
                        font,
                        XBrushes.Black,
                        new XRect(
                            MillimetersToPoints(text.XMillimeters),
                            MillimetersToPoints(text.YMillimeters),
                            MillimetersToPoints(Math.Max(text.WidthMillimeters, 1)),
                            MillimetersToPoints(Math.Max(text.HeightMillimeters, text.FontSizeMillimeters * 1.6))),
                        XStringFormats.TopLeft);
                }

                if (element is OfdImageElement image && image.Data.Length > 0)
                {
                    using var xImage = XImage.FromStream(() => new MemoryStream(image.Data));
                    graphics.DrawImage(
                        xImage,
                        MillimetersToPoints(image.XMillimeters),
                        MillimetersToPoints(image.YMillimeters),
                        MillimetersToPoints(image.WidthMillimeters),
                        MillimetersToPoints(image.HeightMillimeters));
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

        foreach (var element in page.Elements)
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
                    var x = ClampInt((int)Math.Round(imageElement.XMillimeters * 6d), 0, widthPx - 1);
                    var y = ClampInt((int)Math.Round(imageElement.YMillimeters * 6d), 0, heightPx - 1);
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
}
