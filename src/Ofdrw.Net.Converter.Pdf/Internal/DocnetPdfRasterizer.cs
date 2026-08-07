using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Docnet.Core;
using Docnet.Core.Models;
using PdfSharpCore.Pdf.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Rasterizes PDF pages with Docnet/Pdfium for high layout fidelity (tables, grids).
/// </summary>
internal sealed class DocnetPdfRasterizer
{
    private const double PointsPerInch = 72.0;

    public async Task<byte[]?> TryRasterizePageAsync(
        byte[] pdfBytes,
        int zeroBasedPageIndex,
        int dpi,
        CancellationToken cancellationToken)
    {
        if (pdfBytes is null || pdfBytes.Length == 0)
        {
            return null;
        }

        dpi = Math.Max(72, Math.Min(dpi, 300));

        try
        {
            using var pdf = PdfReader.Open(new MemoryStream(pdfBytes, writable: false), PdfDocumentOpenMode.Import);
            if (zeroBasedPageIndex < 0 || zeroBasedPageIndex >= pdf.PageCount)
            {
                return null;
            }

            var page = pdf.Pages[zeroBasedPageIndex];
            var widthPx = ToPixels(page.Width.Point, dpi);
            var heightPx = ToPixels(page.Height.Point, dpi);

            using var reader = DocLib.Instance.GetDocReader(pdfBytes, new PageDimensions(widthPx, heightPx));
            using var pageReader = reader.GetPageReader(zeroBasedPageIndex);
            var raw = pageReader.GetImage();
            var pixelWidth = pageReader.GetPageWidth();
            var pixelHeight = pageReader.GetPageHeight();
            if (raw is null || raw.Length == 0 || pixelWidth <= 0 || pixelHeight <= 0)
            {
                return null;
            }

            using var image = Image.LoadPixelData<Bgra32>(raw, pixelWidth, pixelHeight);
            using var pngStream = new MemoryStream();
            await image.SaveAsync(pngStream, new PngEncoder(), cancellationToken).ConfigureAwait(false);
            return pngStream.ToArray();
        }
        catch (OperationCanceledException)
        {
            throw;
        }
        catch
        {
            return null;
        }
    }

    private static int ToPixels(double points, int dpi)
    {
        return Math.Max(1, (int)Math.Ceiling(points * dpi / PointsPerInch));
    }
}
