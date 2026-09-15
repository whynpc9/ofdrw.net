using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Docnet.Core;
using Docnet.Core.Models;
using Docnet.Core.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>One lazy PDFium document session per conversion, with page-scoped image buffers.</summary>
internal sealed class DocnetPdfRasterizer : IDisposable
{
    private readonly string _pdfPath;
    private readonly int _dpi;
    private readonly long _maximumPixels;
    private IDocReader? _reader;
    private bool _openFailed;

    internal DocnetPdfRasterizer(string pdfPath, int dpi, long maximumPixels)
    {
        _pdfPath = pdfPath;
        _dpi = Math.Max(72, Math.Min(dpi, 300));
        _maximumPixels = maximumPixels;
    }

    internal string? LastFailure { get; private set; }

    internal async Task<byte[]?> TryRasterizePageAsync(int zeroBasedPageIndex, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (_openFailed) return null;
        try
        {
            if (_reader is null)
            {
                try
                {
                    _reader = DocLib.Instance.GetDocReader(_pdfPath, new PageDimensions(_dpi / 72d));
                }
                catch
                {
                    _openFailed = true;
                    throw;
                }
            }
            using var page = _reader.GetPageReader(zeroBasedPageIndex);
            var width = page.GetPageWidth();
            var height = page.GetPageHeight();
            if (width <= 0 || height <= 0 || (long)width * height > _maximumPixels)
                throw new InvalidDataException("PDF page exceeds the configured decoded pixel limit.");
            cancellationToken.ThrowIfCancellationRequested();
            var raw = page.GetImage();
            cancellationToken.ThrowIfCancellationRequested();
            using var image = Image.LoadPixelData<Bgra32>(raw, width, height);
            using var png = new MemoryStream();
            await image.SaveAsync(png, new PngEncoder(), cancellationToken).ConfigureAwait(false);
            return png.ToArray();
        }
        catch (OperationCanceledException) { throw; }
        catch (InvalidDataException) { throw; }
        catch (Exception exception) when (exception is not OutOfMemoryException)
        {
            LastFailure = exception.Message;
            return null;
        }
    }

    public void Dispose() => _reader?.Dispose();
}
