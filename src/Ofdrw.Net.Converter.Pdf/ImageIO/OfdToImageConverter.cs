using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;
using SixLabors.ImageSharp.Processing;

namespace Ofdrw.Net.Converter.Pdf.Converters;

/// <summary>Exports one zero-based OFD page through the existing PDF renderer and PDFium, composited onto white.</summary>
public sealed class OfdToImageConverter
{
    private readonly OfdToImageOptions _options;
    /// <summary>Creates a PNG exporter at 144 DPI with default budgets.</summary>
    public OfdToImageConverter() : this(new OfdToImageOptions()) { }
    /// <summary>Copies and validates resolution, encoding and budgets.</summary>
    public OfdToImageConverter(OfdToImageOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        ImageIoBudget.Positive(options.PixelsPerMillimeter, nameof(options.PixelsPerMillimeter));
        if (!Enum.IsDefined(typeof(OfdImageFormat), options.Format) || options.JpegQuality < 1 || options.JpegQuality > 100 ||
            options.MaxPixels <= 0 || options.MaxRasterWorkingBytes <= 0 || options.MaxIntermediatePdfBytes <= 0 || options.MaxOutputBytes <= 0)
            throw new ArgumentException("Encoding, quality and positive image budgets are required.", nameof(options));
        var load = options.PackageLoadOptions ?? throw new ArgumentException("Package load budgets are required.", nameof(options));
        if (load.MaxInputBytes <= 0 || load.MaxEntryCount <= 0 || load.MaxPageCount <= 0 || load.MaxEntryUncompressedBytes <= 0 ||
            load.MaxTotalUncompressedBytes <= 0 || !(load.MaxCompressionRatio > 0) || double.IsInfinity(load.MaxCompressionRatio))
            throw new ArgumentException("Positive finite package load budgets are required.", nameof(options));
        _options = new OfdToImageOptions
        {
            PixelsPerMillimeter = options.PixelsPerMillimeter, Format = options.Format, JpegQuality = options.JpegQuality,
            MaxPixels = options.MaxPixels, MaxRasterWorkingBytes = options.MaxRasterWorkingBytes,
            MaxIntermediatePdfBytes = options.MaxIntermediatePdfBytes, MaxOutputBytes = options.MaxOutputBytes,
            PackageLoadOptions = new OfdPackageLoadOptions
            {
                MaxInputBytes = load.MaxInputBytes, MaxEntryCount = load.MaxEntryCount, MaxPageCount = load.MaxPageCount,
                MaxEntryUncompressedBytes = load.MaxEntryUncompressedBytes, MaxTotalUncompressedBytes = load.MaxTotalUncompressedBytes,
                MaxCompressionRatio = load.MaxCompressionRatio
            }
        };
    }

    /// <summary>Encodes the selected page at ceil(mm × ppm) pixels per axis. Leaves caller streams open.</summary>
    /// <remarks>Validation, rendering, encoding and pre-publication cancellation leave the output untouched.
    /// Final stream copy I/O or cancellation may leave partial bytes; use a staged file with atomic replacement for file publication.
    /// Cancellation is checked around native rendering, which cannot be interrupted in progress. PDF rendering fidelity limits still apply.</remarks>
    public async Task ConvertAsync(Stream ofdInput, Stream imageOutput, int pageIndex = 0, CancellationToken cancellationToken = default)
    {
        ImageIoBudget.Streams(ofdInput, imageOutput);
        cancellationToken.ThrowIfCancellationRequested();
        if (pageIndex < 0) throw new ArgumentOutOfRangeException(nameof(pageIndex));
        var package = await new OfdReader().ReadAsync(ofdInput, _options.PackageLoadOptions, cancellationToken).ConfigureAwait(false);
        var pages = package.Pages.OrderBy(page => page.Index).ToList();
        if (pageIndex >= pages.Count) throw new ArgumentOutOfRangeException(nameof(pageIndex), "Page index is outside the OFD document.");
        var page = pages[pageIndex];
        ImageIoBudget.Page(page.WidthMillimeters, page.HeightMillimeters);
        var width = ImageIoBudget.Dimension(page.WidthMillimeters, _options.PixelsPerMillimeter);
        var height = ImageIoBudget.Dimension(page.HeightMillimeters, _options.PixelsPerMillimeter);
        ImageIoBudget.Pixels(width, height, _options.MaxPixels, _options.MaxRasterWorkingBytes);
        // PDFium truncates tiny axes to zero. Increase native sampling only enough to guarantee one pixel per axis,
        // then normalize to the requested ceil grid. The additional sampling must also fit the budgets.
        var renderPpm = Math.Max(_options.PixelsPerMillimeter,
            1.000001d / Math.Min(page.WidthMillimeters, page.HeightMillimeters));
        var renderWidth = ImageIoBudget.Dimension(page.WidthMillimeters, renderPpm);
        var renderHeight = ImageIoBudget.Dimension(page.HeightMillimeters, renderPpm);
        ImageIoBudget.Pixels(renderWidth, renderHeight, _options.MaxPixels, _options.MaxRasterWorkingBytes);
        using var pdf = new ImageIoStagingStream(_options.MaxIntermediatePdfBytes);
        await new OfdToPdfConverter(new OfdToPdfOptions
        {
            PackageLoadOptions = _options.PackageLoadOptions,
            MaxDecodedImagePixels = Math.Min(_options.MaxPixels, _options.MaxRasterWorkingBytes / 16)
        }).ConvertPackageAsync(package, pdf, new[] { pageIndex }, cancellationToken, strictAppearanceBudgets: true).ConfigureAwait(false);
        await pdf.CloseWriterAsync(cancellationToken).ConfigureAwait(false);
        // Use pixels-per-point directly: unlike the PDF import rasterizer there is no DPI clamp.
        using var reader = DocLib.Instance.GetDocReader(pdf.PathOnDisk, new PageDimensions(renderPpm * 25.4d / 72d));
        cancellationToken.ThrowIfCancellationRequested();
        using var raster = reader.GetPageReader(0);
        var actualWidth = raster.GetPageWidth();
        var actualHeight = raster.GetPageHeight();
        ImageIoBudget.Pixels(actualWidth, actualHeight, _options.MaxPixels, _options.MaxRasterWorkingBytes);
        if (Math.Abs(actualWidth - renderWidth) > 1 || Math.Abs(actualHeight - renderHeight) > 1)
            throw new InvalidDataException("PDF raster dimensions differ from the requested OFD page geometry.");
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
    }

}
