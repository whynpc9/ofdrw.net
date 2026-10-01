using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;

namespace Ofdrw.Net.Converter.Pdf.Converters;

/// <summary>Imports PNG/JPEG encoded pixels into OFD, in input order, one centered image per page.</summary>
public sealed class ImageToOfdConverter
{
    private readonly ImageToOfdOptions _options;
    /// <summary>Creates a natural-size importer at 144 DPI with default budgets.</summary>
    public ImageToOfdConverter() : this(new ImageToOfdOptions()) { }
    /// <summary>Copies and validates placement and resource budgets.</summary>
    public ImageToOfdConverter(ImageToOfdOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        ImageIoBudget.Positive(options.PixelsPerMillimeter, nameof(options.PixelsPerMillimeter));
        if (options.MaxPageCount <= 0 || options.MaxEntryCount <= 0 || options.MaxInputBytesPerImage <= 0 ||
            options.MaxInputBytesPerImage > int.MaxValue || options.MaxTotalInputBytes <= 0 ||
            options.MaxPixelsPerImage <= 0 || options.MaxRasterWorkingBytes <= 0 || options.MaxOutputBytes <= 0)
            throw new ArgumentException("Positive image import budgets are required; each encoded image must fit a managed byte array.", nameof(options));
        if (options.PageSize is not null) ImageIoBudget.ImportExtent(options.PageSize.WidthMillimeters, options.PageSize.HeightMillimeters);
        _options = new ImageToOfdOptions
        {
            PixelsPerMillimeter = options.PixelsPerMillimeter,
            PageSize = options.PageSize is null ? null : new OfdPageSize
            { WidthMillimeters = options.PageSize.WidthMillimeters, HeightMillimeters = options.PageSize.HeightMillimeters },
            MaxPageCount = options.MaxPageCount, MaxEntryCount = options.MaxEntryCount,
            MaxInputBytesPerImage = options.MaxInputBytesPerImage, MaxTotalInputBytes = options.MaxTotalInputBytes,
            MaxPixelsPerImage = options.MaxPixelsPerImage, MaxRasterWorkingBytes = options.MaxRasterWorkingBytes,
            MaxOutputBytes = options.MaxOutputBytes
        };
    }

    /// <summary>Imports one PNG/JPEG. Leaves caller streams open.</summary>
    public Task ConvertAsync(Stream imageInput, Stream ofdOutput, CancellationToken cancellationToken = default)
        => ConvertAsync(new[] { imageInput }, ofdOutput, cancellationToken);

    /// <summary>Imports one page per input in order, preserving original encoded image data.</summary>
    /// <remarks>Embedded DPI and EXIF orientation are ignored; placement uses the raw encoded pixel axes.
    /// Each image is fully decoded after checking dimensions. Animated/multiple-frame images are rejected.
    /// Validation, encoding and pre-publication cancellation leave the output untouched. Final stream copy I/O or cancellation may leave partial bytes;
    /// use a staged file with atomic replacement for file publication. Input streams remain open.</remarks>
    public async Task ConvertAsync(IReadOnlyList<Stream> imageInputs, Stream ofdOutput, CancellationToken cancellationToken = default)
    {
        if (imageInputs is null) throw new ArgumentNullException(nameof(imageInputs));
        if (ofdOutput is null) throw new ArgumentNullException(nameof(ofdOutput));
        cancellationToken.ThrowIfCancellationRequested();
        if (imageInputs.Count == 0 || imageInputs.Count > _options.MaxPageCount ||
            5L + imageInputs.Count * 2L > _options.MaxEntryCount)
            throw new ArgumentException("Input image count exceeds the page/entry budget or is empty.", nameof(imageInputs));
        // Validate every stream before consuming any input.
        foreach (var input in imageInputs) ImageIoBudget.Streams(input, ofdOutput);
        var package = new OfdDocumentPackage();
        long total = 0;
        for (var index = 0; index < imageInputs.Count; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var available = Math.Min(_options.MaxInputBytesPerImage, _options.MaxTotalInputBytes - total);
            if (available <= 0) throw new InvalidDataException("Total encoded image input budget exceeded.");
            using var staged = new MemoryStream();
            total += await BoundedStreamCopy.CopyAsync(imageInputs[index], staged, available, "Encoded image", cancellationToken).ConfigureAwait(false);
            staged.Position = 0;
            var identified = await Image.IdentifyWithFormatAsync(staged, cancellationToken).ConfigureAwait(false);
            var info = identified.ImageInfo;
            var format = identified.Format?.Name;
            if (info is null || (format != "PNG" && format != "JPEG"))
                throw new InvalidDataException("Only PNG and JPEG encoded images are supported.");
            ImageIoBudget.Pixels(info.Width, info.Height, _options.MaxPixelsPerImage, _options.MaxRasterWorkingBytes);
            if (format == "PNG") ValidatePngChunks(staged.GetBuffer(), (int)staged.Length, cancellationToken);
            staged.Position = 0;
            // Identification alone does not validate the payload. Decode completely, then promptly release the pixel buffer.
            using (var decoded = await Image.LoadAsync(staged, cancellationToken).ConfigureAwait(false))
            {
                if (decoded.Frames.Count != 1) throw new InvalidDataException("Only single-frame PNG/JPEG images are supported.");
            }
            double pageWidth, pageHeight, imageWidth, imageHeight;
            if (_options.PageSize is null)
            {
                pageWidth = imageWidth = info.Width / _options.PixelsPerMillimeter;
                pageHeight = imageHeight = info.Height / _options.PixelsPerMillimeter;
                ImageIoBudget.ImportExtent(pageWidth, pageHeight);
            }
            else
            {
                pageWidth = _options.PageSize.WidthMillimeters;
                pageHeight = _options.PageSize.HeightMillimeters;
                // Fit in pixel space: even when 1/ppm overflows, the finite fixed-page factors keep placement bounded.
                var millimetersPerPixel = Math.Min(1d / _options.PixelsPerMillimeter,
                    Math.Min(pageWidth / info.Width, pageHeight / info.Height));
                // Multiplication can overshoot the limiting axis by an ULP (e.g. 10000/145*145).
                imageWidth = Math.Min(pageWidth, info.Width * millimetersPerPixel);
                imageHeight = Math.Min(pageHeight, info.Height * millimetersPerPixel);
                ImageIoBudget.ImportExtent(imageWidth, imageHeight);
            }
            var page = new OfdPage { Index = index, WidthMillimeters = pageWidth, HeightMillimeters = pageHeight };
            page.Elements.Add(new OfdImageElement
            {
                FileName = $"image-{index + 1}.{(format == "PNG" ? "png" : "jpg")}",
                MediaType = format == "PNG" ? "image/png" : "image/jpeg", Data = staged.ToArray(),
                XMillimeters = (pageWidth - imageWidth) / 2d, YMillimeters = (pageHeight - imageHeight) / 2d,
                WidthMillimeters = imageWidth, HeightMillimeters = imageHeight
            });
            package.Pages.Add(page);
        }
        using var encoded = new ImageIoStagingStream(_options.MaxOutputBytes);
        await new OfdPackageWriter().WriteAsync(package, encoded, cancellationToken).ConfigureAwait(false);
        await encoded.PublishAsync(ofdOutput, cancellationToken).ConfigureAwait(false);
    }

    private static void ValidatePngChunks(byte[] data, int length, CancellationToken token)
    {
        // ImageSharp 2.x ignores APNG animation chunks and would report one decoded frame.
        // Inspect the bounded encoded container before decoding or preserving it in the OFD.
        var offset = 8;
        while (offset <= length - 12)
        {
            token.ThrowIfCancellationRequested();
            var chunkLength = ((long)data[offset] << 24) | ((long)data[offset + 1] << 16) |
                ((long)data[offset + 2] << 8) | data[offset + 3];
            if (chunkLength > length - offset - 12) throw new InvalidDataException("Truncated PNG chunk.");
            var kind = System.Text.Encoding.ASCII.GetString(data, offset + 4, 4);
            if (kind is "acTL" or "fcTL" or "fdAT") throw new InvalidDataException("Animated PNG inputs are not supported.");
            offset += (int)chunkLength + 12;
            if (kind == "IEND")
            {
                if (chunkLength != 0 || offset != length) throw new InvalidDataException("Invalid PNG end chunk or trailing bytes.");
                return;
            }
        }
        throw new InvalidDataException("PNG end chunk is missing.");
    }
}
