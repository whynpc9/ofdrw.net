using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using JBig2Decoder.NETStandard;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Reader.Readers;

/// <summary>Creates a temporary render payload for the restricted JBIG2 and ImageMask profile.</summary>
/// <remarks>The source media and SourceXml remain authoritative and are never rewritten.</remarks>
public static class OfdImageRenderProjection
{
    private const long MaxMaskedPixels = 16_000_000;

    /// <summary>Returns the original bytes for ordinary images and a PNG for admitted JBIG2 or masked images.</summary>
    public static (byte[] Data, string MediaType) Prepare(OfdImageElement image, long maximumPixels,
        CancellationToken cancellationToken = default)
    {
        if (image is null) throw new ArgumentNullException(nameof(image));
        cancellationToken.ThrowIfCancellationRequested();
        var source = image.SourceXml is null ? null : OfdXmlEnvelope.Parse(image.SourceXml);
        var sourceMask = source?.Attribute("ImageMask")?.Value;
        if (sourceMask != image.MaskResourceId)
            throw new NotSupportedException("OFD ImageMask reference changed after resource resolution.");
        // Preserve the established admission and exception contracts of ordinary
        // image exports. The stricter XML/resource preflight belongs to this
        // restricted JBIG2 and ImageMask projection only.
        if (image.MaskResourceId is null && image.MediaType != "image/jbig2")
            return (image.Data, image.MediaType);
        if (source is not null && (!OfdGraphicXmlContract.HasKnownChildren(source) ||
                OfdGraphicXmlContract.HasUnsupportedReferences(source, allowImageMask: sourceMask is not null) ||
                source.DescendantsAndSelf().Attributes().Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute))))
            throw new NotSupportedException("OFD image has unsupported drawing attributes or references.");
        if (image.SourceMediaAmbiguous || image.Data is not { Length: > 0 })
            throw new NotSupportedException("OFD image resource is unresolved or ambiguous.");
        if (source is not null && EncodedImageAdmission.IdentifyFormat(image.Data) != image.MediaType)
            throw new NotSupportedException("OFD image media declaration differs from its encoded payload.");
        var limit = Math.Min(maximumPixels, image.MaskResourceId is null ? RestrictedJbig2Admission.MaxPixels : MaxMaskedPixels);
        var baseBytes = image.MediaType == "image/jbig2" ? DecodeJbig2(image.Data, limit, cancellationToken) : image.Data;
        if (image.MaskResourceId is null) return (baseBytes, "image/png");
        if (image.MaskSourceAmbiguous || image.MaskData is not { Length: > 0 } maskBytes ||
            image.MaskMediaType is not ("image/png" or "image/jbig2"))
            throw new NotSupportedException("OFD ImageMask resource is unresolved or outside the binary image profile.");
        if (EncodedImageAdmission.IdentifyFormat(maskBytes) != image.MaskMediaType)
            throw new NotSupportedException("OFD ImageMask media declaration differs from its encoded payload.");
        if (image.MaskMediaType == "image/png")
        {
            if (maskBytes.Length < 33 || maskBytes[24] != 1 || maskBytes[25] != 0)
                throw new NotSupportedException("OFD ImageMask must use one-bit grayscale PNG or restricted JBIG2.");
            EncodedImageAdmission.Validate(maskBytes, "image/png", cancellationToken);
        }
        else maskBytes = DecodeJbig2(maskBytes, limit, cancellationToken);

        var information = Image.Identify(baseBytes);
        if (information is null || information.Width <= 0 || information.Height <= 0 ||
            (long)information.Width * information.Height > limit)
            throw new InvalidDataException("OFD masked image exceeds the decoded pixel limit.");
        var maskInformation = Image.Identify(maskBytes);
        if (maskInformation is null || maskInformation.Width != information.Width ||
            maskInformation.Height != information.Height)
            throw new NotSupportedException("OFD ImageMask dimensions differ from the image.");
        using var baseImage = Image.Load<Rgba32>(baseBytes);
        using var mask = Image.Load<Rgba32>(maskBytes);
        if (baseImage.Width != information.Width || baseImage.Height != information.Height ||
            mask.Width != information.Width || mask.Height != information.Height)
            throw new NotSupportedException("OFD ImageMask dimensions differ from the image.");
        for (var y = 0; y < mask.Height; y++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            for (var x = 0; x < mask.Width; x++)
            {
                var pixel = mask[x, y];
                if (pixel.A != 255 || pixel.R != pixel.G || pixel.R != pixel.B || pixel.R is not (0 or 255))
                    throw new NotSupportedException("OFD ImageMask contains a non-binary pixel.");
                var basePixel = baseImage[x, y];
                // White retains the source pixel; black clears it.
                basePixel.A = pixel.R == 255 ? basePixel.A : (byte)0;
                baseImage[x, y] = basePixel;
            }
        }
        using var output = new MemoryStream();
        baseImage.Save(output, new PngEncoder { ColorType = PngColorType.RgbWithAlpha });
        return (output.ToArray(), "image/png");
    }

    private static byte[] DecodeJbig2(byte[] encoded, long maximumPixels, CancellationToken token)
    {
        var admitted = RestrictedJbig2Admission.Validate(encoded, token);
        if ((long)admitted.Width * admitted.Height > maximumPixels)
            throw new InvalidDataException("JBIG2 image exceeds the decoded pixel limit.");
        token.ThrowIfCancellationRequested();
        byte[] rgb;
        int width, height;
        try { rgb = new JBIG2StreamDecoder().DecodeJBIG2(encoded, out width, out height); }
        catch (Exception error) when (error is not OutOfMemoryException && error is not OperationCanceledException)
        { throw new InvalidDataException("JBIG2 decoding failed.", error); }
        token.ThrowIfCancellationRequested();
        if (width != admitted.Width || height != admitted.Height || rgb.Length != checked(width * height * 3))
            throw new InvalidDataException("JBIG2 decoder dimensions or pixel count disagree with the admitted page.");
        using var pixels = Image.LoadPixelData<Rgb24>(rgb, width, height);
        using var output = new MemoryStream();
        pixels.SaveAsPng(output);
        return output.ToArray();
    }
}
