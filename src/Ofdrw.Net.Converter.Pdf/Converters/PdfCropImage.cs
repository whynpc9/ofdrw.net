using System;
using System.IO;
using System.Threading;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Converters;

// A PDF-only edge-clamped raster. The OFD keeps the caller's original bytes
// and source-to-destination CTM; this copy prevents readers from filtering
// neighboring source pixels across the destination clip.
internal static class PdfCropImage
{
    internal static (byte[] Data, double U, double V, double Width, double Height) Create(
        byte[] encoded, (double X, double Y, double Width, double Height) crop,
        int alpha, long maximumPixels, CancellationToken token, Func<byte[], Image<Rgba32>>? decode)
    {
        using var source = decode is null ? Image.Load<Rgba32>(encoded) : decode(encoded);
        if (crop.X < 0 || crop.Y < 0 || crop.Width <= 0 || crop.Height <= 0 ||
            crop.X + crop.Width > source.Width || crop.Y + crop.Height > source.Height)
            throw new InvalidDataException("PDF source crop lies outside the image pixel axes.");
        var area = crop.Width * crop.Height;
        var factor = Math.Max(1, Math.Min(16, (int)Math.Floor(Math.Sqrt(4_000_000d / area))));
        int width, height;
        long outputPixels;
        do
        {
            width = checked((int)Math.Ceiling(crop.Width * factor));
            height = checked((int)Math.Ceiling(crop.Height * factor));
            outputPixels = checked(((long)width + 2) * ((long)height + 2));
            if (outputPixels <= maximumPixels) break;
            factor--;
        } while (factor >= 1);
        if (outputPixels > maximumPixels)
            throw new InvalidDataException("PDF source crop exceeds the raster pixel budget after edge padding.");
        var stepX = crop.Width / width;
        var stepY = crop.Height / height;
        using var pixels = new Image<Rgba32>(width + 2, height + 2);
        for (var row = -1; row <= height; row++)
        {
            token.ThrowIfCancellationRequested();
            var sourceY = Math.Max(0, Math.Min(source.Height - 1,
                (int)Math.Floor(crop.Y + (Math.Max(0, Math.Min(height - 1, row)) + 0.5) * stepY)));
            for (var column = -1; column <= width; column++)
            {
                var sourceX = Math.Max(0, Math.Min(source.Width - 1,
                    (int)Math.Floor(crop.X + (Math.Max(0, Math.Min(width - 1, column)) + 0.5) * stepX)));
                var color = source[sourceX, sourceY];
                if (alpha < 255) color.A = (byte)Math.Round(color.A * alpha / 255d);
                pixels[column + 1, row + 1] = color;
            }
        }
        using var output = new MemoryStream();
        pixels.Save(output, new PngEncoder { ColorType = PngColorType.RgbWithAlpha });
        return (output.ToArray(), (crop.X - stepX) / source.Width, (crop.Y - stepY) / source.Height,
            (crop.Width + 2 * stepX) / source.Width, (crop.Height + 2 * stepY) / source.Height);
    }
}
