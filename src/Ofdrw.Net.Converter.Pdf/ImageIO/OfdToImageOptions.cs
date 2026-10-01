using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Resolution, encoding and budgets for one OFD page export.</summary>
public sealed class OfdToImageOptions
{
    /// <summary>Pixels per millimeter; default 144 DPI. Must be finite and positive.</summary>
    public double PixelsPerMillimeter { get; set; } = 144d / 25.4d;
    /// <summary>Output encoding, independent of the destination filename.</summary>
    public OfdImageFormat Format { get; set; } = OfdImageFormat.Png;
    /// <summary>JPEG quality from 1 to 100, default 90.</summary>
    public int JpegQuality { get; set; } = 90;
    /// <summary>Archive entry, compressed input and expansion budgets; copied at construction.</summary>
    public OfdPackageLoadOptions PackageLoadOptions { get; set; } = new();
    /// <summary>Maximum pixels in the output and each decoded embedded image.</summary>
    public long MaxPixels { get; set; } = 40_000_000;
    /// <summary>Maximum estimated raster buffers (16 bytes per pixel); excludes package/PDF state and native overhead.</summary>
    public long MaxRasterWorkingBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum intermediate single-page PDF bytes.</summary>
    public long MaxIntermediatePdfBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum encoded image bytes, enforced while encoding.</summary>
    public long MaxOutputBytes { get; set; } = 64L * 1024 * 1024;
}
