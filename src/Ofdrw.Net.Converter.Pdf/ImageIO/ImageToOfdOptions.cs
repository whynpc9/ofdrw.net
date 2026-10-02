using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Placement and budgets for PNG/JPEG inputs, one image per OFD page.</summary>
public sealed class ImageToOfdOptions
{
    /// <summary>Pixels per millimeter for natural image size; ignores image DPI/EXIF orientation. Must be finite and positive.</summary>
    public double PixelsPerMillimeter { get; set; } = 144d / 25.4d;
    /// <summary>Optional fixed page size in millimeters. Null uses each image's natural size. Images shrink to fit and never enlarge.</summary>
    public OfdPageSize? PageSize { get; set; }
    /// <summary>Maximum number of input images/pages.</summary>
    public int MaxPageCount { get; set; } = 1_000;
    /// <summary>Maximum encoded bytes for each image.</summary>
    public long MaxInputBytesPerImage { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum total encoded input bytes retained in the package.</summary>
    public long MaxTotalInputBytes { get; set; } = 128L * 1024 * 1024;
    /// <summary>Maximum decoded pixels per image; checked before full decode.</summary>
    public long MaxPixelsPerImage { get; set; } = 40_000_000;
    /// <summary>Maximum estimated decode buffers (16 bytes per pixel), excluding retained input, package entries and native overhead.</summary>
    public long MaxRasterWorkingBytes { get; set; } = 256L * 1024 * 1024;
    /// <summary>Maximum OFD ZIP entries: four fixed XML files, one per page, and one per distinct encoded image resource.</summary>
    public int MaxEntryCount { get; set; } = 10_000;
    /// <summary>Maximum encoded OFD package bytes, enforced while writing.</summary>
    public long MaxOutputBytes { get; set; } = 256L * 1024 * 1024;
}
