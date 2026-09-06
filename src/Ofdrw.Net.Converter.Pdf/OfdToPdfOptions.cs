using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Resource budgets for reading and rendering OFD documents.</summary>
public sealed class OfdToPdfOptions
{
    public OfdPackageLoadOptions PackageLoadOptions { get; set; } = new();
    public long MaxDecodedImagePixels { get; set; } = 40_000_000;
}
