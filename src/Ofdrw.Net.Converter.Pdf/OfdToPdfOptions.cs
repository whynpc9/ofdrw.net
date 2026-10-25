using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Resource budgets for reading and rendering OFD documents.</summary>
public sealed class OfdToPdfOptions
{
    public OfdPackageLoadOptions PackageLoadOptions { get; set; } = new();
    public long MaxDecodedImagePixels { get; set; } = 40_000_000;
    /// <summary>Optional strict, conversion-local bindings for selected name-only OFD fonts.</summary>
    public OfdNameOnlyFontBindings? NameOnlyFontBindings { get; set; }
}
