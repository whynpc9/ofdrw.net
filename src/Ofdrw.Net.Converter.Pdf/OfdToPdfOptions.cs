using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>Resource budgets for reading and rendering OFD documents.</summary>
public sealed class OfdToPdfOptions
{
    /// <summary>CGTransform position interpretation; defaults to the historical global profile.</summary>
    public OfdGlyphMappingStrategy GlyphMappingStrategy { get; set; } = OfdGlyphMappingStrategy.LegacyConcatenatedGlobal;
    public OfdPackageLoadOptions PackageLoadOptions { get; set; } = new();
    public long MaxDecodedImagePixels { get; set; } = 40_000_000;
    /// <summary>Optional strict, conversion-local bindings for selected name-only OFD fonts.</summary>
    public OfdNameOnlyFontBindings? NameOnlyFontBindings { get; set; }
}
