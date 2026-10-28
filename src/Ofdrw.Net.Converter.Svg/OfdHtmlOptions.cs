using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Svg;

/// <summary>Finite limits for static HTML previews. Reader/model limits remain independent.</summary>
public sealed class OfdHtmlOptions
{
    /// <summary>CGTransform position interpretation; defaults to the historical global profile.</summary>
    public OfdGlyphMappingStrategy GlyphMappingStrategy { get; set; } = OfdGlyphMappingStrategy.LegacyConcatenatedGlobal;
    /// <summary>Plain text document title; never interpreted as markup.</summary>
    public string Title { get; set; } = "OFD preview";
    /// <summary>Maximum title length in UTF-16 code units.</summary>
    public int MaxTitleCharacters { get; set; } = 1024;
    /// <summary>Maximum output page occurrences, including duplicates.</summary>
    public int MaxPages { get; set; } = 200;
    /// <summary>Maximum complete UTF-8 HTML size, without a BOM.</summary>
    public long MaxOutputBytes { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum UTF-8 section size, including its resources and label.</summary>
    public long MaxPageBytes { get; set; } = 16L * 1024 * 1024;
    /// <summary>Maximum conservative reservation of generated elements, attributes and text nodes.</summary>
    public long MaxGeneratedNodes { get; set; } = 100_000;
    /// <summary>Maximum conservative UTF-16 generation capacity, including intermediate copies and escaping.</summary>
    public long MaxGeneratedCharacters { get; set; } = 32L * 1024 * 1024;
    /// <summary>Maximum cumulative decoded image pixels; every image occurrence is charged before PNG inflation or JPEG entropy traversal.</summary>
    public long MaxDecodedImagePixels { get; set; } = 100_000_000;
    /// <summary>Maximum sum of original embedded image/font bytes, charged again per occurrence.</summary>
    public long MaxEmbeddedResourceBytes { get; set; } = 32L * 1024 * 1024;
}
