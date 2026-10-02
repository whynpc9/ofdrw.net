using System.IO;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
namespace Ofdrw.Net.Core.Fonts;
internal static class EmbeddedFontCoverage
{
    internal static bool HasExplicitGlyphReferences(OfdTextElement text) =>
        !string.IsNullOrWhiteSpace(text.SourceXml) && XElement.Parse(text.SourceXml!).Descendants()
            .Any(node => node.Name.LocalName == "CGTransform");

    internal static void Validate(string text, OpenTypeCmap cmap, string name)
    {
        var previous = -1;
        foreach (var scalar in OpenTypeFace.Scalars(text))
        {
            if (OpenTypeCmap.IsVariationSelector(scalar))
            {
                if (previous < 0 || !cmap.SupportsVariation(previous, scalar))
                    throw new InvalidDataException($"Font '{name}' lacks variation sequence U+{previous:X}/U+{scalar:X}.");
            }
            else if (scalar != '\r' && scalar != '\n' && scalar != '\t' && cmap.Glyph(scalar) == 0)
                throw new InvalidDataException($"Font '{name}' lacks U+{scalar:X4}; bind a configured fallback font before writing.");
            previous = scalar;
        }
    }
}
