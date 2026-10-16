using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
namespace Ofdrw.Net.Core.Fonts;
internal static class EmbeddedFontCoverage
{
    internal static bool HasExplicitGlyphReferences(OfdTextElement text) => HasExplicitGlyphReferences(text.SourceXml);
    internal static bool HasExplicitGlyphReferences(string? sourceXml)
    {
        if (string.IsNullOrWhiteSpace(sourceXml)) return false;
        try { return XElement.Parse(sourceXml!).DescendantsAndSelf().Any(node => node.Name.LocalName == "CGTransform"); }
        catch (System.Xml.XmlException exception)
        {
            throw new System.NotSupportedException("Malformed SourceXml cannot safely preserve glyph references.", exception);
        }
    }

    internal static void Validate(OfdTextElement text, OpenTypeCmap cmap, string name, CancellationToken cancellationToken = default)
        => Validate(text, cmap, name, cancellationToken, null);

    // Per-operation test seam: observe completed scalar validation without a
    // global hook, asynchronous timing, or changing the existing entry point.
    internal static void Validate(OfdTextElement text, OpenTypeCmap cmap, string name,
        CancellationToken cancellationToken, System.Action<int>? scalarValidated)
    {
        if (text.Runs.Count == 0)
        {
            Validate(text.Text, cmap, name, cancellationToken, OfdTextGeometry.IsWhiteSpace(text.Text, cancellationToken), scalarValidated);
            return;
        }
        foreach (var run in text.Runs)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var deltas = OfdTextGeometry.PositioningDeltas(run, cancellationToken);
            var positioned = deltas.X.Count > 0 || deltas.Y.Count > 0;
            // Only independently positioned blank glyphs may lack an outline.
            // A mixed unpositioned run must resolve a real covering font.
            if (positioned)
                foreach (var glyph in OfdTextGeometry.EnumerateGlyphs(run.Text, cancellationToken))
                    Validate(glyph, cmap, name, cancellationToken, OfdTextGeometry.IsWhiteSpace(glyph, cancellationToken), scalarValidated);
            else Validate(run.Text, cmap, name, cancellationToken, OfdTextGeometry.IsWhiteSpace(run.Text, cancellationToken), scalarValidated);
        }
    }

    internal static void Validate(string text, OpenTypeCmap cmap, string name, CancellationToken cancellationToken = default)
        => Validate(text, cmap, name, cancellationToken, false);

    private static void Validate(string text, OpenTypeCmap cmap, string name, CancellationToken cancellationToken, bool blankGlyph)
        => Validate(text, cmap, name, cancellationToken, blankGlyph, null);

    private static void Validate(string text, OpenTypeCmap cmap, string name, CancellationToken cancellationToken,
        bool blankGlyph, System.Action<int>? scalarValidated)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var previous = -1;
        foreach (var scalar in OpenTypeFace.Scalars(text))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (OpenTypeCmap.IsVariationSelector(scalar))
            {
                if (previous < 0 || !cmap.SupportsVariation(previous, scalar))
                    throw new InvalidDataException($"Font '{name}' lacks variation sequence U+{previous:X}/U+{scalar:X}.");
            }
            else if (scalar != 9 && scalar != 10 && scalar != 13 && !(blankGlyph && scalar <= 0xFFFF && char.IsWhiteSpace((char)scalar)) && !UnicodeFontSubsetProfile.IsNonRenderingControl(scalar) && cmap.Glyph(scalar) == 0)
                throw new InvalidDataException($"Font '{name}' lacks U+{scalar:X4}; bind a configured fallback font before writing.");
            previous = scalar;
            scalarValidated?.Invoke(scalar);
        }
    }
}
