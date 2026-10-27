using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// First CGTransform render profile: one source TextCode, independent BMP
// characters and a complete one-to-one mapping to original font GIDs. Keep
// the source XML as the owner of this transient projection.
internal sealed class OfdExplicitGlyphRun
{
    private OfdExplicitGlyphRun(OfdTextRun run, int[] glyphIds)
    { Run = run; GlyphIds = glyphIds; }

    internal OfdTextRun Run { get; }
    internal IReadOnlyList<int> GlyphIds { get; }

    internal static OfdExplicitGlyphRun? Read(OfdTextElement text)
    {
        if (string.IsNullOrWhiteSpace(text.SourceXml)) return null;
        XElement root;
        try { root = OfdXmlEnvelope.Parse(text.SourceXml!); }
        catch (System.Xml.XmlException error)
        { throw new NotSupportedException("Malformed CGTransform source XML cannot be rendered.", error); }
        var transforms = root.DescendantsAndSelf().Where(node => node.Name.LocalName == "CGTransform").ToArray();
        if (transforms.Length == 0) return null;
        if (root.Name.LocalName != "TextObject" || transforms.Any(node => node.Parent != root) ||
            text.Runs.Count != 1 || root.Elements(root.Name.Namespace + "TextCode").Count() != 1)
            throw new NotSupportedException("CGTransform rendering requires one source TextObject and one TextCode.");
        var code = root.Element(root.Name.Namespace + "TextCode")!;
        if (code.HasElements || code.Nodes().Any(node => node is not XText) ||
            OfdTextRenderProjection.Decode(code.Value, CancellationToken.None) != text.Runs[0].Text ||
            text.Text != text.Runs[0].Text ||
            code.Attribute("DeltaX")?.Value != text.Runs[0].DeltaX ||
            code.Attribute("DeltaY")?.Value != text.Runs[0].DeltaY ||
            !SameAxis(code.Attribute("X"), text.Runs[0].XMillimeters) ||
            !SameAxis(code.Attribute("Y"), text.Runs[0].YMillimeters))
            throw new NotSupportedException("CGTransform TextCode must match the unedited source run.");
        var logical = text.Runs[0].Text;
        if (logical.Length is < 1 or > 1024 || logical.Any(value => char.IsSurrogate(value) ||
                char.IsControl(value) || char.IsWhiteSpace(value) ||
                CharUnicodeInfo.GetUnicodeCategory(value) is UnicodeCategory.Format or
                    UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark))
            throw new NotSupportedException("CGTransform one-to-one rendering requires independent, nonblank BMP characters.");
        var glyphIds = new int[logical.Length];
        var assigned = new bool[logical.Length];
        var unicodeByGlyph = new Dictionary<int, char>();
        foreach (var transform in transforms)
        {
            if (transform.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                    attribute.Name.LocalName is not ("CodePosition" or "CodeCount" or "GlyphCount")))
                throw new NotSupportedException("CGTransform has unsupported attributes.");
            var position = Number(transform.Attribute("CodePosition"), -1);
            var codeCount = Number(transform.Attribute("CodeCount"), 1);
            var glyphCount = Number(transform.Attribute("GlyphCount"), 1);
            if (position < 0 || codeCount <= 0 || glyphCount != codeCount ||
                position > logical.Length - codeCount ||
                transform.Elements().Count() != 1 ||
                transform.Element(root.Name.Namespace + "Glyphs") is not { } glyphs || glyphs.HasElements ||
                glyphs.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                glyphs.Nodes().Any(node => node is not XText))
                throw new NotSupportedException("CGTransform span or glyph count is outside the one-to-one profile.");
            var tokens = glyphs.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length != glyphCount) throw new NotSupportedException("CGTransform Glyphs count does not match GlyphCount.");
            for (var index = 0; index < glyphCount; index++)
            {
                var slot = position + index;
                if (assigned[slot] || !int.TryParse(tokens[index], NumberStyles.None,
                        CultureInfo.InvariantCulture, out var gid) || gid <= 0 || gid > ushort.MaxValue)
                    throw new NotSupportedException("CGTransform contains overlapping or invalid original GIDs.");
                if (unicodeByGlyph.TryGetValue(gid, out var prior) && prior != logical[slot])
                    throw new NotSupportedException("One original GID has ambiguous PDF Unicode semantics.");
                unicodeByGlyph[gid] = logical[slot];
                glyphIds[slot] = gid;
                assigned[slot] = true;
            }
        }
        if (assigned.Any(value => !value))
            throw new NotSupportedException("CGTransform must map every TextCode character before rendering.");
        if (!string.IsNullOrWhiteSpace(text.Runs[0].DeltaY) &&
            string.IsNullOrWhiteSpace(text.Runs[0].DeltaX))
            throw new NotSupportedException("CGTransform DeltaY without DeltaX is outside the render profile.");
        CountDeltas(text.Runs[0].DeltaX, logical.Length - 1);
        CountDeltas(text.Runs[0].DeltaY, logical.Length - 1);
        return new OfdExplicitGlyphRun(text.Runs[0], glyphIds);
    }

    private static void CountDeltas(string? expression, int expected)
    {
        if (string.IsNullOrWhiteSpace(expression)) return;
        var tokens = expression!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var count = 0;
        for (var index = 0; index < tokens.Length;)
        {
            var repeat = 1;
            if (string.Equals(tokens[index], "g", StringComparison.OrdinalIgnoreCase))
            {
                if (++index >= tokens.Length || !int.TryParse(tokens[index++], NumberStyles.None,
                        CultureInfo.InvariantCulture, out repeat) || repeat < 0)
                    throw new NotSupportedException("Invalid CGTransform repeated positioning delta.");
            }
            if (index >= tokens.Length || !double.TryParse(tokens[index++], NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var value) || double.IsNaN(value) || double.IsInfinity(value) ||
                repeat > expected - count)
                throw new NotSupportedException("Invalid CGTransform positioning delta.");
            count += repeat;
        }
        if (count != expected)
            throw new NotSupportedException("CGTransform requires complete glyph-slot positioning deltas.");
    }

    private static int Number(XAttribute? attribute, int fallback) => attribute is null ? fallback :
        int.TryParse(attribute.Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value) ? value : -1;

    private static bool SameAxis(XAttribute? attribute, double actual) => attribute is not null &&
        double.TryParse(attribute.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var value) &&
        !double.IsNaN(value) && !double.IsInfinity(value) && value == actual;
}
