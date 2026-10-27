using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Core.Models;

// A bounded source-character-span to original-GID projection. Source XML owns
// the mapping; this transient view never changes the OFD's logical text.
internal sealed class OfdExplicitGlyphSpan
{
    internal OfdExplicitGlyphSpan(int codePosition, int codeCount, int glyphPosition,
        int glyphCount, string logicalText)
    {
        CodePosition = codePosition;
        CodeCount = codeCount;
        GlyphPosition = glyphPosition;
        GlyphCount = glyphCount;
        LogicalText = logicalText;
    }

    internal int CodePosition { get; }
    internal int CodeCount { get; }
    internal int GlyphPosition { get; }
    internal int GlyphCount { get; }
    internal string LogicalText { get; }
}

internal sealed class OfdExplicitGlyphRun
{
    private OfdExplicitGlyphRun(IReadOnlyList<OfdTextRun> runs, int[] glyphIds,
        IReadOnlyList<OfdExplicitGlyphSpan> spans)
    { Runs = runs; GlyphIds = glyphIds; Spans = spans; }

    internal IReadOnlyList<OfdTextRun> Runs { get; }
    internal string LogicalText => string.Concat(Runs.Select(run => run.Text));
    internal IReadOnlyList<int> GlyphIds { get; }
    internal IReadOnlyList<OfdExplicitGlyphSpan> Spans { get; }
    internal string? SingleGlyphSourceSpan => Spans.Count == 1 &&
        Spans[0].CodeCount > 1 && Spans[0].GlyphCount == 1 ? Spans[0].LogicalText : null;
    internal string? PathGlyphSourceSpan => Spans.Count == 1 &&
        Spans[0].GlyphCount > 1 && Spans[0].CodeCount != Spans[0].GlyphCount
            ? Spans[0].LogicalText : null;

    internal IReadOnlyList<(double X, double Y)> PathGlyphOrigins(CancellationToken token)
    {
        if (PathGlyphSourceSpan is null || Runs.Count != 1)
            throw new InvalidOperationException("A complete source span is required for original-GID path placement.");
        var source = Runs[0];
        var deltaX = OfdTextGeometry.ExpandDeltas(source.DeltaX, GlyphIds.Count - 1, token);
        var deltaY = OfdTextGeometry.ExpandDeltas(source.DeltaY, GlyphIds.Count - 1, token);
        var origins = new (double X, double Y)[GlyphIds.Count];
        var x = source.XMillimeters; var y = source.YMillimeters;
        for (var index = 0; index < origins.Length; index++)
        {
            token.ThrowIfCancellationRequested();
            origins[index] = (x, y);
            if (index < deltaX.Count) x += deltaX[index];
            if (index < deltaY.Count) y += deltaY[index];
            if (double.IsNaN(x) || double.IsInfinity(x) ||
                double.IsNaN(y) || double.IsInfinity(y))
                throw new NotSupportedException("CGTransform glyph-slot positions must remain finite.");
        }
        return origins;
    }

    internal static OfdExplicitGlyphRun? Read(OfdTextElement text, OfdFontResource? resource = null,
        IDictionary<OfdFontResource, OpenTypeCmap>? cmapCache = null, FontParseBudget? parseBudget = null)
    {
        if (string.IsNullOrWhiteSpace(text.SourceXml)) return null;
        XElement root;
        try { root = OfdXmlEnvelope.Parse(text.SourceXml!); }
        catch (System.Xml.XmlException error)
        { throw new NotSupportedException("Malformed CGTransform source XML cannot be rendered.", error); }
        var transforms = root.DescendantsAndSelf().Where(node => node.Name.LocalName == "CGTransform").ToArray();
        if (transforms.Length == 0) return null;
        var codes = root.Elements(root.Name.Namespace + "TextCode").ToArray();
        if (root.Name.LocalName != "TextObject" || transforms.Any(node => node.Parent != root) ||
            codes.Length == 0 || codes.Length != text.Runs.Count)
            throw new NotSupportedException("CGTransform rendering requires matching source TextCode runs.");
        double? priorX = null, priorY = null;
        for (var index = 0; index < codes.Length; index++)
        {
            var code = codes[index]; var run = text.Runs[index];
            if (code.HasElements || code.Nodes().Any(node => node is not XText) ||
                OfdTextRenderProjection.Decode(code.Value, CancellationToken.None) != run.Text ||
                code.Attribute("DeltaX")?.Value != run.DeltaX ||
                code.Attribute("DeltaY")?.Value != run.DeltaY ||
                !SameAxis(code.Attribute("X"), priorX, run.XMillimeters) ||
                !SameAxis(code.Attribute("Y"), priorY, run.YMillimeters))
                throw new NotSupportedException("CGTransform TextCode must match the unedited source runs.");
            priorX = run.XMillimeters; priorY = run.YMillimeters;
        }
        var logical = string.Concat(text.Runs.Select(run => run.Text));
        if (text.Text != logical || text.Runs.Any(run => run.Text.Length == 0))
            throw new NotSupportedException("CGTransform requires nonempty matching source runs.");
        // Admit only global, unambiguous BMP positions over the
        // concatenated runs. Per-TextCode position resets need separate evidence.
        if (logical.Length is < 1 or > 1024 || logical.Any(value => char.IsSurrogate(value) ||
                char.IsControl(value) || char.IsWhiteSpace(value) ||
                CharUnicodeInfo.GetUnicodeCategory(value) == UnicodeCategory.Format))
            throw new NotSupportedException("CGTransform glyph rendering requires independent, nonblank BMP characters.");
        // Table 46 positions the drawn glyphs. For a complete multi-glyph
        // source span, an explicit delta for every glyph gap removes any need
        // to infer advances or how source-code deltas map to glyph positions.
        // A combining mark is allowed after a base within this complete span;
        // the older independent-character profiles retain their rejection.
        if (logical.Length <= 8 && codes.Length == 1 && transforms.Length == 1 &&
            !IsCombiningMark(logical[0]) &&
            !string.IsNullOrWhiteSpace(text.Runs[0].DeltaX))
        {
            var transform = transforms[0];
            var codeCount = Number(transform.Attribute("CodeCount"), 1);
            var glyphCount = Number(transform.Attribute("GlyphCount"), 1);
            if (glyphCount is >= 2 and <= 8 && glyphCount != codeCount &&
                transform.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
                    attribute.Name.LocalName is "CodePosition" or "CodeCount" or "GlyphCount") &&
                Number(transform.Attribute("CodePosition"), -1) == 0 &&
                codeCount == logical.Length && transform.Elements().Count() == 1 &&
                transform.Element(root.Name.Namespace + "Glyphs") is { } glyphs &&
                !glyphs.HasElements &&
                glyphs.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
                glyphs.Nodes().All(node => node is XText) &&
                glyphs.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries) is { } tokens &&
                tokens.Length == glyphCount)
            {
                var ids = new int[glyphCount];
                var seen = new HashSet<int>();
                for (var index = 0; index < glyphCount; index++)
                {
                    if (!int.TryParse(tokens[index], NumberStyles.None, CultureInfo.InvariantCulture,
                            out var gid) || gid is <= 0 or > ushort.MaxValue || !seen.Add(gid))
                        throw new NotSupportedException("CGTransform contains overlapping or invalid original GIDs.");
                    ids[index] = gid;
                }
                CountDeltas(text.Runs[0].DeltaX, glyphCount - 1);
                CountDeltas(text.Runs[0].DeltaY, glyphCount - 1);
                var positioned = new OfdExplicitGlyphRun(text.Runs, ids,
                    [new OfdExplicitGlyphSpan(0, codeCount, 0, glyphCount, logical)]);
                _ = positioned.PathGlyphOrigins(CancellationToken.None);
                return positioned;
            }
        }
        if (logical.Any(IsCombiningMark))
            throw new NotSupportedException("CGTransform glyph rendering requires independent, nonblank BMP characters.");
        // A whole, unpositioned source span can map to one original GID. A
        // private PDF code is then given the exact source span in ToUnicode.
        // Mixed clusters and per-glyph deltas need a separate placement model.
        if (logical.Length is >= 2 and <= 8 && codes.Length == 1 && transforms.Length == 1 &&
            string.IsNullOrWhiteSpace(text.Runs[0].DeltaX) &&
            string.IsNullOrWhiteSpace(text.Runs[0].DeltaY))
        {
            var transform = transforms[0];
            if (transform.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
                    attribute.Name.LocalName is "CodePosition" or "CodeCount" or "GlyphCount") &&
                Number(transform.Attribute("CodePosition"), -1) == 0 &&
                Number(transform.Attribute("CodeCount"), 1) == logical.Length &&
                Number(transform.Attribute("GlyphCount"), 1) == 1 &&
                transform.Elements().Count() == 1 &&
                transform.Element(root.Name.Namespace + "Glyphs") is { } glyphs &&
                !glyphs.HasElements &&
                glyphs.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
                glyphs.Nodes().All(node => node is XText) &&
                glyphs.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries) is { Length: 1 } tokens &&
                int.TryParse(tokens[0], NumberStyles.None, CultureInfo.InvariantCulture, out var gid) &&
                gid is > 0 and <= ushort.MaxValue)
                return new OfdExplicitGlyphRun(text.Runs, [gid],
                    [new OfdExplicitGlyphSpan(0, logical.Length, 0, 1, logical)]);
        }
        // One source character and two original glyph slots can share one
        // origin. Require a complete unpositioned TextCode so no adjacent
        // cluster or character/glyph delta interpretation is inferred.
        if (logical.Length == 1 && codes.Length == 1 && transforms.Length == 1 &&
            string.IsNullOrWhiteSpace(text.Runs[0].DeltaX) &&
            string.IsNullOrWhiteSpace(text.Runs[0].DeltaY))
        {
            var transform = transforms[0];
            if (transform.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
                    attribute.Name.LocalName is "CodePosition" or "CodeCount" or "GlyphCount") &&
                Number(transform.Attribute("CodePosition"), -1) == 0 &&
                Number(transform.Attribute("CodeCount"), 1) == 1 &&
                Number(transform.Attribute("GlyphCount"), 1) == 2 &&
                transform.Elements().Count() == 1 &&
                transform.Element(root.Name.Namespace + "Glyphs") is { } glyphs &&
                !glyphs.HasElements &&
                glyphs.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
                glyphs.Nodes().All(node => node is XText) &&
                glyphs.Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries) is { Length: 2 } tokens &&
                int.TryParse(tokens[0], NumberStyles.None, CultureInfo.InvariantCulture, out var first) &&
                int.TryParse(tokens[1], NumberStyles.None, CultureInfo.InvariantCulture, out var second) &&
                first is > 0 and <= ushort.MaxValue && second is > 0 and <= ushort.MaxValue && first != second)
                return new OfdExplicitGlyphRun(text.Runs, [first, second],
                    [new OfdExplicitGlyphSpan(0, 1, 0, 2, logical)]);
        }
        var glyphIds = new int[logical.Length];
        var assigned = new bool[logical.Length];
        var unicodeByGlyph = new Dictionary<int, char>();
        var spans = new List<OfdExplicitGlyphSpan>();
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
            spans.Add(new OfdExplicitGlyphSpan(position, codeCount, position, glyphCount,
                logical.Substring(position, codeCount)));
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
        var mixed = assigned.Any(value => !value);
        if (mixed)
        {
            if (text.Runs.Count > 1)
            {
                var priorEnd = 0;
                foreach (var span in spans)
                {
                    if (span.CodePosition < priorEnd)
                        throw new NotSupportedException("Mixed CGTransform across TextCode runs requires ascending global positions.");
                    priorEnd = span.CodePosition + span.CodeCount;
                }
            }
            // An unmapped character keeps the embedded face's direct Unicode
            // cmap GID. Never substitute a host font or .notdef for it.
            if (resource?.Data is not { Length: > 0 })
                throw new NotSupportedException("Mixed CGTransform requires an explicitly embedded font.");
            parseBudget ??= new FontParseBudget();
            parseBudget.Spend(0);
            if (cmapCache is null || !cmapCache.TryGetValue(resource, out var cmap))
            {
                // Selecting and parsing a face copies font tables. Charge the
                // conversion-wide budget before that work and reuse the cmap
                // for every mixed TextObject selecting this resource.
                parseBudget.Spend(resource.Data.LongLength);
                var face = new OpenTypeFace(OpenTypeCollection.SelectFace(resource.Data,
                    resource.CollectionFaceIndex));
                cmap = new OpenTypeCmap(face, parseBudget);
                cmapCache?.Add(resource, cmap);
            }
            if (cmap.IsSymbol)
                throw new NotSupportedException("Mixed CGTransform requires a direct Unicode cmap.");
            for (var index = 0; index < assigned.Length; index++)
            {
                parseBudget.Spend(0);
                if (assigned[index]) continue;
                var gid = cmap.DirectGlyph(logical[index]);
                if (gid <= 0 || gid > ushort.MaxValue)
                    throw new NotSupportedException("Mixed CGTransform ordinary character has no direct embedded-font glyph.");
                if (unicodeByGlyph.TryGetValue(gid, out var prior) && prior != logical[index])
                    throw new NotSupportedException("One original GID has ambiguous PDF Unicode semantics.");
                unicodeByGlyph[gid] = logical[index];
                glyphIds[index] = gid;
            }
        }
        foreach (var run in text.Runs)
        {
            if (mixed && run.Text.Length > 1 && string.IsNullOrWhiteSpace(run.DeltaX))
                throw new NotSupportedException("Mixed CGTransform requires explicit inter-glyph positioning deltas.");
            if (!string.IsNullOrWhiteSpace(run.DeltaY) && string.IsNullOrWhiteSpace(run.DeltaX))
                throw new NotSupportedException("CGTransform DeltaY without DeltaX is outside the render profile.");
            CountDeltas(run.DeltaX, run.Text.Length - 1);
            CountDeltas(run.DeltaY, run.Text.Length - 1);
        }
        return new OfdExplicitGlyphRun(text.Runs, glyphIds, spans);
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

    private static bool IsCombiningMark(char value) =>
        CharUnicodeInfo.GetUnicodeCategory(value) is UnicodeCategory.NonSpacingMark or
            UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark;

    private static bool SameAxis(XAttribute? attribute, double? previous, double actual) =>
        (attribute is not null || previous is not null) &&
        double.TryParse(attribute?.Value ?? previous?.ToString("R", CultureInfo.InvariantCulture),
            NumberStyles.Float, CultureInfo.InvariantCulture, out var value) &&
        !double.IsNaN(value) && !double.IsInfinity(value) && value == actual;
}
