using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Fonts;

// A symbol cmap is an encoded glyph repertoire, not a Unicode cmap. Only
// literal BMP private-use text with a direct, unambiguous mapping is admitted.
// The alias is confined to an export; the OFD keeps its original font bytes.
internal static class SymbolCmapProfile
{
    private const int FirstPrivateUse = 0xE000;
    private const int LastPrivateUse = 0xF8FF;
    private const int MaximumMappedGlyphs = 4096;

    internal static void ValidateText(OfdTextElement text, OpenTypeCmap symbol,
        CancellationToken cancellationToken = default)
    {
        if (!symbol.IsSymbol) throw new ArgumentException("Expected a Windows symbol cmap.", nameof(symbol));
        var values = text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text);
        foreach (var value in values)
            foreach (var scalar in OpenTypeFace.Scalars(value))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (scalar < FirstPrivateUse || scalar > LastPrivateUse || symbol.DirectGlyph(scalar) == 0)
                    throw new NotSupportedException($"Windows symbol cmap text U+{scalar:X4} has unmodeled character semantics without a direct private-use mapping; the original OFD font remains preserved.");
            }
    }

    internal static byte[] CreateUnicodeAlias(byte[] selectedFace, OpenTypeCmap symbol,
        FontParseBudget budget)
    {
        if (!symbol.IsSymbol) throw new ArgumentException("Expected a Windows symbol cmap.", nameof(symbol));
        if (selectedFace.LongLength > 64L * 1024 * 1024)
            throw new NotSupportedException("Symbol font exceeds the bounded static TrueType profile.");
        var face = new OpenTypeFace(selectedFace);
        if (face.Version != 0x00010000 || !face.Tables.ContainsKey("glyf") || !face.Tables.ContainsKey("loca") ||
            new[] { "CFF ", "CFF2", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
                "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "sbix" }.Any(face.Tables.ContainsKey))
            throw new NotSupportedException("Symbol export requires a static monochrome TrueType face.");
        if (!face.Tables.TryGetValue("OS/2", out var os2) || os2.Length < 10 ||
            (OpenTypeFace.U16(os2, 8) & (0x0002 | 0x0200)) != 0)
            throw new NotSupportedException("Symbol font does not permit the verified export embedding profile.");
        var glyphCount = OpenTypeFace.U16(face.Table("maxp", 6), 4);
        var mapping = new Dictionary<int, int>();
        var scalarByGlyph = new Dictionary<int, int>();
        for (var scalar = FirstPrivateUse; scalar <= LastPrivateUse; scalar++)
        {
            budget.Spend(1);
            var glyph = symbol.DirectGlyph(scalar);
            if (glyph == 0) continue;
            if (glyph >= glyphCount) throw new InvalidDataException("Symbol cmap references an absent glyph.");
            if (scalarByGlyph.TryGetValue(glyph, out var previous) && previous != scalar)
                throw new NotSupportedException("Symbol cmap maps one glyph to multiple private-use characters; PDF text extraction would be ambiguous.");
            if (mapping.Count == MaximumMappedGlyphs)
                throw new NotSupportedException("Symbol cmap exceeds the bounded private-use repertoire.");
            mapping.Add(scalar, glyph);
            scalarByGlyph.Add(glyph, scalar);
        }
        if (mapping.Count == 0)
            throw new NotSupportedException("Symbol cmap has no direct BMP private-use mappings.");
        face.Tables["cmap"] = OpenTypeCmap.Build(mapping);
        var alias = face.Build();
        var unicode = new OpenTypeCmap(new OpenTypeFace(alias), budget);
        if (unicode.IsSymbol || mapping.Any(item => unicode.DirectGlyph(item.Key) != item.Value))
            throw new InvalidDataException("Symbol Unicode export alias changed a glyph mapping.");
        return alias;
    }
}
