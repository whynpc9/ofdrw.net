using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>
/// Exact, conversion-local font bytes for OFD resources that declare a font name
/// without a FontFile. Supplying this policy makes selected name-only text strict:
/// every selected name must be present and covered by its bound face.
/// </summary>
public sealed class OfdNameOnlyFontBindings
{
    private const int MaximumFaceBytes = 16 * 1024 * 1024;
    private const int MaximumTotalBytes = 64 * 1024 * 1024;
    private const int MaximumFaces = 256;
    private static readonly string[] UnsupportedTables =
    [
        "CFF ", "CFF2", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
        "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "EBDT", "EBLC", "EBSC", "sbix",
        "acnt", "ankr", "bdat", "bloc", "bsln", "feat", "kerx", "lcar",
        "morx", "mort", "opbd", "prop", "trak", "Zapf", "Silf", "Glat", "Gloc", "Feat", "Sill"
    ];
    private readonly Dictionary<string, byte[]> _faces = new(StringComparer.Ordinal);

    /// <summary>Copies the caller's exact face bytes; names match OFD FontName with ordinal comparison.</summary>
    public OfdNameOnlyFontBindings(IReadOnlyDictionary<string, byte[]> faces)
    {
        if (faces is null) throw new ArgumentNullException(nameof(faces));
        long total = 0;
        foreach (var pair in faces)
        {
            if (_faces.Count >= MaximumFaces) throw new ArgumentException("More than 256 bound font names were supplied.", nameof(faces));
            if (string.IsNullOrWhiteSpace(pair.Key)) throw new ArgumentException("A nonblank OFD FontName is required.", nameof(faces));
            if (pair.Value is null || pair.Value.Length is 0 or > MaximumFaceBytes)
                throw new ArgumentException("A bound TrueType face must contain 1 to 16 MiB of bytes.", nameof(faces));
            total += pair.Value.Length;
            if (total > MaximumTotalBytes) throw new ArgumentException("Bound font bytes exceed 64 MiB.", nameof(faces));
            _faces.Add(pair.Key, (byte[])pair.Value.Clone());
        }
    }

    internal (byte[] Data, OpenTypeCmap Coverage) Resolve(OfdFontResource resource, FontParseBudget budget)
    {
        if (!_faces.TryGetValue(resource.FontName, out var data))
            throw new NotSupportedException($"Selected name-only font '{resource.FontName}' has no explicit face binding.");
        if (resource.Bold || resource.Italic || resource.CollectionFaceIndex != 0)
            throw new NotSupportedException("The explicit name-only profile requires one regular standalone TrueType face.");
        var face = new OpenTypeFace(data);
        var style = face.Style;
        if (face.Version != 0x00010000 || !face.Tables.ContainsKey("glyf") || !face.Tables.ContainsKey("loca") ||
            UnsupportedTables.Any(face.Tables.ContainsKey) || style.Bold || style.Italic)
            throw new NotSupportedException("The explicit name-only profile requires a regular static monochrome TrueType glyf face.");
        if (!face.Tables.TryGetValue("OS/2", out var os2) || os2.Length < 10 ||
            (OpenTypeFace.U16(os2, 8) & (0x0002 | 0x0100 | 0x0200)) != 0)
            throw new NotSupportedException("The bound font does not permit the verified embedding profile.");
        var glyphCount = OpenTypeFace.U16(face.Table("maxp", 6), 4);
        var locaFormat = OpenTypeFace.U16(face.Table("head", 54), 50);
        if (glyphCount == 0 || locaFormat is not (0 or 1))
            throw new NotSupportedException("The bound font has unsupported glyph locations.");
        OpenTypeFace.Require(face.Table("loca"), 0, (glyphCount + 1) * (locaFormat == 0 ? 2 : 4));
        var coverage = new OpenTypeCmap(face, budget);
        if (coverage.IsSymbol)
            throw new NotSupportedException("The bound font uses Windows symbol cmap semantics.");
        return (data, coverage);
    }

    internal static void ValidateBmpText(OfdTextElement text)
    {
        var values = text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text);
        if (values.Any(value => OpenTypeFace.Scalars(value).Any(scalar => scalar > 0xFFFF)))
            throw new NotSupportedException("Name-only explicit face binding is limited to BMP text.");
    }
}
