using System;
using System.IO;
using System.Security.Cryptography;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Layout.Text;

/// <summary>Outline technology of a supplied static face.</summary>
public enum TextFontOutlineFormat
{
    /// <summary>TrueType quadratic outlines (glyf/loca).</summary>
    TrueType,
    /// <summary>Static OpenType/CFF1 cubic outlines ('CFF ' table, including CID-keyed fonts).</summary>
    Cff,
}

/// <summary>An immutable, explicitly supplied static TrueType or OpenType/CFF1 face used by shaping and measurement.</summary>
public sealed class TextFontFace
{
    private const int MaximumFontBytes = 32 * 1024 * 1024;
    private readonly byte[] _faceBytes;
    private readonly OpenTypeCmap _cmap;
    private readonly int _glyphCount;

    /// <summary>Selects one face from supplied font bytes. Font names never trigger host discovery.</summary>
    /// <remarks>Static TrueType (glyf) and OpenType/CFF1 faces, including CID-keyed CFF such as Source Han
    /// Sans and Noto Sans CJK OTF, are admitted. Variable (fvar/gvar), CFF2 and color faces are rejected.</remarks>
    public TextFontFace(string name, byte[] data, int collectionFaceIndex = 0)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A font name is required.", nameof(name));
        if (data is null) throw new ArgumentNullException(nameof(data));
        if (data.Length == 0 || data.Length > MaximumFontBytes)
            throw new TextLayoutException(TextLayoutError.InvalidFont, "Font bytes are empty or exceed the 32 MiB face budget.");
        if (collectionFaceIndex < 0) throw new ArgumentOutOfRangeException(nameof(collectionFaceIndex));
        Name = name;
        CollectionFaceIndex = collectionFaceIndex;
        var copy = (byte[])data.Clone();
        SourceSha256 = Hash(copy);
        try
        {
            _faceBytes = OpenTypeCollection.SelectFace(copy, collectionFaceIndex, MaximumFontBytes);
            var face = new OpenTypeFace(_faceBytes);
            if (face.Tables.ContainsKey("fvar") || face.Tables.ContainsKey("CFF2") || face.Tables.ContainsKey("gvar") ||
                face.Tables.ContainsKey("COLR") || face.Tables.ContainsKey("SVG ") || face.Tables.ContainsKey("CBDT") ||
                face.Tables.ContainsKey("sbix"))
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Variable, CFF2 and color faces are outside the text profile; supply a static monochrome face.");
            var trueType = face.Version == 0x00010000 && face.Tables.ContainsKey("glyf") && face.Tables.ContainsKey("loca") &&
                !face.Tables.ContainsKey("CFF ");
            var cff = face.Version == 0x4F54544F && face.Tables.ContainsKey("CFF ") && !face.Tables.ContainsKey("glyf");
            if (!trueType && !cff)
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Only static TrueType (glyf) and OpenType/CFF1 outline faces are in the text profile.");
            OutlineFormat = trueType ? TextFontOutlineFormat.TrueType : TextFontOutlineFormat.Cff;
            var head = face.Table("head", 54);
            var hhea = face.Table("hhea", 36);
            var hmtx = face.Table("hmtx", 4);
            var maxp = face.Table("maxp", trueType ? 32 : 6);
            if (OpenTypeFace.U32(maxp, 0) != (trueType ? 0x00010000u : 0x00005000u) &&
                !(cff && OpenTypeFace.U32(maxp, 0) == 0x00010000u))
                throw new TextLayoutException(TextLayoutError.InvalidFont, "Invalid maxp version for the outline format.");
            _glyphCount = OpenTypeFace.U16(maxp, 4);
            if (_glyphCount == 0)
                throw new TextLayoutException(TextLayoutError.InvalidFont, "The selected font has no glyphs.");
            if (trueType) ValidateGlyphTables(face, head, hhea, hmtx, _glyphCount);
            else
            {
                ValidateHorizontalMetrics(hhea, hmtx, _glyphCount);
                // The CFF CharStrings INDEX must agree with maxp so that every shaped GID has a program.
                try
                {
                    if (CffFontProgram.Parse(face).GlyphCount != _glyphCount)
                        throw new TextLayoutException(TextLayoutError.InvalidFont, "CFF CharStrings count differs from maxp.");
                }
                catch (NotSupportedException exception)
                {
                    throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                        "The CFF table is outside the single-font CFF1 profile.", exception);
                }
            }
            UnitsPerEm = OpenTypeFace.U16(head, 18);
            if (UnitsPerEm < 16 || UnitsPerEm > 16384)
                throw new TextLayoutException(TextLayoutError.InvalidFont, "Invalid font units per em.");
            Ascender = unchecked((short)OpenTypeFace.U16(hhea, 4));
            Descender = unchecked((short)OpenTypeFace.U16(hhea, 6));
            LineGap = unchecked((short)OpenTypeFace.U16(hhea, 8));
            if (Ascender <= Descender || Ascender <= 0)
                throw new TextLayoutException(TextLayoutError.InvalidFont, "Invalid horizontal font extents.");
            _cmap = new OpenTypeCmap(face);
            (Bold, Italic) = face.Style;
            // Decoration metrics: post underline and OS/2 strikeout, each with a conventional fallback.
            UnderlinePosition = -UnitsPerEm / 10; UnderlineThickness = Math.Max(1, UnitsPerEm / 17);
            if (face.Tables.TryGetValue("post", out var post) && post.Length >= 12)
            {
                var position = unchecked((short)OpenTypeFace.U16(post, 8));
                var thickness = unchecked((short)OpenTypeFace.U16(post, 10));
                if (thickness > 0 && thickness <= UnitsPerEm / 4 && position < 0 && position > -UnitsPerEm / 2)
                { UnderlinePosition = position; UnderlineThickness = thickness; }
            }
            StrikeoutPosition = (int)Math.Round(Ascender * 0.3); StrikeoutThickness = UnderlineThickness;
            if (face.Tables.TryGetValue("OS/2", out var os2) && os2.Length >= 30)
            {
                var size = unchecked((short)OpenTypeFace.U16(os2, 26));
                var position = unchecked((short)OpenTypeFace.U16(os2, 28));
                if (size > 0 && size <= UnitsPerEm / 4 && position > 0 && position < Ascender)
                { StrikeoutPosition = position; StrikeoutThickness = size; }
            }
        }
        catch (TextLayoutException) { throw; }
        catch (NotSupportedException exception)
        {
            throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                "The selected font face needs an unsupported managed font profile.", exception);
        }
        catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException || exception is IndexOutOfRangeException)
        {
            throw new TextLayoutException(TextLayoutError.InvalidFont, "The selected font face cannot be parsed.", exception);
        }
        FaceSha256 = Hash(_faceBytes);
    }

    /// <summary>Caller supplied display name; it is never an identity or resolver key.</summary>
    public string Name { get; }
    /// <summary>Zero based face index selected from the supplied collection.</summary>
    public int CollectionFaceIndex { get; }
    /// <summary>SHA-256 of the complete caller supplied font bytes.</summary>
    public string SourceSha256 { get; }
    /// <summary>SHA-256 of the selected standalone face bytes used by shaping and embedding.</summary>
    public string FaceSha256 { get; }
    /// <summary>OpenType design units per em, used for one conversion to millimeters.</summary>
    public int UnitsPerEm { get; }
    /// <summary>Horizontal ascender in design units.</summary>
    public int Ascender { get; }
    /// <summary>Horizontal descender in design units.</summary>
    public int Descender { get; }
    /// <summary>Horizontal line gap in design units.</summary>
    public int LineGap { get; }
    /// <summary>Outline technology of the selected face.</summary>
    public TextFontOutlineFormat OutlineFormat { get; }
    /// <summary>Underline top-edge position below the baseline in design units (negative, Y up).</summary>
    public int UnderlinePosition { get; }
    /// <summary>Underline thickness in design units.</summary>
    public int UnderlineThickness { get; }
    /// <summary>Strikeout top-edge position above the baseline in design units (Y up).</summary>
    public int StrikeoutPosition { get; }
    /// <summary>Strikeout thickness in design units.</summary>
    public int StrikeoutThickness { get; }
    /// <summary>Whether the actual selected face declares a bold style.</summary>
    public bool Bold { get; }
    /// <summary>Whether the actual selected face declares an italic style.</summary>
    public bool Italic { get; }

    /// <summary>Returns an independent copy; callers cannot mutate this face.</summary>
    public byte[] GetFaceBytes() => (byte[])_faceBytes.Clone();

    /// <summary>Checks Unicode cmap coverage without selecting an implicit host fallback.</summary>
    public bool Covers(int unicodeScalar) => unicodeScalar >= 0 && unicodeScalar <= 0x10ffff &&
        CheckedGlyph(unicodeScalar) != 0;

    internal int NominalGlyph(int unicodeScalar) => CheckedGlyph(unicodeScalar);
    internal int GlyphCount => _glyphCount;
    internal int CoverageLookupWorkUpperBound => _cmap.GlyphLookupWorkUpperBound;
    internal int FaceByteLength => _faceBytes.Length;

    private static void ValidateHorizontalMetrics(byte[] hhea, byte[] hmtx, int glyphCount)
    {
        var horizontalMetrics = OpenTypeFace.U16(hhea, 34);
        if (horizontalMetrics == 0 || horizontalMetrics > glyphCount ||
            hmtx.Length < horizontalMetrics * 4 + (glyphCount - horizontalMetrics) * 2)
            throw new TextLayoutException(TextLayoutError.InvalidFont, "The font glyph count exceeds its horizontal metrics.");
    }

    private static void ValidateGlyphTables(OpenTypeFace face, byte[] head, byte[] hhea, byte[] hmtx, int glyphCount)
    {
        ValidateHorizontalMetrics(hhea, hmtx, glyphCount);

        var locationFormat = OpenTypeFace.U16(head, 50);
        if (locationFormat != 0 && locationFormat != 1)
            throw new TextLayoutException(TextLayoutError.InvalidFont, "Invalid TrueType glyph location format.");
        var loca = face.Table("loca");
        var glyf = face.Table("glyf");
        var locationSize = locationFormat == 0 ? 2 : 4;
        OpenTypeFace.Require(loca, 0, (glyphCount + 1) * locationSize);
        var previous = 0;
        for (var index = 0; index <= glyphCount; index++)
        {
            var offset = locationFormat == 0
                ? OpenTypeFace.U16(loca, index * 2) * 2L
                : OpenTypeFace.U32(loca, index * 4);
            if (offset < previous || offset > glyf.Length)
                throw new TextLayoutException(TextLayoutError.InvalidFont, "Invalid TrueType glyph location.");
            previous = (int)offset;
        }
    }

    private int CheckedGlyph(int unicodeScalar)
    {
        try
        {
            var glyph = _cmap.Glyph(unicodeScalar);
            if (glyph >= _glyphCount)
                throw new InvalidDataException("The font cmap maps a scalar beyond maxp.numGlyphs.");
            return glyph;
        }
        catch (Exception exception) when (exception is InvalidDataException || exception is OverflowException ||
            exception is IndexOutOfRangeException)
        {
            throw new TextLayoutException(TextLayoutError.InvalidFont,
                "The selected font cmap failed during glyph lookup.", exception);
        }
    }

    private static string Hash(byte[] data)
    {
        using var sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(data)).Replace("-", "").ToLowerInvariant();
    }
}
