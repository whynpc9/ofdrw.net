using System;
using System.IO;
using System.Security.Cryptography;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Layout.Text;

/// <summary>An immutable, explicitly supplied static TrueType face used by shaping and measurement.</summary>
public sealed class TextFontFace
{
    private const int MaximumFontBytes = 32 * 1024 * 1024;
    private readonly byte[] _faceBytes;
    private readonly OpenTypeCmap _cmap;

    /// <summary>Selects one face from supplied font bytes. Font names never trigger host discovery.</summary>
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
            if (face.Version != 0x00010000 || !face.Tables.ContainsKey("glyf") || !face.Tables.ContainsKey("loca") ||
                face.Tables.ContainsKey("fvar") || face.Tables.ContainsKey("CFF ") || face.Tables.ContainsKey("CFF2") ||
                face.Tables.ContainsKey("COLR") || face.Tables.ContainsKey("SVG ") || face.Tables.ContainsKey("CBDT") ||
                face.Tables.ContainsKey("sbix"))
                throw new TextLayoutException(TextLayoutError.UnsupportedFontProfile,
                    "Only static, monochrome TrueType outline faces are in the initial text profile.");
            var head = face.Table("head", 54);
            var hhea = face.Table("hhea", 36);
            face.Table("hmtx", 4);
            face.Table("maxp", 6);
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

    private int CheckedGlyph(int unicodeScalar)
    {
        try { return _cmap.Glyph(unicodeScalar); }
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
