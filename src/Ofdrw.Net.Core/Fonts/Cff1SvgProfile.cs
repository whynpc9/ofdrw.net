using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Fonts;

// Structural admission for an embedded, static OpenType/CFF1 face selected by
// SVG text. This does not interpret Type 2 CharStrings or substitute a GID for
// Unicode; the SVG renderer consumes the original font and logical text.
internal static class Cff1SvgProfile
{
    private const int MaximumFaceBytes = 16 * 1024 * 1024;
    private static readonly string[] UnsupportedTables =
    [
        "CFF2", "glyf", "loca", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
        "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "EBDT", "EBLC", "sbix",
        "acnt", "ankr", "bdat", "bloc", "bsln", "feat", "kerx", "lcar",
        "morx", "mort", "opbd", "prop", "trak", "Zapf", "Silf", "Glat", "Gloc", "Feat", "Sill"
    ];

    internal static bool IsCffOpenType(OpenTypeFace face) =>
        face.Version == 0x4F54544F && (face.Tables.ContainsKey("CFF ") || face.Tables.ContainsKey("CFF2"));

    internal static int Validate(byte[] source, OpenTypeFace face)
    {
        if (source.Length is < 12 or > MaximumFaceBytes || face.Version != 0x4F54544F ||
            !face.Tables.ContainsKey("CFF ") || UnsupportedTables.Any(face.Tables.ContainsKey))
            throw new NotSupportedException("SVG text requires a bounded static monochrome OpenType/CFF1 face.");
        var style = face.Style;
        if (style.Bold || style.Italic)
            throw new NotSupportedException("The SVG CFF1 profile requires a regular face without synthetic style.");
        var os2 = face.Table("OS/2", 10);
        if ((OpenTypeFace.U16(os2, 8) & (0x0002 | 0x0200)) != 0)
            throw new NotSupportedException("The selected CFF1 font does not permit this embedding profile.");
        var glyphCount = OpenTypeFace.U16(face.Table("maxp", 6), 4);
        if (glyphCount == 0 || OpenTypeFace.U16(face.Table("head", 54), 18) is < 16 or > 16384)
            throw new InvalidDataException("Invalid CFF1 face metrics.");
        var cff = face.Table("CFF ", 4);
        if (cff[0] != 1 || cff[2] < 4 || cff[2] > cff.Length || cff[3] is < 1 or > 4)
            throw new NotSupportedException("SVG text requires a valid CFF1 header.");
        var names = ReadIndex(cff, cff[2]);
        var top = ReadIndex(cff, names.End);
        if (names.Count != 1 || top.Count != 1)
            throw new NotSupportedException("SVG text requires one standalone CFF1 font and Top DICT.");
        var strings = ReadIndex(cff, top.End);
        var globalSubrs = ReadIndex(cff, strings.End);
        var charStringsOffset = ReadCharStringsOffset(cff, top.ItemStart(0), top.ItemEnd(0));
        if (charStringsOffset < globalSubrs.End)
            throw new InvalidDataException("CFF1 CharStrings INDEX overlaps the header indexes.");
        var charStrings = ReadIndex(cff, charStringsOffset);
        if (charStrings.Count != glyphCount)
            throw new InvalidDataException("CFF1 CharStrings count differs from maxp glyph count.");
        for (var gid = 0; gid < charStrings.Count; gid++)
            if (charStrings.ItemStart(gid) == charStrings.ItemEnd(gid))
                throw new InvalidDataException("CFF1 CharStrings INDEX contains an empty glyph program.");
        return glyphCount;
    }

    internal static void ValidateSelectedText(OfdTextElement text, OfdFontResource font,
        OpenTypeCmap coverage, int glyphCount, FontParseBudget budget)
    {
        if (font.Bold || font.Italic || text.Weight >= 600 || text.Italic)
            throw new NotSupportedException("SVG CFF1 text requires a regular face without synthetic emphasis.");
        foreach (var value in text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text))
            foreach (var scalar in OpenTypeFace.Scalars(value))
            {
                budget.Spend(coverage.GlyphLookupWorkUpperBound);
                if (coverage.Glyph(scalar) >= glyphCount)
                    throw new InvalidDataException("CFF1 cmap GID exceeds the CharStrings INDEX.");
            }
    }

    internal static void ValidateVariationGlyphs(OpenTypeCmap coverage, int glyphCount, FontParseBudget budget)
    {
        foreach (var gid in coverage.VariationGlyphs)
        {
            budget.Spend(1);
            if (gid == 0 || gid >= glyphCount)
                throw new InvalidDataException("CFF1 non-default UVS cmap GID is outside the CharStrings INDEX.");
        }
    }

    private static int ReadCharStringsOffset(byte[] cff, int start, int end)
    {
        var operands = new List<long?>();
        int? charStrings = null;
        for (var cursor = start; cursor < end;)
        {
            var value = cff[cursor++];
            if (value == 28 || value == 29)
            {
                var bytes = value == 28 ? 2 : 4;
                OpenTypeFace.Require(cff, cursor, bytes);
                if (cursor + bytes > end) throw new InvalidDataException("Truncated CFF1 Top DICT operand.");
                long number = value == 28 ? (short)OpenTypeFace.U16(cff, cursor) :
                    unchecked((int)OpenTypeFace.U32(cff, cursor));
                operands.Add(number); cursor += bytes;
            }
            else if (value == 30)
            {
                // A real number is legal in a DICT, but not as a CharStrings
                // INDEX offset. Bound its nibble scan even for hostile bytes.
                var terminated = false;
                for (var bytes = 0; bytes < 32 && cursor < end; bytes++)
                {
                    var pair = cff[cursor++];
                    if ((pair >> 4) == 15 || (pair & 15) == 15) { terminated = true; break; }
                }
                if (!terminated) throw new InvalidDataException("Unterminated CFF1 Top DICT real operand.");
                operands.Add(null);
            }
            else if (value >= 32 && value <= 254)
            {
                long number;
                if (value <= 246) number = value - 139;
                else
                {
                    if (cursor >= end) throw new InvalidDataException("Truncated CFF1 Top DICT operand.");
                    var tail = cff[cursor++];
                    number = value <= 250 ? (value - 247) * 256 + tail + 108 :
                        -(value - 251) * 256 - tail - 108;
                }
                operands.Add(number);
            }
            else if (value == 12)
            {
                if (cursor >= end) throw new InvalidDataException("Truncated CFF1 escaped operator.");
                var op = cff[cursor++];
                if (op == 30) throw new NotSupportedException("CID-keyed CFF1 is outside the SVG text profile.");
                if (op == 6 && (operands.Count != 1 || operands[0] != 2))
                    throw new NotSupportedException("SVG text requires Type 2 CFF1 CharStrings.");
                operands.Clear();
            }
            else if (value <= 21)
            {
                if (value == 17)
                {
                    if (charStrings.HasValue || operands.Count != 1 || operands[0] is not long offset ||
                        offset < 0 || offset > int.MaxValue)
                        throw new InvalidDataException("Invalid CFF1 CharStrings INDEX offset.");
                    charStrings = (int)offset;
                }
                operands.Clear();
            }
            else throw new InvalidDataException("Unsupported CFF1 Top DICT operand.");
            if (operands.Count > 48) throw new InvalidDataException("CFF1 Top DICT operand stack exceeded.");
        }
        return charStrings ?? throw new InvalidDataException("CFF1 Top DICT lacks CharStrings INDEX.");
    }

    private static Index ReadIndex(byte[] cff, int start)
    {
        var count = OpenTypeFace.U16(cff, start);
        if (count == 0) return new Index(start + 2, start + 2, Array.Empty<int>());
        var offSizePos = checked(start + 2);
        OpenTypeFace.Require(cff, offSizePos, 1);
        var offSize = cff[offSizePos];
        if (offSize is < 1 or > 4) throw new InvalidDataException("Invalid CFF1 INDEX offset size.");
        var offsetsStart = checked(offSizePos + 1);
        var dataStart = checked(offsetsStart + (count + 1) * offSize);
        OpenTypeFace.Require(cff, offsetsStart, (count + 1) * offSize);
        var offsets = new int[count + 1];
        for (var i = 0; i <= count; i++)
        {
            uint value = 0;
            for (var j = 0; j < offSize; j++) value = (value << 8) | cff[offsetsStart + i * offSize + j];
            if (value == 0 || value > int.MaxValue || i == 0 && value != 1 || i > 0 && value < offsets[i - 1])
                throw new InvalidDataException("Invalid CFF1 INDEX offsets.");
            offsets[i] = (int)value;
        }
        var end = checked(dataStart + offsets[count] - 1);
        OpenTypeFace.Require(cff, dataStart, end - dataStart);
        return new Index(dataStart, end, offsets);
    }

    private readonly struct Index(int dataStart, int end, int[] offsets)
    {
        internal int Count => offsets.Length == 0 ? 0 : offsets.Length - 1;
        internal int End { get; } = end;
        internal int ItemStart(int index) => checked(dataStart + offsets[index] - 1);
        internal int ItemEnd(int index) => checked(dataStart + offsets[index + 1] - 1);
    }
}
