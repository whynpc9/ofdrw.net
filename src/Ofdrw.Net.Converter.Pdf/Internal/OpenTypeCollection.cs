using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Copies each face in a TrueType/OpenType collection into a standalone sfnt.
/// Windows CJK fonts such as SimSun ship as .ttc; OFD FontFile and PDFsharp need TTF.
/// </summary>
internal static class OpenTypeCollection
{
    internal static IReadOnlyList<byte[]> ExtractFaces(byte[] source)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (source.Length < 12) throw new InvalidDataException("Font data is too small.");
        var tag = Encoding.ASCII.GetString(source, 0, 4);
        if (tag != "ttcf") return new[] { source };

        var count = checked((int)U32(source, 8));
        if (count <= 0 || count > 64) throw new InvalidDataException("TrueType collection face count is invalid.");
        Require(source, 12, count * 4);
        var faces = new List<byte[]>(count);
        for (var index = 0; index < count; index++)
        {
            var offset = checked((int)U32(source, 12 + index * 4));
            faces.Add(CopySfnt(source, offset));
        }

        return faces;
    }

    private static byte[] CopySfnt(byte[] source, int offset)
    {
        Require(source, offset, 12);
        var count = U16(source, offset + 4);
        Require(source, offset + 12, count * 16);
        var tables = new List<(string Tag, byte[] Data)>(count);
        for (var index = 0; index < count; index++)
        {
            var record = offset + 12 + index * 16;
            var tag = Encoding.ASCII.GetString(source, record, 4);
            var tableOffset = checked((int)U32(source, record + 8));
            var length = checked((int)U32(source, record + 12));
            Require(source, tableOffset, length);
            if (tag == "DSIG") continue;
            var data = new byte[length];
            Buffer.BlockCopy(source, tableOffset, data, 0, length);
            tables.Add((tag, data));
        }

        if (!tables.Any(table => table.Tag == "name") || !tables.Any(table => table.Tag == "head"))
            throw new InvalidDataException("Collection face must contain OpenType name and head tables.");

        var size = checked(12 + tables.Count * 16 + tables.Sum(table => Align(table.Data.Length)));
        var result = new byte[size];
        Buffer.BlockCopy(source, offset, result, 0, 4);
        Put16(result, 4, tables.Count);
        var power = 1;
        var selector = 0;
        while (power * 2 <= tables.Count) { power *= 2; selector++; }
        Put16(result, 6, power * 16);
        Put16(result, 8, selector);
        Put16(result, 10, tables.Count * 16 - power * 16);
        var cursor = 12 + tables.Count * 16;
        var headOffset = -1;
        for (var index = 0; index < tables.Count; index++)
        {
            var table = tables[index];
            var record = 12 + index * 16;
            Encoding.ASCII.GetBytes(table.Tag).CopyTo(result, record);
            Put32(result, record + 4, Checksum(table.Data));
            Put32(result, record + 8, (uint)cursor);
            Put32(result, record + 12, (uint)table.Data.Length);
            table.Data.CopyTo(result, cursor);
            if (table.Tag == "head") headOffset = cursor;
            cursor += Align(table.Data.Length);
        }

        Require(result, headOffset, 12);
        Put32(result, headOffset + 8, 0);
        Put32(result, headOffset + 8, unchecked(0xB1B0AFBAu - Checksum(result)));
        return result;
    }

    private static int Align(int value) => checked((value + 3) & ~3);

    private static void Require(byte[] data, int offset, int count)
    {
        if (offset < 0 || count < 0 || (long)offset + count > data.Length)
            throw new InvalidDataException("Invalid OpenType collection bounds.");
    }

    private static int U16(byte[] data, int offset)
    {
        Require(data, offset, 2);
        return (data[offset] << 8) | data[offset + 1];
    }

    private static uint U32(byte[] data, int offset)
    {
        Require(data, offset, 4);
        return ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) |
               ((uint)data[offset + 2] << 8) | data[offset + 3];
    }

    private static void Put16(byte[] data, int offset, int value)
    {
        data[offset] = (byte)(value >> 8);
        data[offset + 1] = (byte)value;
    }

    private static void Put32(byte[] data, int offset, uint value)
    {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static uint Checksum(byte[] data)
    {
        uint sum = 0;
        for (var index = 0; index < data.Length; index += 4)
        {
            uint word = 0;
            for (var skip = 0; skip < 4; skip++)
                word = (word << 8) | (index + skip < data.Length ? data[index + skip] : 0u);
            sum = unchecked(sum + word);
        }

        return sum;
    }
}
