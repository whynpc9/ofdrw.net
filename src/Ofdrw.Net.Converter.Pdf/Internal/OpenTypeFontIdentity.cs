using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Gives PDFsharp's internal name-based caches the same content identity as its
/// resolver. Glyphs, metrics, layout tables and copyright records are retained.
/// The caller's original font is never modified.
/// </summary>
internal static class OpenTypeFontIdentity
{
    internal static byte[] WithUniqueNames(byte[] source, string identity)
    {
        Require(source, 0, 12);
        var count = U16(source, 4);
        Require(source, 12, count * 16);
        var tables = new List<(string Tag, byte[] Data)>();
        for (var index = 0; index < count; index++)
        {
            var record = 12 + index * 16;
            var tag = Encoding.ASCII.GetString(source, record, 4);
            var offset = checked((int)U32(source, record + 8));
            var length = checked((int)U32(source, record + 12));
            Require(source, offset, length);
            if (tag == "DSIG") continue; // a signature cannot survive any font-table edit
            var data = new byte[length];
            Buffer.BlockCopy(source, offset, data, 0, length);
            if (tag == "name") data = RewriteNames(data, identity);
            if (tag == "head") { Require(data, 0, 12); Put32(data, 8, 0); }
            tables.Add((tag, data));
        }
        if (!tables.Any(table => table.Tag == "name") || !tables.Any(table => table.Tag == "head"))
            throw new InvalidDataException("Embedded font must contain OpenType name and head tables.");

        var size = checked(12 + tables.Count * 16 + tables.Sum(table => Align(table.Data.Length)));
        var result = new byte[size];
        Buffer.BlockCopy(source, 0, result, 0, 4);
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
        Put32(result, headOffset + 8, unchecked(0xB1B0AFBAu - Checksum(result)));
        return result;
    }

    private static byte[] RewriteNames(byte[] table, string identity)
    {
        Require(table, 0, 6);
        var version = U16(table, 0);
        if (version > 1) throw new NotSupportedException("Unsupported OpenType naming table version.");
        var count = U16(table, 2);
        var storage = U16(table, 4);
        Require(table, 6, count * 12);
        var recordsEnd = 6 + count * 12;
        var languageCount = version == 1 ? U16(table, recordsEnd) : 0;
        var headerLength = recordsEnd + (version == 1 ? 2 + languageCount * 4 : 0);
        Require(table, 0, headerLength);
        var records = new byte[headerLength];
        Buffer.BlockCopy(table, 0, records, 0, headerLength);
        Put16(records, 4, headerLength);
        using var strings = new MemoryStream();
        var stringOffsets = new Dictionary<string, int>(StringComparer.Ordinal);
        int Intern(byte[] bytes, int offset, int length)
        {
            var key = Convert.ToBase64String(bytes, offset, length);
            if (stringOffsets.TryGetValue(key, out var known)) return known;
            var position = checked((int)strings.Length);
            strings.Write(bytes, offset, length);
            stringOffsets.Add(key, position);
            return position;
        }
        for (var index = 0; index < count; index++)
        {
            var record = 6 + index * 12;
            var platform = U16(table, record);
            var name = U16(table, record + 6);
            var length = U16(table, record + 8);
            var offset = U16(table, record + 10);
            Require(table, storage + offset, length);
            byte[] value;
            if (name is 1 or 3 or 4 or 6 or 16 or 18 or 21 && platform is 0 or 1 or 3)
            {
                var text = name == 6 ? "O" + identity.Substring(0, Math.Min(62, identity.Length)) : "Ofdrw-" + identity;
                value = platform == 1 ? Encoding.ASCII.GetBytes(text) : Encoding.BigEndianUnicode.GetBytes(text);
            }
            else
            {
                value = new byte[length];
                Buffer.BlockCopy(table, storage + offset, value, 0, length);
            }
            Put16(records, record + 8, value.Length);
            Put16(records, record + 10, Intern(value, 0, value.Length));
        }
        for (var index = 0; index < languageCount; index++)
        {
            var record = recordsEnd + 2 + index * 4;
            var length = U16(table, record);
            var offset = U16(table, record + 2);
            Require(table, storage + offset, length);
            Put16(records, record + 2, Intern(table, storage + offset, length));
        }
        using var output = new MemoryStream();
        output.Write(records, 0, records.Length);
        strings.Position = 0;
        strings.CopyTo(output);
        return output.ToArray();
    }

    private static int Align(int value) => checked((value + 3) & ~3);
    private static void Require(byte[] data, int offset, int count)
    {
        if (offset < 0 || count < 0 || (long)offset + count > data.Length)
            throw new InvalidDataException("Invalid OpenType table bounds.");
    }
    private static int U16(byte[] data, int offset) { Require(data, offset, 2); return (data[offset] << 8) | data[offset + 1]; }
    private static uint U32(byte[] data, int offset)
    {
        Require(data, offset, 4);
        return ((uint)data[offset] << 24) | ((uint)data[offset + 1] << 16) | ((uint)data[offset + 2] << 8) | data[offset + 3];
    }
    private static void Put16(byte[] data, int offset, int value)
    {
        if (value < 0 || value > ushort.MaxValue) throw new InvalidDataException("OpenType naming data exceeds its 16-bit bounds.");
        data[offset] = (byte)(value >> 8); data[offset + 1] = (byte)value;
    }
    private static void Put32(byte[] data, int offset, uint value)
    {
        data[offset] = (byte)(value >> 24); data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8); data[offset + 3] = (byte)value;
    }
    private static uint Checksum(byte[] data)
    {
        uint sum = 0;
        for (var index = 0; index < data.Length; index += 4)
        {
            uint word = 0;
            for (var offset = 0; offset < 4; offset++) word = (word << 8) | (index + offset < data.Length ? data[index + offset] : 0u);
            sum = unchecked(sum + word);
        }
        return sum;
    }
}
