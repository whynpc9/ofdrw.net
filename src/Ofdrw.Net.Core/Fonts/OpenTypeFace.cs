using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Ofdrw.Net.Core.Fonts;

// The shared, bounded sfnt reader/writer used by collection extraction, PDF
// name isolation, coverage validation and package subsetting. Table bytes are
// snapshots: no converter may modify the caller's font buffer.
internal sealed class OpenTypeFace
{
    internal uint Version { get; }
    internal Dictionary<string, byte[]> Tables { get; } = new(StringComparer.Ordinal);

    internal OpenTypeFace(byte[] data, int faceOffset = 0)
    {
        Require(data, faceOffset, 12);
        Version = U32(data, faceOffset);
        if (Version != 0x00010000 && Version != 0x4F54544F && Version != 0x74727565)
            throw new InvalidDataException("Expected a standalone TrueType/OpenType face.");
        var count = U16(data, faceOffset + 4);
        if (count == 0 || count > 4095) throw new InvalidDataException("Invalid OpenType table count.");
        Require(data, faceOffset + 12, count * 16);
        var records = new List<(string Tag, int Offset, int Length)>();
        for (var index = 0; index < count; index++)
        {
            var record = faceOffset + 12 + index * 16;
            var tag = Encoding.ASCII.GetString(data, record, 4);
            var offset = Size(U32(data, record + 8));
            var length = Size(U32(data, record + 12));
            Require(data, offset, length);
            if (records.Any(value => value.Tag == tag)) throw new InvalidDataException("Duplicate OpenType table: " + tag);
            records.Add((tag, offset, length));
        }
        long previousEnd = -1;
        foreach (var record in records.Where(value => value.Length > 0).OrderBy(value => value.Offset))
        {
            if (record.Offset < previousEnd || record.Offset < faceOffset + 12L + count * 16 && record.Offset + (long)record.Length > faceOffset)
                throw new InvalidDataException("Overlapping OpenType tables/directory.");
            previousEnd = record.Offset + (long)record.Length;
        }
        foreach (var record in records)
        {
            var bytes = new byte[record.Length];
            Buffer.BlockCopy(data, record.Offset, bytes, 0, record.Length);
            Tables.Add(record.Tag, bytes);
        }
    }

    internal byte[] Table(string tag, int minimumLength = 0)
    {
        if (!Tables.TryGetValue(tag, out var table)) throw new InvalidDataException("Missing OpenType table: " + tag);
        Require(table, 0, minimumLength);
        return table;
    }

    internal (bool Bold, bool Italic) Style
    {
        get
        {
            if (Tables.TryGetValue("OS/2", out var os2))
            { var flags = U16(os2, 62); return ((flags & 0x20) != 0, (flags & 1) != 0); }
            var macFlags = U16(Table("head", 54), 44);
            return ((macFlags & 1) != 0, (macFlags & 2) != 0);
        }
    }

    internal byte[] Build()
    {
        var tables = Tables.Where(pair => pair.Key != "DSIG").OrderBy(pair => pair.Key, StringComparer.Ordinal).ToArray();
        var head = Table("head", 54);
        Put32(head, 8, 0);
        var size = checked(12 + tables.Length * 16);
        foreach (var table in tables) size = checked(size + Align(table.Value.Length));
        var output = new byte[size];
        Put32(output, 0, Version); Put16(output, 4, tables.Length);
        var power = 1; var selector = 0;
        while (power * 2 <= tables.Length) { power *= 2; selector++; }
        Put16(output, 6, power * 16); Put16(output, 8, selector);
        Put16(output, 10, tables.Length * 16 - power * 16);
        var cursor = 12 + tables.Length * 16; var headOffset = 0;
        for (var index = 0; index < tables.Length; index++)
        {
            var table = tables[index]; var record = 12 + index * 16;
            Encoding.ASCII.GetBytes(table.Key).CopyTo(output, record);
            Put32(output, record + 4, Checksum(table.Value));
            Put32(output, record + 8, (uint)cursor); Put32(output, record + 12, (uint)table.Value.Length);
            table.Value.CopyTo(output, cursor);
            if (table.Key == "head") headOffset = cursor;
            cursor = checked(cursor + Align(table.Value.Length));
        }
        Put32(output, headOffset + 8, unchecked(0xB1B0AFBAu - Checksum(output)));
        return output;
    }

    internal static IEnumerable<int> Scalars(string text)
    {
        for (var i = 0; i < text.Length; i++)
        {
            var value = text[i];
            if (!char.IsSurrogate(value)) yield return value;
            else if (char.IsHighSurrogate(value) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1]))
                yield return char.ConvertToUtf32(value, text[++i]);
            else throw new InvalidDataException("Font usage contains invalid UTF-16.");
        }
    }

    internal static int Size(uint value)
    {
        if (value > int.MaxValue) throw new InvalidDataException("OpenType size exceeds supported bounds.");
        return (int)value;
    }
    internal static int Align(int value) => checked((value + 3) & ~3);
    internal static void Require(byte[] bytes, int offset, int length)
    {
        if (offset < 0 || length < 0 || (long)offset + length > bytes.Length)
            throw new InvalidDataException("Invalid OpenType table bounds.");
    }
    internal static int U16(byte[] bytes, int offset)
    { Require(bytes, offset, 2); return (bytes[offset] << 8) | bytes[offset + 1]; }
    internal static uint U32(byte[] bytes, int offset)
    { Require(bytes, offset, 4); return ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3]; }
    internal static void Put16(byte[] bytes, int offset, int value)
    {
        if (value < 0 || value > ushort.MaxValue) throw new InvalidDataException("OpenType value exceeds 16-bit bounds.");
        bytes[offset] = (byte)(value >> 8); bytes[offset + 1] = (byte)value;
    }
    internal static void Put32(byte[] bytes, int offset, uint value)
    { bytes[offset] = (byte)(value >> 24); bytes[offset + 1] = (byte)(value >> 16); bytes[offset + 2] = (byte)(value >> 8); bytes[offset + 3] = (byte)value; }
    internal static uint Checksum(byte[] bytes)
    {
        uint sum = 0;
        for (var i = 0; i < bytes.Length; i += 4)
        {
            uint word = 0;
            for (var j = 0; j < 4; j++) word = (word << 8) | (i + j < bytes.Length ? bytes[i + j] : 0u);
            sum = unchecked(sum + word);
        }
        return sum;
    }
}
