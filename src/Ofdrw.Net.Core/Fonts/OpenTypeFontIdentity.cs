using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>
/// Gives PDFsharp's internal name-based caches the same content identity as its
/// resolver. Glyphs, metrics, layout tables and copyright records are retained.
/// The caller's original font is never modified.
/// </summary>
internal static class OpenTypeFontIdentity
{
    internal static byte[] WithUniqueNames(byte[] source, string identity)
    {
        var face = new OpenTypeFace(source);
        face.Tables["name"] = RewriteNames(face.Table("name", 6), identity);
        return face.Build();
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

    private static void Require(byte[] data, int offset, int count) => OpenTypeFace.Require(data, offset, count);
    private static int U16(byte[] data, int offset) => OpenTypeFace.U16(data, offset);
    private static void Put16(byte[] data, int offset, int value) => OpenTypeFace.Put16(data, offset, value);
}
