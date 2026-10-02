using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
namespace Ofdrw.Net.Core.Fonts;

internal static class OpenTypeCollection
{
    internal static byte[] SelectFace(byte[] source, int index, long maximumFaceBytes = long.MaxValue)
    {
        OpenTypeFace.Require(source, 0, 12);
        if (OpenTypeFace.U32(source, 0) != 0x74746366)
        {
            if (index != 0) throw new InvalidDataException("Standalone font has only face zero.");
            if (source.LongLength > maximumFaceBytes) throw new InvalidDataException("Font exceeds MaximumFontBytes.");
            return source;
        }
        var version = OpenTypeFace.U32(source, 4);
        if (version != 0x00010000 && version != 0x00020000) throw new InvalidDataException("Invalid collection version.");
        var count = OpenTypeFace.Size(OpenTypeFace.U32(source, 8));
        if (count <= 0 || count > 64 || index < 0 || index >= count) throw new InvalidDataException("Invalid collection face index/count.");
        OpenTypeFace.Require(source, 12, count * 4);
        var faceOffset = OpenTypeFace.Size(OpenTypeFace.U32(source, 12 + index * 4));
        OpenTypeFace.Require(source, faceOffset, 12);
        var tableCount = OpenTypeFace.U16(source, faceOffset + 4);
        if (tableCount == 0 || tableCount > 4095) throw new InvalidDataException("Invalid OpenType table count.");
        OpenTypeFace.Require(source, faceOffset + 12, tableCount * 16);
        long outputSize = 12;
        for (var table = 0; table < tableCount; table++)
        {
            var record = faceOffset + 12 + table * 16;
            if (source[record] == 'D' && source[record + 1] == 'S' && source[record + 2] == 'I' && source[record + 3] == 'G') continue;
            outputSize += 16L + OpenTypeFace.Align(OpenTypeFace.Size(OpenTypeFace.U32(source, record + 12)));
            if (outputSize > maximumFaceBytes) throw new InvalidDataException("Selected TTC face exceeds MaximumFontBytes.");
        }
        return new OpenTypeFace(source, faceOffset).Build();
    }
    // Host resolver face keys are opaque. A collection must not be guessed as
    // face zero. Full/PostScript names take precedence over family names.
    internal static byte[] SelectNamedFace(byte[] source, string faceName, bool allowSingleFace = false)
    {
        OpenTypeFace.Require(source, 0, 12);
        if (OpenTypeFace.U32(source, 0) != 0x74746366) return source;
        if (source.LongLength > 256L * 1024 * 1024) throw new NotSupportedException("Host collection exceeds the input font budget.");
        var version = OpenTypeFace.U32(source, 4);
        if (version != 0x00010000 && version != 0x00020000) throw new InvalidDataException("Invalid collection version.");
        var count = OpenTypeFace.Size(OpenTypeFace.U32(source, 8));
        if (count <= 0 || count > 64) throw new InvalidDataException("Invalid collection face count.");
        OpenTypeFace.Require(source, 12, count * 4);
        if (count == 1 && allowSingleFace) return SelectFace(source, 0, 64L * 1024 * 1024);
        var strong = new HashSet<int>(); var family = new HashSet<int>(); long decodedBytes = 0;
        var unicode = new UnicodeEncoding(true, false, true);
        for (var index = 0; index < count; index++)
        {
            var offset = OpenTypeFace.Size(OpenTypeFace.U32(source, 12 + index * 4));
            OpenTypeFace.Require(source, offset, 12);
            var signature = OpenTypeFace.U32(source, offset);
            if (signature != 0x00010000 && signature != 0x4F54544F && signature != 0x74727565) throw new InvalidDataException("Invalid collection face signature.");
            var tables = OpenTypeFace.U16(source, offset + 4);
            if (tables == 0 || tables > 4095) throw new InvalidDataException("Invalid collection table count.");
            OpenTypeFace.Require(source, offset + 12, tables * 16);
            var named = false;
            for (var table = 0; table < tables; table++)
            {
                var record = offset + 12 + table * 16;
                if (OpenTypeFace.U32(source, record) != 0x6E616D65) continue;
                if (named) throw new InvalidDataException("Duplicate naming table."); named = true;
                var start = OpenTypeFace.Size(OpenTypeFace.U32(source, record + 8));
                var length = OpenTypeFace.Size(OpenTypeFace.U32(source, record + 12));
                OpenTypeFace.Require(source, start, length);
                if (length < 6 || length > 1024 * 1024) throw new InvalidDataException("Invalid or over-budget naming table.");
                var namingVersion = OpenTypeFace.U16(source, start);
                if (namingVersion > 1) throw new NotSupportedException("Unsupported font naming table.");
                var names = OpenTypeFace.U16(source, start + 2); var storage = OpenTypeFace.U16(source, start + 4);
                var recordsEnd = 6 + names * 12;
                if (recordsEnd > length) throw new InvalidDataException("Invalid naming records.");
                if (namingVersion == 1)
                {
                    if (recordsEnd + 2 > length) throw new InvalidDataException("Missing font language records.");
                    recordsEnd += 2 + OpenTypeFace.U16(source, start + recordsEnd) * 4;
                }
                if (recordsEnd > length || storage < recordsEnd || storage > length) throw new InvalidDataException("Invalid naming storage.");
                for (var item = 0; item < names; item++)
                {
                    var name = start + 6 + item * 12; var id = OpenTypeFace.U16(source, name + 6);
                    if (id != 1 && id != 4 && id != 6 && id != 16) continue;
                    var platform = OpenTypeFace.U16(source, name); var encoding = OpenTypeFace.U16(source, name + 2);
                    var size = OpenTypeFace.U16(source, name + 8); var relative = OpenTypeFace.U16(source, name + 10);
                    if (storage + (long)relative + size > length) throw new InvalidDataException("Name string exceeds its table.");
                    decodedBytes += size; if (decodedBytes > 1024 * 1024) throw new NotSupportedException("Host collection exceeds the naming work budget.");
                    var position = start + storage + relative; string value;
                    if (platform == 0 || platform == 3 && (encoding == 0 || encoding == 1 || encoding == 10))
                    {
                        if (size % 2 != 0) throw new InvalidDataException("Invalid UTF16 font name.");
                        try { value = unicode.GetString(source, position, size); }
                        catch (DecoderFallbackException exception) { throw new InvalidDataException("Malformed UTF16 font name.", exception); }
                    }
                    else if (platform == 1 && encoding == 0)
                    {
                        var ascii = true;
                        for (var character = 0; character < size; character++) if (source[position + character] > 127) { ascii = false; break; }
                        if (!ascii) continue; // Never substitute undecodable MacRoman names.
                        value = Encoding.ASCII.GetString(source, position, size);
                    }
                    else continue;
                    if (!string.Equals(value, faceName, StringComparison.OrdinalIgnoreCase)) continue;
                    (id == 4 || id == 6 ? strong : family).Add(index);
                }
            }
        }
        var matches = strong.Count > 0 ? strong : family;
        if (matches.Count != 1) throw new NotSupportedException("Host collection does not identify one requested font face.");
        foreach (var index in matches) return SelectFace(source, index, 64L * 1024 * 1024);
        throw new NotSupportedException("Host collection face was not identified.");
    }

    internal static IReadOnlyList<byte[]> ExtractFaces(byte[] source, long maximumExtractedBytes = 256L * 1024 * 1024)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        OpenTypeFace.Require(source, 0, 12);
        if (OpenTypeFace.U32(source, 0) != 0x74746366) return new[] { source };
        var version = OpenTypeFace.U32(source, 4);
        if (version != 0x00010000 && version != 0x00020000) throw new InvalidDataException("Invalid collection version.");
        var count = OpenTypeFace.Size(OpenTypeFace.U32(source, 8));
        if (count <= 0 || count > 64) throw new InvalidDataException("Invalid collection face count.");
        OpenTypeFace.Require(source, 12, count * 4);
        if (maximumExtractedBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maximumExtractedBytes));
        long total = 0;
        var faces = new List<byte[]>(count);
        for (var index = 0; index < count; index++)
        {
            var face = new OpenTypeFace(source, OpenTypeFace.Size(OpenTypeFace.U32(source, 12 + index * 4)));
            long size = 12 + face.Tables.Count * 16;
            foreach (var table in face.Tables.Values) size += OpenTypeFace.Align(table.Length);
            if (size > maximumExtractedBytes - total) throw new InvalidDataException("Collection faces exceed the extracted font byte budget.");
            var bytes = face.Build(); total += bytes.LongLength; faces.Add(bytes);
        }
        return faces;
    }
}
