using System;
using System.Collections.Generic;
using System.IO;
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
