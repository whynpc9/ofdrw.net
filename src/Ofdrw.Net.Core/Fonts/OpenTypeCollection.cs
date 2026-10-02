using System;
using System.Collections.Generic;
using System.IO;
namespace Ofdrw.Net.Core.Fonts;

internal static class OpenTypeCollection
{
    internal static byte[] SelectFace(byte[] source, int index)
    {
        OpenTypeFace.Require(source, 0, 12);
        if (OpenTypeFace.U32(source, 0) != 0x74746366)
        {
            if (index != 0) throw new InvalidDataException("Standalone font has only face zero.");
            return source;
        }
        var version = OpenTypeFace.U32(source, 4);
        if (version != 0x00010000 && version != 0x00020000) throw new InvalidDataException("Invalid collection version.");
        var count = OpenTypeFace.Size(OpenTypeFace.U32(source, 8));
        if (count <= 0 || count > 64 || index < 0 || index >= count) throw new InvalidDataException("Invalid collection face index/count.");
        OpenTypeFace.Require(source, 12, count * 4);
        return new OpenTypeFace(source, OpenTypeFace.Size(OpenTypeFace.U32(source, 12 + index * 4))).Build();
    }
    internal static IReadOnlyList<byte[]> ExtractFaces(byte[] source)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        OpenTypeFace.Require(source, 0, 12);
        if (OpenTypeFace.U32(source, 0) != 0x74746366) return new[] { source };
        var version = OpenTypeFace.U32(source, 4);
        if (version != 0x00010000 && version != 0x00020000) throw new InvalidDataException("Invalid collection version.");
        var count = OpenTypeFace.Size(OpenTypeFace.U32(source, 8));
        if (count <= 0 || count > 64) throw new InvalidDataException("Invalid collection face count.");
        OpenTypeFace.Require(source, 12, count * 4);
        var faces = new List<byte[]>(count);
        for (var index = 0; index < count; index++)
            faces.Add(new OpenTypeFace(source, OpenTypeFace.Size(OpenTypeFace.U32(source, 12 + index * 4))).Build());
        return faces;
    }
}
