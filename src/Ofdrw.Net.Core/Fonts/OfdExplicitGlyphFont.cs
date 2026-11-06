using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Fonts;

// A selected, embeddable source face and its exact original-GID outlines.
internal sealed class OfdExplicitGlyphFont
{
    private OfdExplicitGlyphFont(byte[] faceBytes, TrueTypeGlyphOutline[] outlines)
    { FaceBytes = faceBytes; Outlines = outlines; }

    internal OfdExplicitGlyphFont Slice(int offset, int count) =>
        new(FaceBytes, Outlines.Skip(offset).Take(count).ToArray());

    internal byte[] FaceBytes { get; }
    internal IReadOnlyList<TrueTypeGlyphOutline> Outlines { get; }

    internal static OfdExplicitGlyphFont Read(OfdFontResource? resource, OfdTextElement text,
        OfdExplicitGlyphRun run, bool embedFont)
    {
        if (resource?.Data is not { Length: > 0 })
            throw new NotSupportedException("CGTransform rendering requires an explicitly embedded font.");
        if (text.Weight >= 600 && !resource.Bold || text.Italic && !resource.Italic)
            throw new NotSupportedException("CGTransform synthetic bold or italic is outside the explicit outline profile.");
        var outlines = GlyphOutlineSources.Open(resource.Data, resource.CollectionFaceIndex);
        var faceBytes = OpenTypeCollection.SelectFace(resource.Data, resource.CollectionFaceIndex);
        if (embedFont)
        {
            var face = new OpenTypeFace(faceBytes);
            if (!face.Tables.TryGetValue("OS/2", out var os2) || os2.Length < 10 ||
                (OpenTypeFace.U16(os2, 8) & (0x0002 | 0x0100 | 0x0200)) != 0)
                throw new NotSupportedException("CGTransform font does not permit the verified embedding profile.");
        }
        var values = new TrueTypeGlyphOutline[run.GlyphIds.Count];
        var cache = new Dictionary<int, TrueTypeGlyphOutline>();
        long emittedPoints = 0;
        for (var index = 0; index < values.Length; index++)
        {
            var gid = run.GlyphIds[index];
            if (!cache.TryGetValue(gid, out var outline))
                cache.Add(gid, outline = outlines.Read(gid));
            emittedPoints += outline.Contours.Sum(contour => (long)contour.Points.Count);
            if (emittedPoints > 131072)
                throw new NotSupportedException("CGTransform emitted outline points exceed the bounded render profile.");
            values[index] = outline;
        }
        // A visible character needs ink; in a complete cluster map, a cluster's hidden joiner or selector
        // glyph is blank as long as another glyph of the same cluster has an outline.
        for (var index = 0; index < values.Length; index++)
        {
            if (values[index].Contours.Count != 0 || run.IsWhitespaceSlot(index)) continue;
            if (run.ClusterOfSlot(index) is { } cluster &&
                Enumerable.Range(cluster.First, cluster.Count).Any(slot => values[slot].Contours.Count != 0))
                continue;
            throw new NotSupportedException("CGTransform mapped a visible character to an empty original GID.");
        }
        return new OfdExplicitGlyphFont(faceBytes, values);
    }
}
