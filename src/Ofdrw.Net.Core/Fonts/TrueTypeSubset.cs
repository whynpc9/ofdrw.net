using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using static Ofdrw.Net.Core.Fonts.OpenTypeFace;
namespace Ofdrw.Net.Core.Fonts;

internal static class TrueTypeSubset
{
    private static readonly HashSet<string> KnownTables = new(StringComparer.Ordinal)
    {
        "head", "hhea", "maxp", "OS/2", "hmtx", "cmap", "loca", "glyf", "name", "post",
        "cvt ", "fpgm", "prep", "gasp", "kern", "GPOS", "GSUB", "GDEF", "BASE",
        "STAT", "vhea", "vmtx", "VORG", "VDMX", "hdmx", "LTSH", "PCLT", "DSIG"
    };
    internal static byte[] Create(OpenTypeFace face, IEnumerable<int> scalars, CancellationToken cancellationToken, long maximumOperations,
        out int keptGlyphs, out int originalGlyphs)
    {
        foreach (var tag in face.Tables.Keys)
            if (!KnownTables.Contains(tag)) throw new NotSupportedException("Unmodeled font table '" + tag + "' requires full preservation.");
        var head = face.Table("head", 54); var maxp = face.Table("maxp", 6); var glyf = face.Table("glyf"); var loca = face.Table("loca");
        originalGlyphs = U16(maxp, 4);
        if (originalGlyphs == 0) throw new InvalidDataException("Font has no glyphs.");
        var hhea = face.Table("hhea", 36); var metrics = U16(hhea, 34);
        if (metrics == 0 || metrics > originalGlyphs) throw new InvalidDataException("Invalid horizontal metrics count.");
        Require(face.Table("hmtx"), 0, metrics * 4 + (originalGlyphs - metrics) * 2);
        var format = U16(head, 50);
        if (format != 0 && format != 1) throw new InvalidDataException("Invalid loca format.");
        Require(loca, 0, (originalGlyphs + 1) * (format == 0 ? 2 : 4));
        var offsets = new int[originalGlyphs + 1];
        for (var i = 0; i <= originalGlyphs; i++)
        {
            offsets[i] = format == 0 ? U16(loca, i * 2) * 2 : Size(U32(loca, i * 4));
            if (offsets[i] > glyf.Length || (i > 0 && offsets[i] < offsets[i - 1])) throw new InvalidDataException("Invalid glyph offset.");
        }
        var cmap = new OpenTypeCmap(face);
        var mapping = new Dictionary<int, int>(); var keep = new HashSet<int> { 0 };
        foreach (var scalar in scalars.Distinct())
        {
            cancellationToken.ThrowIfCancellationRequested(); var gid = cmap.Glyph(scalar);
            if (gid >= originalGlyphs) throw new InvalidDataException("cmap references a nonexistent glyph.");
            if (gid != 0) { keep.Add(gid); mapping[scalar] = gid; }
        }
        foreach (var gid in cmap.VariationGlyphs)
        {
            if (gid >= originalGlyphs) throw new InvalidDataException("Variation references a nonexistent glyph.");
            keep.Add(gid);
        }
        long operations = 0;
        void Spend()
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (++operations > maximumOperations) throw new InvalidDataException("Font glyph closure exceeds MaximumGlyphClosureOperations.");
        }
        var previous = -1;
        while (previous != keep.Count)
        {
            previous = keep.Count;
            if (face.Tables.TryGetValue("GSUB", out var gsub)) OpenTypeGlyphClosure.AddSubstitutions(gsub, keep, originalGlyphs, Spend, cancellationToken);
            var visiting = new HashSet<int>(); var done = new HashSet<int>();
            foreach (var gid in keep.ToArray()) Composite(gid, 0);
            void Composite(int gid, int depth)
            {
                Spend();
                if (gid >= offsets.Length - 1) throw new InvalidDataException("Composite references a nonexistent glyph.");
                if (done.Contains(gid)) return;
                if (depth > 64 || !visiting.Add(gid)) throw new InvalidDataException("Recursive or excessively deep composite glyph.");
                var start = offsets[gid]; var length = offsets[gid + 1] - start;
                if (length > 0)
                {
                    if (length < 10) throw new InvalidDataException("Truncated glyph header.");
                    if ((short)U16(glyf, start) < 0)
                    {
                        var cursor = start + 10; int flags;
                        do
                        {
                            if (cursor + 4 > start + length) throw new InvalidDataException("Truncated composite glyph.");
                            flags = U16(glyf, cursor); var child = U16(glyf, cursor + 2); keep.Add(child); Composite(child, depth + 1);
                            cursor += 4 + ((flags & 1) != 0 ? 4 : 2);
                            if ((flags & 8) != 0) cursor += 2;
                            else if ((flags & 64) != 0) cursor += 4;
                            else if ((flags & 128) != 0) cursor += 8;
                            if (cursor > start + length) throw new InvalidDataException("Truncated composite transform.");
                        } while ((flags & 32) != 0);
                        if ((flags & 256) != 0)
                        {
                            if (cursor + 2 > start + length) throw new InvalidDataException("Truncated composite instructions.");
                            var instructions = U16(glyf, cursor);
                            if (cursor + 2 + instructions > start + length) throw new InvalidDataException("Truncated composite instructions.");
                        }
                    }
                }
                visiting.Remove(gid); done.Add(gid);
            }
        }
        keptGlyphs = keep.Count;
        using var output = new MemoryStream(); var newLoca = new byte[(originalGlyphs + 1) * 4];
        for (var gid = 0; gid < originalGlyphs; gid++)
        {
            cancellationToken.ThrowIfCancellationRequested(); Put32(newLoca, gid * 4, (uint)output.Length);
            if (!keep.Contains(gid)) continue;
            var length = offsets[gid + 1] - offsets[gid]; output.Write(glyf, offsets[gid], length);
            while (output.Length % 4 != 0) output.WriteByte(0);
        }
        Put32(newLoca, originalGlyphs * 4, (uint)output.Length);
        face.Tables["glyf"] = output.ToArray(); face.Tables["loca"] = newLoca; Put16(head, 50, 1);
        face.Tables["cmap"] = OpenTypeCmap.Build(mapping, cmap.VariationSequences);
        return face.Build();
    }
}
