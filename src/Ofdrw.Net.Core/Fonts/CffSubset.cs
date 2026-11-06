using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using static Ofdrw.Net.Core.Fonts.OpenTypeFace;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>GID-preserving subset of a static OpenType/CFF1 face.</summary>
/// <remarks>Every original GID keeps its slot. CharStrings of unused glyphs become a bare endchar and
/// unused subroutines a bare return, rewritten in place and zero padded, so no CFF offset changes and
/// the OFD's deflate compression removes the padding. The cmap is reduced to the used scalars.</remarks>
internal static class CffSubset
{
    private static readonly HashSet<string> KnownTables = new(StringComparer.Ordinal)
    {
        "head", "hhea", "maxp", "OS/2", "hmtx", "cmap", "name", "post", "CFF ", "GPOS", "GSUB", "GDEF", "BASE",
        "STAT", "vhea", "vmtx", "VORG", "DSIG", "kern", "meta", "JSTF", "MATH"
    };

    internal static byte[] Create(OpenTypeFace face, IEnumerable<int> scalars, CancellationToken cancellationToken,
        long maximumOperations, FontParseBudget parseBudget, out int keptGlyphs, out int originalGlyphs,
        OpenTypeCmap? parsedCmap = null, IEnumerable<int>? actualGlyphIds = null)
    {
        parseBudget.Spend(0);
        foreach (var tag in face.Tables.Keys)
            if (!KnownTables.Contains(tag)) throw new NotSupportedException("Unmodeled font table '" + tag + "' requires full preservation.");
        var outlines = new CffCharStringOutlines(face.Build());
        var program = outlines.Program;
        originalGlyphs = program.GlyphCount;
        var cmap = parsedCmap ?? new OpenTypeCmap(face, parseBudget);
        var mapping = new Dictionary<int, int>(); var keep = new HashSet<int> { 0 };
        foreach (var scalar in scalars.Distinct())
        {
            cancellationToken.ThrowIfCancellationRequested(); var gid = cmap.Glyph(scalar);
            if (gid >= originalGlyphs) throw new InvalidDataException("cmap references a nonexistent glyph.");
            if (gid != 0) { keep.Add(gid); mapping[scalar] = gid; }
        }
        if (actualGlyphIds is not null) foreach (var gid in actualGlyphIds.Distinct())
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (gid <= 0 || gid >= originalGlyphs) throw new InvalidDataException("Shaped glyph reference is outside the original font.");
            keep.Add(gid);
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
        if (face.Tables.TryGetValue("GSUB", out var gsub))
        {
            var previous = -1;
            while (previous != keep.Count)
            {
                previous = keep.Count;
                OpenTypeGlyphClosure.AddSubstitutions(gsub, keep, originalGlyphs, Spend, cancellationToken);
            }
        }
        keptGlyphs = keep.Count;
        var global = new HashSet<int>();
        var local = Enumerable.Range(0, program.FdCount).Select(_ => new HashSet<int>()).ToArray();
        foreach (var gid in keep)
        {
            Spend();
            outlines.CollectSubroutines(gid, global, local[program.FdIndex(gid)], cancellationToken);
        }
        var cff = (byte[])program.Bytes.Clone();
        void Blank(CffIndex index, int item, byte op)
        {
            var start = index.ItemStart(item); var end = index.ItemEnd(item);
            if (end <= start) return;
            cff[start] = op;
            Array.Clear(cff, start + 1, end - start - 1);
        }
        for (var gid = 0; gid < originalGlyphs; gid++)
            if (!keep.Contains(gid)) Blank(program.CharStrings, gid, 14);
        for (var index = 0; index < program.GlobalSubrs.Count; index++)
            if (!global.Contains(index)) Blank(program.GlobalSubrs, index, 11);
        for (var fd = 0; fd < program.FdCount; fd++)
        {
            var subrs = program.LocalSubrsOfFd(fd);
            // FDs may share one Subrs INDEX; blank an entry only when no FD using it calls it.
            var used = new HashSet<int>();
            for (var other = 0; other < program.FdCount; other++)
                if (program.LocalSubrsOfFd(other).Start == subrs.Start) used.UnionWith(local[other]);
            for (var index = 0; index < subrs.Count; index++)
                if (!used.Contains(index)) Blank(subrs, index, 11);
        }
        face.Tables["CFF "] = cff;
        face.Tables["cmap"] = OpenTypeCmap.Build(mapping, cmap.VariationSequences);
        return face.Build();
    }
}
