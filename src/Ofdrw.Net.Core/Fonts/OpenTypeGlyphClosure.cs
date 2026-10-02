using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using static Ofdrw.Net.Core.Fonts.OpenTypeFace;
namespace Ofdrw.Net.Core.Fonts;

internal static class OpenTypeGlyphClosure
{
    // Preserve GIDs and the GSUB table. All lookups are considered regardless
    // of script/feature enablement: contextual references cannot escape this
    // conservative fixed point. GPOS changes positions, never glyph identity.
    internal static void AddSubstitutions(byte[] table, HashSet<int> glyphs, int glyphCount, Action spend, CancellationToken cancellationToken)
    {
        Require(table, 0, 10);
        var version = U32(table, 0);
        if (version != 0x00010000) throw new NotSupportedException("GSUB feature variations are preserved without subsetting.");
        var list = U16(table, 8); var count = U16(table, list); Require(table, list + 2, count * 2);
        var changed = true;
        while (changed)
        {
            spend(); cancellationToken.ThrowIfCancellationRequested(); var previous = glyphs.Count;
            for (var i = 0; i < count; i++)
            {
                spend(); cancellationToken.ThrowIfCancellationRequested();
                var lookup = list + U16(table, list + 2 + i * 2); var type = U16(table, lookup); var subCount = U16(table, lookup + 4);
                Require(table, lookup + 6, subCount * 2);
                for (var j = 0; j < subCount; j++) Visit(type, lookup + U16(table, lookup + 6 + j * 2), 0);
            }
            changed = glyphs.Count != previous;
        }
        void Add(int gid)
        {
            spend();
            if (gid >= glyphCount) throw new InvalidDataException("GSUB references a nonexistent glyph.");
            glyphs.Add(gid);
        }
        void Visit(int type, int offset, int depth)
        {
            if (depth > 1) throw new InvalidDataException("Recursive GSUB extension.");
            spend();
            var format = U16(table, offset);
            if (type == 7)
            {
                if (format != 1) throw new NotSupportedException("Unknown GSUB extension format.");
                var nextType = U16(table, offset + 2);
                if (nextType == 7) throw new InvalidDataException("Recursive GSUB extension.");
                Visit(nextType, checked(offset + Size(U32(table, offset + 4))), depth + 1); return;
            }
            if (type == 5 || type == 6)
            {
                if (format < 1 || format > 3) throw new NotSupportedException("Unknown contextual GSUB format.");
                // Referenced lookup outputs are handled by the outer loop. No
                // contextual rule itself introduces a glyph.
                return;
            }
            if (type < 1 || type > 8) throw new NotSupportedException("Unknown GSUB lookup type.");
            var coverage = Coverage(offset + U16(table, offset + 2));
            if (type == 1 && format == 1)
            {
                var delta = U16(table, offset + 4);
                foreach (var gid in coverage) { spend(); if (glyphs.Contains(gid)) Add((gid + delta) & 0xFFFF); }
            }
            else if (type == 1 && format == 2)
            {
                var number = U16(table, offset + 4);
                if (number != coverage.Count) throw new InvalidDataException("GSUB coverage count mismatch.");
                Require(table, offset + 6, number * 2);
                for (var i = 0; i < number; i++) { spend(); if (glyphs.Contains(coverage[i])) Add(U16(table, offset + 6 + i * 2)); }
            }
            else if ((type == 2 || type == 3) && format == 1)
            {
                var number = U16(table, offset + 4);
                if (number != coverage.Count) throw new InvalidDataException("GSUB coverage count mismatch.");
                Require(table, offset + 6, number * 2);
                for (var i = 0; i < number; i++)
                {
                    spend();
                    if (!glyphs.Contains(coverage[i])) continue;
                    var sequence = offset + U16(table, offset + 6 + i * 2); var length = U16(table, sequence);
                    Require(table, sequence + 2, length * 2);
                    for (var j = 0; j < length; j++) Add(U16(table, sequence + 2 + j * 2));
                }
            }
            else if (type == 4 && format == 1)
            {
                var number = U16(table, offset + 4);
                if (number != coverage.Count) throw new InvalidDataException("GSUB coverage count mismatch.");
                Require(table, offset + 6, number * 2);
                for (var i = 0; i < number; i++)
                {
                    spend();
                    if (!glyphs.Contains(coverage[i])) continue;
                    var set = offset + U16(table, offset + 6 + i * 2); var length = U16(table, set); Require(table, set + 2, length * 2);
                    for (var j = 0; j < length; j++)
                    {
                        spend();
                        var liga = set + U16(table, set + 2 + j * 2); var components = U16(table, liga + 2);
                        if (components == 0) throw new InvalidDataException("Invalid GSUB ligature.");
                        Require(table, liga + 4, (components - 1) * 2); var match = true;
                        for (var k = 0; k < components - 1; k++) { spend(); if (!glyphs.Contains(U16(table, liga + 4 + k * 2))) match = false; }
                        if (match) Add(U16(table, liga));
                    }
                }
            }
            else if (type == 8 && format == 1)
            {
                var cursor = offset + 4;
                for (var group = 0; group < 2; group++)
                { var length = U16(table, cursor); Require(table, cursor + 2, length * 2); cursor += 2 + length * 2; }
                var number = U16(table, cursor);
                if (number != coverage.Count) throw new InvalidDataException("GSUB coverage count mismatch.");
                Require(table, cursor + 2, number * 2);
                for (var i = 0; i < number; i++) { spend(); if (glyphs.Contains(coverage[i])) Add(U16(table, cursor + 2 + i * 2)); }
            }
            else throw new NotSupportedException("Unknown GSUB substitution format.");
        }
        List<int> Coverage(int offset)
        {
            var result = new List<int>(); var format = U16(table, offset); var number = U16(table, offset + 2);
            if (format == 1)
            {
                Require(table, offset + 4, number * 2);
                for (var i = 0; i < number; i++) { spend(); result.Add(U16(table, offset + 4 + i * 2)); }
            }
            else if (format == 2)
            {
                Require(table, offset + 4, number * 6);
                for (var i = 0; i < number; i++)
                {
                    spend();
                    var record = offset + 4 + i * 6; var start = U16(table, record); var end = U16(table, record + 2);
                    if (start > end || U16(table, record + 4) != result.Count) throw new InvalidDataException("Invalid GSUB coverage range.");
                    for (var gid = start; gid <= end; gid++) { spend(); result.Add(gid); }
                    if (result.Count > glyphCount) throw new InvalidDataException("GSUB coverage exceeds glyph count.");
                }
            }
            else throw new NotSupportedException("Unknown GSUB coverage format.");
            foreach (var gid in result) if (gid >= glyphCount) throw new InvalidDataException("GSUB coverage references a nonexistent glyph.");
            return result;
        }
    }
}
