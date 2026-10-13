using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace Ofdrw.Net.Converter.Pdf.Geometry;

// B profile metadata only, AFTER exact trusted identity admission. This is not an
// untrusted-font validator for SixLabors' eager GSUB/GPOS/GDEF/all-glyph parsing.
internal sealed class TrustedFontMetadata
{
    internal readonly struct Glyph
    {
        internal Glyph(short contours, ushort points, short xmin, short ymin, short xmax, short ymax, ushort advance, bool empty)
        { Contours = contours; Points = points; XMin = xmin; YMin = ymin; XMax = xmax; YMax = ymax; Advance = advance; Empty = empty; }
        internal readonly short Contours, XMin, YMin, XMax, YMax;
        internal readonly ushort Points, Advance;
        internal readonly bool Empty;
    }
    private readonly struct Group
    {
        internal Group(uint start, uint end, uint gid) { Start = start; End = end; Gid = gid; }
        internal readonly uint Start, End, Gid;
    }
    private readonly Glyph[] glyphs;
    private readonly Group[]? groups;
    private readonly byte[]? cmap4;
    internal ushort UnitsPerEm { get; }
    internal long SimplePoints { get; }
    internal int GlyphCount => glyphs.Length;
    internal Glyph this[int glyph] => glyphs[glyph];

    private TrustedFontMetadata(Glyph[] glyphs, ushort upem, long points, Group[]? groups, byte[]? cmap4)
    { this.glyphs = glyphs; UnitsPerEm = upem; SimplePoints = points; this.groups = groups; this.cmap4 = cmap4; }

    internal static TrustedFontMetadata Read(Stream input, Action<int> chargeGlyphRecords, CancellationToken token)
    {
        var read = new Reader(input);
        if (read.U32(0) != 0x00010000) throw new NotSupportedException("Only trusted standalone TrueType SFNT is supported.");
        var count = read.U16(4);
        if (count == 0 || count > 64) throw new NotSupportedException("Unsupported trusted font table count.");
        var tables = new Dictionary<string, (long Offset, int Length)>(StringComparer.Ordinal);
        for (var i = 0; i < count; i++)
        {
            token.ThrowIfCancellationRequested(); var at = 12 + i * 16;
            var tag = Encoding.ASCII.GetString(read.Bytes(at, 4));
            var offset = read.U32(at + 8); var length = read.U32(at + 12);
            if (length > int.MaxValue || offset > input.Length || length > input.Length - offset || tables.ContainsKey(tag))
                throw new InvalidDataException("Invalid trusted font table directory.");
            tables.Add(tag, (offset, (int)length));
        }
        foreach (var tag in new[] { "CFF ", "CFF2", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR", "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "sbix" })
            if (tables.ContainsKey(tag)) throw new NotSupportedException("Variable, color and non-TrueType geometry are unsupported.");
        (long Offset, int Length) Table(string tag, int minimum)
        {
            if (!tables.TryGetValue(tag, out var t) || t.Length < minimum) throw new InvalidDataException("Missing or short trusted font table: " + tag);
            return t;
        }
        var head = Table("head", 54); var maxp = Table("maxp", 6); var hhea = Table("hhea", 36);
        var hmtx = Table("hmtx", 4); var loca = Table("loca", 2); var glyf = Table("glyf", 0); var cmap = Table("cmap", 4);
        var os2 = Table("OS/2", 64);
        var os2Version = read.U16(os2.Offset);
        if (os2Version > 5) throw new NotSupportedException("Unsupported trusted OS/2 metadata version.");
        var macStyle = read.U16(head.Offset + 44); var selection = read.U16(os2.Offset + 62);
        // OBLIQUE (bit9, defined in v4+) is independent of ITALIC. Older reserved uses also fail closed.
        if ((macStyle & 3) != 0 || (selection & (1 | 32 | 512)) != 0 || read.U16(os2.Offset + 4) != 400)
            throw new NotSupportedException("Only actual Regular weight-400 faces are supported.");
        var upem = read.U16(head.Offset + 18); var format = read.I16(head.Offset + 50);
        if (upem < 16 || upem > 16384 || format is not (0 or 1)) throw new InvalidDataException("Invalid trusted font units or loca format.");
        var glyphCount = read.U16(maxp.Offset + 4); var metrics = read.U16(hhea.Offset + 34);
        if (glyphCount == 0 || metrics == 0 || metrics > glyphCount) throw new InvalidDataException("Invalid trusted glyph/metric counts.");
        chargeGlyphRecords(glyphCount); // before arrays, across all bound faces.
        var stride = format == 0 ? 2 : 4;
        if ((long)(glyphCount + 1) * stride > loca.Length || (long)metrics * 4 + (glyphCount - metrics) * 2 > hmtx.Length)
            throw new InvalidDataException("Short trusted glyph index or metrics table.");
        var offsets = new int[glyphCount + 1];
        for (var i = 0; i < offsets.Length; i++)
        {
            if ((i & 255) == 0) token.ThrowIfCancellationRequested();
            var offset = format == 0 ? (uint)read.U16(loca.Offset + i * 2L) * 2 : read.U32(loca.Offset + i * 4L);
            if (offset > glyf.Length || i > 0 && offset < offsets[i - 1]) throw new InvalidDataException("Invalid trusted glyph range.");
            offsets[i] = (int)offset;
        }
        var glyphs = new Glyph[glyphCount]; long totalPoints = 0; ushort advance = 0;
        for (var i = 0; i < glyphCount; i++)
        {
            if ((i & 255) == 0) token.ThrowIfCancellationRequested();
            if (i < metrics) advance = read.U16(hmtx.Offset + i * 4L);
            var length = offsets[i + 1] - offsets[i];
            if (length == 0) { glyphs[i] = new Glyph(0, 0, 0, 0, 0, 0, advance, true); continue; }
            if (length < 10) throw new InvalidDataException("Short trusted glyph header.");
            var at = glyf.Offset + offsets[i]; var contours = read.I16(at);
            var xmin = read.I16(at + 2); var ymin = read.I16(at + 4); var xmax = read.I16(at + 6); var ymax = read.I16(at + 8);
            if (xmin > xmax || ymin > ymax) throw new InvalidDataException("Invalid trusted glyph box.");
            var points = 0;
            if (contours > 0)
            {
                if (10L + contours * 2L > length) throw new InvalidDataException("Short trusted glyph endpoints.");
                points = read.U16(at + 10 + (contours - 1) * 2L) + 1;
                totalPoints += points;
                if (points > ushort.MaxValue || totalPoints > 8_000_000) throw new InvalidDataException("Trusted face point admission exceeded.");
            }
            glyphs[i] = new Glyph(contours, (ushort)points, xmin, ymin, xmax, ymax, advance, contours == 0);
        }
        var subtables = read.U16(cmap.Offset + 2);
        if (subtables == 0 || subtables > 32 || 4L + 8L * subtables > cmap.Length) throw new NotSupportedException("Unsupported trusted cmap directory.");
        long selected12 = -1, selected4 = -1;
        for (var i = 0; i < subtables; i++)
        {
            token.ThrowIfCancellationRequested(); var at = cmap.Offset + 4 + i * 8L;
            var platform = read.U16(at); var encoding = read.U16(at + 2); var relative = read.U32(at + 4);
            if (relative > cmap.Length - 2) throw new InvalidDataException("Invalid trusted cmap offset.");
            var start = cmap.Offset + relative; var f = read.U16(start);
            if (platform == 0 || platform == 3 && encoding is 1 or 10)
            { if (f == 12) selected12 = start; else if (f == 4) selected4 = start; }
        }
        Group[]? groups = null; byte[]? bytes4 = null;
        if (selected12 >= 0)
        {
            var length = read.U32(selected12 + 4); var n = read.U32(selected12 + 12);
            if (length > 1024 * 1024 || n > 65_536 || length < 16L + n * 12L || length > cmap.Offset + cmap.Length - selected12)
                throw new InvalidDataException("Trusted cmap12 capacity or range exceeded.");
            groups = new Group[n]; uint previous = 0;
            for (var i = 0; i < groups.Length; i++)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested(); var at = selected12 + 16 + i * 12L;
                var start = read.U32(at); var end = read.U32(at + 4); var gid = read.U32(at + 8);
                if (start > end || end > 0x10ffff || i > 0 && start <= previous || gid + (ulong)end - start >= (ulong)glyphCount)
                    throw new InvalidDataException("Invalid trusted cmap12 group.");
                groups[i] = new Group(start, end, gid); previous = end;
            }
        }
        else if (selected4 >= 0)
        {
            var length = read.U16(selected4 + 2);
            if (length < 16 || length > cmap.Offset + cmap.Length - selected4) throw new InvalidDataException("Invalid trusted cmap4 length.");
            bytes4 = read.Bytes(selected4, length); var segments = U16(bytes4, 6) / 2;
            if (U16(bytes4, 6) % 2 != 0 || segments == 0 || segments > 8192 || 16L + segments * 8L > length) throw new InvalidDataException("Invalid trusted cmap4 segments.");
            var ends = 14; var starts = ends + 2 * segments + 2; var previousEnd = -1;
            for (var i = 0; i < segments; i++)
            {
                if ((i & 255) == 0) token.ThrowIfCancellationRequested();
                var first = U16(bytes4, starts + i * 2); var last = U16(bytes4, ends + i * 2);
                if (first > last || first <= previousEnd) throw new InvalidDataException("Unordered trusted cmap4 segments.");
                previousEnd = last;
            }
        }
        else throw new NotSupportedException("Trusted font has no supported Unicode cmap4/12 mapping.");
        token.ThrowIfCancellationRequested(); return new TrustedFontMetadata(glyphs, upem, totalPoints, groups, bytes4);
    }

    internal int Map(int scalar)
    {
        if (groups is not null)
        {
            var lo = 0; var hi = groups.Length - 1;
            while (lo <= hi)
            {
                var mid = lo + (hi - lo) / 2; var group = groups[mid];
                if ((uint)scalar < group.Start) hi = mid - 1;
                else if ((uint)scalar > group.End) lo = mid + 1;
                else return (int)(group.Gid + (uint)scalar - group.Start);
            }
            return 0;
        }
        if (scalar > 0xffff) return 0;
        var bytes = cmap4!; var n = U16(bytes, 6) / 2; var end = 14; var start = end + 2 * n + 2; var delta = start + 2 * n; var range = delta + 2 * n;
        var lo4 = 0; var hi4 = n - 1;
        while (lo4 <= hi4)
        {
            var mid = lo4 + (hi4 - lo4) / 2;
            if (scalar > U16(bytes, end + mid * 2)) lo4 = mid + 1; else hi4 = mid - 1;
        }
        if (lo4 == n || scalar < U16(bytes, start + lo4 * 2)) return 0;
        var d = unchecked((short)U16(bytes, delta + lo4 * 2)); var r = U16(bytes, range + lo4 * 2);
        if (r == 0) return (scalar + d) & 0xffff;
        var gid = U16(bytes, range + lo4 * 2 + r + 2 * (scalar - U16(bytes, start + lo4 * 2)));
        return gid == 0 ? 0 : (gid + d) & 0xffff;
    }
    private static ushort U16(byte[] bytes, int at)
    {
        if (at < 0 || at > bytes.Length - 2) throw new InvalidDataException("Trusted cmap access is out of range.");
        return (ushort)((bytes[at] << 8) | bytes[at + 1]);
    }
    private sealed class Reader(Stream input)
    {
        internal ushort U16(long at) { Position(at, 2); return (ushort)((Byte() << 8) | Byte()); }
        internal short I16(long at) => unchecked((short)U16(at));
        internal uint U32(long at) { Position(at, 4); return ((uint)Byte() << 24) | ((uint)Byte() << 16) | ((uint)Byte() << 8) | (uint)Byte(); }
        internal byte[] Bytes(long at, int length)
        {
            Position(at, length); var result = new byte[length]; var done = 0;
            while (done < length) { var n = input.Read(result, done, length - done); if (n == 0) throw new EndOfStreamException(); done += n; }
            return result;
        }
        private void Position(long at, int length)
        { if (at < 0 || at > input.Length - length) throw new InvalidDataException("Trusted metadata access is out of range."); input.Position = at; }
        private int Byte() { var value = input.ReadByte(); if (value < 0) throw new EndOfStreamException(); return value; }
    }
}
