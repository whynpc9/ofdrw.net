using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using static Ofdrw.Net.Core.Fonts.OpenTypeFace;
namespace Ofdrw.Net.Core.Fonts;

internal sealed class OpenTypeCmap
{
    private readonly byte[] table;
    private readonly int format4 = -1;
    private readonly int format12 = -1;
    internal bool IsSymbol { get; }
    internal byte[]? VariationSequences { get; }
    internal HashSet<int> VariationGlyphs { get; } = new();
    // Format 4 scans segments; format 12 uses binary search even at the 32 MiB font limit.
    internal int GlyphLookupWorkUpperBound => format12 >= 0 ? 32 :
        (U16(table, format4 + 6) / 2) * (IsSymbol ? 2 : 1);
    private readonly FontParseBudget budget;
    internal OpenTypeCmap(OpenTypeFace face) : this(face, new FontParseBudget()) { }
    internal OpenTypeCmap(OpenTypeFace face, FontParseBudget budget)
    {
        this.budget = budget ?? throw new ArgumentNullException(nameof(budget));
        budget.Spend(0);
        table = face.Table("cmap", 4);
        var symbol4 = -1;
        var count = U16(table, 2); Require(table, 4, count * 8);
        var parsed = new HashSet<(int Offset, int Format)>();
        var variations = new Dictionary<int, byte[]>();
        for (var i = 0; i < count; i++)
        {
            budget.Spend(1);
            var record = 4 + i * 8; var platform = U16(table, record); var encoding = U16(table, record + 2);
            var offset = Size(U32(table, record + 4)); var format = U16(table, offset);
            if (format == 14)
            {
                if (!variations.TryGetValue(offset, out var variation))
                {
                    var length = Size(U32(table, offset + 2)); Require(table, offset, length);
                    budget.Spend(length);
                    variation = new byte[length]; Buffer.BlockCopy(table, offset, variation, 0, length);
                    ReadVariations(variation);
                    variations.Add(offset, variation);
                }
                VariationSequences = variation;
            }
            if (platform == 3 && encoding == 0 && format == 4) symbol4 = offset;
            if (platform != 0 && !(platform == 3 && (encoding == 1 || encoding == 10))) continue;
            if (format == 4) format4 = offset;
            if (format == 12) { if (parsed.Add((offset, format))) Validate12(offset); format12 = offset; }
        }
        if (format4 < 0 && format12 < 0 && symbol4 >= 0) { format4 = symbol4; IsSymbol = true; }
        // Format 12 is authoritative in GlyphCore. An unused legacy/symbol
        // format 4 cannot invalidate that selected repertoire; validate it only
        // when lookup will actually consume it.
        if (format12 < 0 && format4 >= 0) Validate4(format4);
        if (format4 < 0 && format12 < 0) throw new NotSupportedException("Font needs a Unicode cmap format 4 or 12.");
    }
    private void ReadVariations(byte[] variation)
    {
        Require(variation, 0, 10); var count = Size(U32(variation, 6)); Require(variation, 10, checked(count * 11));
        var arrays = new HashSet<(int Offset, bool NonDefault)>();
        budget.Spend(count);
        for (var i = 0; i < count; i++)
        {
            budget.Spend(0);
            var record = 10 + i * 11; var defaults = Size(U32(variation, record + 3)); var nonDefaults = Size(U32(variation, record + 7));
            if (defaults != 0 && arrays.Add((defaults, false)))
            {
                var number = Size(U32(variation, defaults));
                Require(variation, defaults + 4, checked(number * 4)); budget.Spend(number);
            }
            if (nonDefaults != 0 && arrays.Add((nonDefaults, true)))
            {
                var number = Size(U32(variation, nonDefaults)); Require(variation, nonDefaults + 4, checked(number * 5));
                budget.Spend(number);
                for (var j = 0; j < number; j++)
                { budget.Spend(0); VariationGlyphs.Add(U16(variation, nonDefaults + 4 + j * 5 + 3)); }
            }
        }
    }
    internal static bool IsVariationSelector(int scalar) => scalar is >= 0xFE00 and <= 0xFE0F or >= 0xE0100 and <= 0xE01EF;
    internal bool SupportsVariation(int scalar, int selector)
    {
        budget.Spend(0);
        if (VariationSequences is not { } variation) return false;
        var count = Size(U32(variation, 6));
        for (var i = 0; i < count; i++)
        {
            budget.Spend(1);
            var record = 10 + i * 11;
            if (U24(variation, record) != selector) continue;
            var defaults = Size(U32(variation, record + 3)); var nonDefaults = Size(U32(variation, record + 7));
            if (defaults != 0)
            {
                var number = Size(U32(variation, defaults));
                for (var j = 0; j < number; j++)
                { budget.Spend(1); var p = defaults + 4 + j * 4; var start = U24(variation, p); if (scalar >= start && scalar <= start + variation[p + 3]) return Glyph(scalar) != 0; }
            }
            if (nonDefaults != 0)
            {
                var number = Size(U32(variation, nonDefaults));
                for (var j = 0; j < number; j++)
                { budget.Spend(1); var p = nonDefaults + 4 + j * 5; if (U24(variation, p) == scalar) return U16(variation, p + 3) != 0; }
            }
        }
        return false;
    }
    private static int U24(byte[] bytes, int offset)
    { Require(bytes, offset, 3); return (bytes[offset] << 16) | (bytes[offset + 1] << 8) | bytes[offset + 2]; }
    private void Validate4(int offset)
    {
        Require(table, offset, 16); var length = U16(table, offset + 2); Require(table, offset, length);
        var count = U16(table, offset + 6) / 2;
        if (count == 0 || U16(table, offset + 6) % 2 != 0 || 16L + count * 8 > length)
            throw new InvalidDataException("Invalid cmap format 4 segments.");
        budget.Spend(length + (long)count);
        var previous = -1;
        for (var i = 0; i < count; i++)
        {
            budget.Spend(0);
            var end = U16(table, offset + 14 + i * 2); var start = U16(table, offset + 16 + count * 2 + i * 2);
            if (start > end || start <= previous) throw new InvalidDataException("Unordered cmap segments.");
            previous = end;
        }
    }
    private void Validate12(int offset)
    {
        Require(table, offset, 16); var length = Size(U32(table, offset + 4)); Require(table, offset, length);
        var count = Size(U32(table, offset + 12));
        if (16L + count * 12L > length) throw new InvalidDataException("Invalid cmap format 12 groups.");
        budget.Spend(length + (long)count);
        long previous = -1;
        for (var i = 0; i < count; i++)
        {
            budget.Spend(0);
            var p = offset + 16 + i * 12; var start = U32(table, p); var end = U32(table, p + 4);
            if (start > end || start <= previous || end > 0x10FFFF || U32(table, p + 8) + (ulong)(end - start) > ushort.MaxValue)
                throw new InvalidDataException("Invalid Unicode cmap group.");
            previous = end;
        }
    }
    internal int Glyph(int scalar)
    {
        budget.Spend(0);
        var glyph = GlyphCore(scalar);
        return IsSymbol && glyph == 0 && scalar <= 0xFF ? GlyphCore(scalar + 0xF000) : glyph;
    }
    // Symbol-font byte fallback is useful for inspecting legacy files, but it
    // must not be mistaken for a direct Unicode mapping by an export outlet.
    internal int DirectGlyph(int scalar)
    {
        budget.Spend(0);
        return GlyphCore(scalar);
    }
    // PDFsharp may give two Unicode scalars that share one font GID the same
    // PDF code. A path-painted CGTransform span uses this face only for its
    // invisible semantic text, so that GID must have one Unicode meaning in
    // the selected cmap even if the other scalar is used by another TextObject.
    internal bool HasDistinctUnicodeScalarForGlyph(int gid, int sourceScalar)
    {
        budget.Spend(0);
        if (gid <= 0) return false;
        if (format12 >= 0)
        {
            var count = Size(U32(table, format12 + 12));
            for (var index = 0; index < count; index++)
            {
                budget.Spend(1);
                var group = format12 + 16 + index * 12;
                var firstScalar = U32(table, group);
                var lastScalar = U32(table, group + 4);
                var firstGlyph = U32(table, group + 8);
                if ((ulong)gid < firstGlyph || (ulong)gid > firstGlyph + lastScalar - firstScalar) continue;
                var scalar = firstScalar + (uint)((ulong)gid - firstGlyph);
                if (scalar != sourceScalar) return true;
            }
            return false;
        }
        // A delta segment has at most one scalar for this GID. Only explicit
        // glyph-array segments need a per-position scan; their disjoint spans
        // contain at most 65,536 positions in total.
        var count4 = U16(table, format4 + 6) / 2;
        for (var index = 0; index < count4; index++)
        {
            budget.Spend(1);
            var start = U16(table, format4 + 16 + count4 * 2 + index * 2);
            var end = U16(table, format4 + 14 + index * 2);
            var delta = U16(table, format4 + 16 + count4 * 4 + index * 2);
            var address = format4 + 16 + count4 * 6 + index * 2;
            var range = U16(table, address);
            if (range == 0)
            {
                var scalar = (gid - delta) & 0xFFFF;
                if (scalar >= start && scalar <= end && scalar != sourceScalar &&
                    !char.IsSurrogate((char)scalar)) return true;
                continue;
            }
            for (var scalar = start; scalar <= end; scalar++)
            {
                budget.Spend(1);
                var target = address + range + (scalar - start) * 2;
                if (target + 2 > format4 + U16(table, format4 + 2))
                    throw new InvalidDataException("cmap glyph reference escapes subtable.");
                var raw = U16(table, target);
                if (raw != 0 && ((raw + delta) & 0xFFFF) == gid && scalar != sourceScalar &&
                    !char.IsSurrogate((char)scalar)) return true;
            }
        }
        return false;
    }
    private int GlyphCore(int scalar)
    {
        if (format12 >= 0)
        {
            var count = Size(U32(table, format12 + 12)); var lo = 0; var hi = count - 1;
            while (lo <= hi)
            {
                var mid = lo + (hi - lo) / 2; var p = format12 + 16 + mid * 12;
                var first = U32(table, p); var last = U32(table, p + 4);
                if (scalar < first) hi = mid - 1;
                else if (scalar > last) lo = mid + 1;
                else return checked((int)(U32(table, p + 8) + scalar - first));
            }
            // A format 12 Unicode repertoire is authoritative, even when a
            // legacy BMP subtable disagrees; every consumer uses this rule.
            return 0;
        }
        if (scalar > 0xFFFF) return 0;
        var count4 = U16(table, format4 + 6) / 2;
        for (var i = 0; i < count4; i++)
        {
            if (scalar > U16(table, format4 + 14 + i * 2)) continue;
            if (scalar < U16(table, format4 + 16 + count4 * 2 + i * 2)) return 0;
            var delta = U16(table, format4 + 16 + count4 * 4 + i * 2);
            var address = format4 + 16 + count4 * 6 + i * 2;
            var range = U16(table, address);
            if (range == 0) return (scalar + delta) & 0xFFFF;
            var target = address + range + (scalar - U16(table, format4 + 16 + count4 * 2 + i * 2)) * 2;
            if (target + 2 > format4 + U16(table, format4 + 2)) throw new InvalidDataException("cmap glyph reference escapes subtable.");
            var glyph = U16(table, target);
            return glyph == 0 ? 0 : (glyph + delta) & 0xFFFF;
        }
        return 0;
    }
    internal static byte[] Build(IReadOnlyDictionary<int, int> mapping, byte[]? variations = null)
    {
        var pairs = mapping.Where(pair => pair.Value != 0).OrderBy(pair => pair.Key).ToArray();
        var bmp = pairs.Where(pair => pair.Key < 0xFFFF).ToArray();
        var segments = new List<(int Start, int End, int Delta)>();
        foreach (var pair in bmp)
        {
            var delta = (pair.Value - pair.Key) & 0xFFFF;
            if (segments.Count > 0 && segments[segments.Count - 1].End + 1 == pair.Key && segments[segments.Count - 1].Delta == delta)
            { var previous = segments[segments.Count - 1]; segments[segments.Count - 1] = (previous.Start, pair.Key, delta); }
            else segments.Add((pair.Key, pair.Key, delta));
        }
        // PDFsharp requires a BMP subtable even for non-BMP-only text. Never
        // produce format-12-only subsets its current parser cannot consume.
        if (segments.Count > 8188) throw new NotSupportedException("Unicode cmap format 4 cannot fit; full font retained for PDF compatibility.");
        segments.Add((0xFFFF, 0xFFFF, 1));
        const bool has4 = true;
        var length4 = 16 + segments.Count * 8;
        var length12 = checked(16 + pairs.Length * 12);
        var records = (has4 ? 3 : 2) + (variations is null ? 0 : 1); var header = 4 + records * 8;
        var result = new byte[checked(header + length4 + length12 + (variations?.Length ?? 0))];
        Put16(result, 2, records);
        void Record(int index, int platform, int encoding, int offset)
        { Put16(result, 4 + index * 8, platform); Put16(result, 6 + index * 8, encoding); Put32(result, 8 + index * 8, (uint)offset); }
        if (has4)
        {
            Record(0, 3, 1, header); var count = segments.Count; var power = 1; var selector = 0;
            while (power * 2 <= count) { power *= 2; selector++; }
            Put16(result, header, 4); Put16(result, header + 2, length4); Put16(result, header + 6, count * 2);
            Put16(result, header + 8, power * 2); Put16(result, header + 10, selector); Put16(result, header + 12, count * 2 - power * 2);
            for (var i = 0; i < count; i++)
            {
                var segment = segments[i];
                Put16(result, header + 14 + i * 2, segment.End); Put16(result, header + 16 + count * 2 + i * 2, segment.Start);
                Put16(result, header + 16 + count * 4 + i * 2, segment.Delta);
            }
        }
        var p12 = header + length4;
        Record(has4 ? 1 : 0, 0, 4, p12); Record(has4 ? 2 : 1, 3, 10, p12);
        Put16(result, p12, 12); Put32(result, p12 + 4, (uint)length12); Put32(result, p12 + 12, (uint)pairs.Length);
        for (var i = 0; i < pairs.Length; i++)
        {
            var p = p12 + 16 + i * 12; Put32(result, p, (uint)pairs[i].Key); Put32(result, p + 4, (uint)pairs[i].Key); Put32(result, p + 8, (uint)pairs[i].Value);
        }
        if (variations is not null)
        { Record(records - 1, 0, 5, p12 + length12); variations.CopyTo(result, p12 + length12); }
        return result;
    }
}
