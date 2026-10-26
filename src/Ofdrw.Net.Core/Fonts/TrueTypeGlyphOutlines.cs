using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Fonts;

// An operation-owned, unhinted design-unit outline. The caller supplies an
// original GID; no Unicode cmap or shaping decision is made here. R02 can use
// the same points for PDF and SVG while retaining TextCode separately.
internal sealed class TrueTypeGlyphOutlines
{
    private const int MaximumFaceBytes = 16 * 1024 * 1024;
    private const int MaximumSourceBytes = 64 * 1024 * 1024;
    private const int MaximumPoints = 32768;
    private const int MaximumComponents = 4096;
    private const int MaximumDepth = 16;
    private static readonly string[] UnsupportedTables =
    [
        "CFF ", "CFF2", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
        "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "EBDT", "EBLC", "sbix"
    ];
    private readonly byte[] glyf;
    private readonly byte[] hmtx;
    private readonly int[] offsets;
    private readonly int metricCount;

    internal TrueTypeGlyphOutlines(byte[] sourceBytes, int faceIndex = 0)
    {
        if (sourceBytes is null) throw new ArgumentNullException(nameof(sourceBytes));
        if (sourceBytes.Length is < 12 or > MaximumSourceBytes)
            throw new NotSupportedException("GID outlines require a bounded TrueType source.");
        var ownedSource = (byte[])sourceBytes.Clone();
        var selected = OpenTypeCollection.SelectFace(ownedSource, faceIndex, MaximumFaceBytes);
        var face = new OpenTypeFace(selected);
        if (face.Version != 0x00010000 || UnsupportedTables.Any(face.Tables.ContainsKey))
            throw new NotSupportedException("GID outlines require a static monochrome TrueType glyf face.");
        var head = face.Table("head", 54);
        GlyphCount = OpenTypeFace.U16(face.Table("maxp", 6), 4);
        UnitsPerEm = OpenTypeFace.U16(head, 18);
        if (GlyphCount == 0 || UnitsPerEm is < 16 or > 16384)
            throw new InvalidDataException("Invalid TrueType GID or units-per-em count.");
        metricCount = OpenTypeFace.U16(face.Table("hhea", 36), 34);
        if (metricCount == 0 || metricCount > GlyphCount)
            throw new InvalidDataException("Invalid TrueType horizontal metric count.");
        hmtx = face.Table("hmtx", metricCount * 4 + (GlyphCount - metricCount) * 2);
        glyf = face.Table("glyf");
        var loca = face.Table("loca");
        var format = OpenTypeFace.U16(head, 50);
        if (format is not (0 or 1)) throw new InvalidDataException("Invalid TrueType loca format.");
        var stride = format == 0 ? 2 : 4;
        OpenTypeFace.Require(loca, 0, checked((GlyphCount + 1) * stride));
        offsets = new int[GlyphCount + 1];
        for (var i = 0; i < offsets.Length; i++)
        {
            offsets[i] = format == 0 ? OpenTypeFace.U16(loca, i * stride) * 2 :
                OpenTypeFace.Size(OpenTypeFace.U32(loca, i * stride));
            if (offsets[i] > glyf.Length || i > 0 && offsets[i] < offsets[i - 1])
                throw new InvalidDataException("Invalid TrueType glyph location.");
        }
        SourceSha256 = BinaryIdentity.Hash(ownedSource);
        SelectedFaceSha256 = BinaryIdentity.Hash(selected);
        FaceIndex = faceIndex;
    }

    internal string SourceSha256 { get; }
    internal string SelectedFaceSha256 { get; }
    internal int FaceIndex { get; }
    internal int GlyphCount { get; }
    internal int UnitsPerEm { get; }

    internal TrueTypeGlyphOutline Read(int glyphId, CancellationToken cancellationToken = default)
    {
        if (glyphId <= 0 || glyphId >= GlyphCount)
            throw new NotSupportedException("The requested original GID is outside the admitted font.");
        var state = new ReadState(cancellationToken);
        var geometry = ReadGlyph(glyphId, state, new HashSet<int>(), 0);
        var frozen = new ReadOnlyCollection<TrueTypeGlyphContour>(geometry.Contours.Select(points =>
            new TrueTypeGlyphContour(new ReadOnlyCollection<TrueTypeOutlinePoint>(points.ToArray()))).ToArray());
        var advance = OpenTypeFace.U16(hmtx, Math.Min(geometry.MetricsGlyphId, metricCount - 1) * 4);
        return new TrueTypeGlyphOutline(SourceSha256, SelectedFaceSha256, FaceIndex,
            glyphId, UnitsPerEm, advance, frozen);
    }

    private GlyphGeometry ReadGlyph(int glyphId, ReadState state, HashSet<int> visiting, int depth)
    {
        state.Visit();
        if (depth > MaximumDepth || !visiting.Add(glyphId))
            throw new InvalidDataException("Recursive or excessively deep TrueType composite glyph.");
        try
        {
            var start = offsets[glyphId]; var end = offsets[glyphId + 1];
            if (start == end) return new GlyphGeometry(new List<List<TrueTypeOutlinePoint>>(), glyphId);
            RequireGlyph(start, end, start, 10);
            var contours = Signed16(glyf, start);
            if (contours >= 0) return new GlyphGeometry(ReadSimple(start, end, contours, state), glyphId);
            if (contours != -1) throw new NotSupportedException("Unknown TrueType composite glyph header.");
            return ReadComposite(glyphId, start, end, state, visiting, depth);
        }
        finally { visiting.Remove(glyphId); }
    }

    private List<List<TrueTypeOutlinePoint>> ReadSimple(int start, int end, int contourCount, ReadState state)
    {
        // A zero-contour glyph may consist solely of its ten-byte header.
        // If more bytes follow, they must contain a complete instruction
        // length and the advertised instruction payload.
        if (contourCount == 0 && end == start + 10)
            return new List<List<TrueTypeOutlinePoint>>();
        var cursor = start + 10;
        RequireGlyph(start, end, cursor, checked(contourCount * 2 + 2));
        var ends = new int[contourCount];
        for (var i = 0; i < ends.Length; i++)
        {
            ends[i] = OpenTypeFace.U16(glyf, cursor); cursor += 2;
            if (i > 0 && ends[i] <= ends[i - 1]) throw new InvalidDataException("Invalid TrueType contour endpoints.");
        }
        var count = contourCount == 0 ? 0 : checked(ends[contourCount - 1] + 1);
        state.AddPoints(count);
        var instructionLength = OpenTypeFace.U16(glyf, cursor); cursor += 2;
        RequireGlyph(start, end, cursor, instructionLength); cursor += instructionLength;
        var flags = new byte[count];
        for (var i = 0; i < count;)
        {
            RequireGlyph(start, end, cursor, 1);
            var flag = glyf[cursor++];
            if ((flag & 0x80) != 0) throw new NotSupportedException("Reserved TrueType point flag.");
            var repeats = 0;
            if ((flag & 0x08) != 0)
            {
                RequireGlyph(start, end, cursor, 1); repeats = glyf[cursor++];
            }
            if (repeats >= count - i) throw new InvalidDataException("TrueType point repeat exceeds contour.");
            for (var j = 0; j <= repeats; j++) flags[i++] = flag;
        }
        var x = new int[count]; var y = new int[count];
        ReadAxis(x, flags, 0x02, 0x10, ref cursor, start, end);
        ReadAxis(y, flags, 0x04, 0x20, ref cursor, start, end);
        var result = new List<List<TrueTypeOutlinePoint>>(contourCount);
        var first = 0;
        foreach (var last in ends)
        {
            var points = new List<TrueTypeOutlinePoint>(last - first + 1);
            for (var i = first; i <= last; i++) points.Add(new TrueTypeOutlinePoint(x[i], y[i], (flags[i] & 1) != 0));
            result.Add(points); first = last + 1;
        }
        return result;
    }

    private void ReadAxis(int[] values, byte[] flags, int shortBit, int positiveBit,
        ref int cursor, int start, int end)
    {
        var current = 0;
        for (var i = 0; i < values.Length; i++)
        {
            var flag = flags[i]; int delta;
            if ((flag & shortBit) != 0)
            {
                RequireGlyph(start, end, cursor, 1);
                delta = (flag & positiveBit) != 0 ? glyf[cursor] : -glyf[cursor]; cursor++;
            }
            else if ((flag & positiveBit) != 0) delta = 0;
            else
            {
                RequireGlyph(start, end, cursor, 2); delta = Signed16(glyf, cursor); cursor += 2;
            }
            current = checked(current + delta);
            if (current is < -1000000 or > 1000000)
                throw new InvalidDataException("TrueType point exceeds bounded design coordinates.");
            values[i] = current;
        }
    }

    private GlyphGeometry ReadComposite(int glyphId, int start, int end, ReadState state,
        HashSet<int> visiting, int depth)
    {
        var result = new List<List<TrueTypeOutlinePoint>>();
        var cursor = start + 10; int flags;
        int? metricsGlyphId = null;
        var hasInstructions = false;
        do
        {
            state.Visit(); RequireGlyph(start, end, cursor, 4);
            flags = OpenTypeFace.U16(glyf, cursor);
            var child = OpenTypeFace.U16(glyf, cursor + 2); cursor += 4;
            if ((flags & ~0x1FEF) != 0 || child >= GlyphCount)
                throw new InvalidDataException("Invalid TrueType composite flags or GID.");
            if ((flags & 0x0002) == 0)
                throw new NotSupportedException("TrueType point-attached composite placement is outside the GID outline profile.");
            var words = (flags & 0x0001) != 0;
            RequireGlyph(start, end, cursor, words ? 4 : 2);
            var dx = words ? Signed16(glyf, cursor) : unchecked((sbyte)glyf[cursor]);
            var dy = words ? Signed16(glyf, cursor + 2) : unchecked((sbyte)glyf[cursor + 1]);
            cursor += words ? 4 : 2;
            // The pinned font has XY-translated components. A transformed
            // component needs a separately proven offset/grid interpretation.
            if ((flags & (0x0008 | 0x0040 | 0x0080)) != 0)
                throw new NotSupportedException("Transformed TrueType composite placement is outside the GID outline profile.");
            if ((flags & 0x1800) == 0x1800)
                throw new InvalidDataException("Conflicting TrueType composite offset flags.");
            var component = ReadGlyph(child, state, visiting, depth + 1);
            if ((flags & 0x0200) != 0)
            {
                if (metricsGlyphId is not null)
                    throw new NotSupportedException("Multiple USE_MY_METRICS components are outside the GID outline profile.");
                metricsGlyphId = component.MetricsGlyphId;
            }
            hasInstructions |= (flags & 0x0100) != 0;
            foreach (var contour in component.Contours)
            {
                var points = new List<TrueTypeOutlinePoint>(contour.Count);
                foreach (var point in contour)
                {
                    var x = point.X + dx;
                    var y = point.Y + dy;
                    if (Math.Abs(x) > 1000000 || Math.Abs(y) > 1000000)
                        throw new InvalidDataException("TrueType component exceeds bounded design coordinates.");
                    points.Add(new TrueTypeOutlinePoint(x, y, point.OnCurve));
                }
                result.Add(points);
            }
        } while ((flags & 0x0020) != 0);
        if (hasInstructions)
        {
            RequireGlyph(start, end, cursor, 2);
            var length = OpenTypeFace.U16(glyf, cursor); cursor += 2;
            RequireGlyph(start, end, cursor, length);
        }
        return new GlyphGeometry(result, metricsGlyphId ?? glyphId);
    }

    private sealed class GlyphGeometry(List<List<TrueTypeOutlinePoint>> contours, int metricsGlyphId)
    {
        internal List<List<TrueTypeOutlinePoint>> Contours { get; } = contours;
        internal int MetricsGlyphId { get; } = metricsGlyphId;
    }

    private static int Signed16(byte[] bytes, int offset) => unchecked((short)OpenTypeFace.U16(bytes, offset));
    private static void RequireGlyph(int start, int end, int offset, int length)
    {
        if (offset < start || length < 0 || (long)offset + length > end)
            throw new InvalidDataException("Truncated TrueType glyph data.");
    }
    private sealed class ReadState(CancellationToken token)
    {
        private int points, components;
        internal void Visit()
        {
            token.ThrowIfCancellationRequested();
            if (++components > MaximumComponents) throw new InvalidDataException("TrueType component budget exceeded.");
        }
        internal void AddPoints(int count)
        {
            if (count > MaximumPoints - points) throw new InvalidDataException("TrueType point budget exceeded.");
            points += count;
        }
    }
}

internal sealed class TrueTypeGlyphOutline(string sourceSha256, string selectedFaceSha256, int faceIndex, int glyphId, int unitsPerEm,
    int advanceWidth, IReadOnlyList<TrueTypeGlyphContour> contours)
{
    internal string SourceSha256 { get; } = sourceSha256;
    internal string SelectedFaceSha256 { get; } = selectedFaceSha256;
    internal int FaceIndex { get; } = faceIndex;
    internal int GlyphId { get; } = glyphId;
    internal int UnitsPerEm { get; } = unitsPerEm;
    internal int AdvanceWidth { get; } = advanceWidth;
    internal IReadOnlyList<TrueTypeGlyphContour> Contours { get; } = contours;

    // Implied on-curve midpoints are resolved once here, rather than
    // independently in PDF and SVG. Coordinates remain unhinted font units.
    internal IReadOnlyList<TrueTypePathCommand> ToPathCommands()
    {
        var commands = new List<TrueTypePathCommand>();
        foreach (var contour in Contours)
        {
            var points = contour.Points;
            if (points.Count == 0) continue;
            TrueTypeOutlinePoint start;
            var first = 0; var last = points.Count - 1;
            if (points[0].OnCurve) { start = points[0]; first = 1; }
            else if (points[last].OnCurve) { start = points[last]; last--; }
            else start = Midpoint(points[last], points[0]);
            commands.Add(new TrueTypePathCommand(TrueTypePathVerb.Move, start.X, start.Y));
            TrueTypeOutlinePoint? pending = null;
            for (var i = first; i <= last + 1; i++)
            {
                var point = i <= last ? points[i] : start;
                if (!point.OnCurve)
                {
                    if (pending is { } previous)
                    {
                        var mid = Midpoint(previous, point);
                        commands.Add(new TrueTypePathCommand(TrueTypePathVerb.Quadratic, mid.X, mid.Y,
                            previous.X, previous.Y));
                    }
                    pending = point;
                }
                else if (pending is { } control)
                {
                    commands.Add(new TrueTypePathCommand(TrueTypePathVerb.Quadratic, point.X, point.Y,
                        control.X, control.Y));
                    pending = null;
                }
                else if (i <= last)
                    commands.Add(new TrueTypePathCommand(TrueTypePathVerb.Line, point.X, point.Y));
            }
            commands.Add(new TrueTypePathCommand(TrueTypePathVerb.Close, start.X, start.Y));
        }
        return new ReadOnlyCollection<TrueTypePathCommand>(commands);
    }

    private static TrueTypeOutlinePoint Midpoint(TrueTypeOutlinePoint first, TrueTypeOutlinePoint second) =>
        new((first.X + second.X) / 2, (first.Y + second.Y) / 2, true);
}

internal sealed class TrueTypeGlyphContour(IReadOnlyList<TrueTypeOutlinePoint> points)
{
    internal IReadOnlyList<TrueTypeOutlinePoint> Points { get; } = points;
}

internal readonly struct TrueTypeOutlinePoint(double x, double y, bool onCurve)
{
    internal double X { get; } = x;
    internal double Y { get; } = y;
    internal bool OnCurve { get; } = onCurve;
}

internal enum TrueTypePathVerb { Move, Line, Quadratic, Close }

internal readonly struct TrueTypePathCommand(TrueTypePathVerb verb, double x, double y,
    double controlX = 0, double controlY = 0)
{
    internal TrueTypePathVerb Verb { get; } = verb;
    internal double X { get; } = x;
    internal double Y { get; } = y;
    internal double ControlX { get; } = controlX;
    internal double ControlY { get; } = controlY;
}
