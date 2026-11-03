using System;
using System.Collections.Generic;
using System.Linq;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Builds the area a glyph gains when it is dilated only horizontally by
/// <c>h</c> on each side, together with an equal inward margin at each side
/// edge so the overlay meets the original glyph without an antialiasing seam.
/// </summary>
/// <remarks>
/// Every output point lies inside the exact horizontal dilation of the glyph,
/// so the overlay never reaches above or below the original contour. Vertical
/// gaps between stacked horizontal strokes (the dense Han counters in issue
/// #107) are therefore not painted at all, independent of the viewer's
/// hinting and hairline policies. The result is a set of disjoint trapezoids
/// with one orientation for a single nonzero fill.
/// </remarks>
internal static class HorizontalEmboldenOverlay
{
    private const double Epsilon = 1e-9;

    internal readonly struct Edge
    {
        internal Edge(double x0, double y0, double x1, double y1)
        {
            if (y0 <= y1) { X0 = x0; Y0 = y0; X1 = x1; Y1 = y1; Winding = 1; }
            else { X0 = x1; Y0 = y1; X1 = x0; Y1 = y0; Winding = -1; }
        }

        internal double X0 { get; }
        internal double Y0 { get; }
        internal double X1 { get; }
        internal double Y1 { get; }
        internal int Winding { get; }
        internal double XAt(double y) => Y1 - Y0 < Epsilon ? X0 : X0 + (X1 - X0) * (y - Y0) / (Y1 - Y0);
    }

    private readonly struct Boundary
    {
        internal Boundary(Edge edge, double offset) { Edge = edge; Offset = offset; }
        internal Edge Edge { get; }
        internal double Offset { get; }
        internal double XAt(double y) => Edge.XAt(y) + Offset;
    }

    /// <summary>Returns closed quadrilaterals (x, y pairs) in the polygon coordinate space.</summary>
    internal static List<(double X, double Y)[]> Build(IReadOnlyList<IReadOnlyList<(double X, double Y)>> contours, double halfWidth)
    {
        var result = new List<(double X, double Y)[]>();
        if (!(halfWidth > 0)) return result;
        var edges = new List<Edge>();
        foreach (var contour in contours)
        {
            for (var index = 0; index < contour.Count; index++)
            {
                var a = contour[index];
                var b = contour[(index + 1) % contour.Count];
                if (Math.Abs(a.Y - b.Y) > Epsilon) edges.Add(new Edge(a.X, a.Y, b.X, b.Y));
            }
        }
        if (edges.Count == 0) return result;
        var levels = edges.SelectMany(edge => new[] { edge.Y0, edge.Y1 }).Distinct().OrderBy(y => y).ToList();
        for (var level = 0; level + 1 < levels.Count; level++)
        {
            var top = levels[level];
            var bottom = levels[level + 1];
            if (bottom - top < Epsilon) continue;
            var active = edges.Where(edge => edge.Y0 <= top + Epsilon && edge.Y1 >= bottom - Epsilon).ToList();
            if (active.Count < 2) continue;
            // Contour edges may cross inside a band (overlapping CFF parts).
            var splits = new List<double> { top, bottom };
            for (var i = 0; i < active.Count; i++)
                for (var j = i + 1; j < active.Count; j++)
                    AddRoot(splits, top, bottom,
                        active[i].XAt(top) - active[j].XAt(top), active[i].XAt(bottom) - active[j].XAt(bottom));
            splits.Sort();
            for (var piece = 0; piece + 1 < splits.Count; piece++)
                EmitBand(result, active, splits[piece], splits[piece + 1], halfWidth);
        }
        return result;
    }

    private static void EmitBand(List<(double X, double Y)[]> result, List<Edge> active, double top, double bottom, double h)
    {
        if (bottom - top < Epsilon) return;
        var intervals = Intervals(active, (top + bottom) / 2);
        if (intervals.Count == 0) return;
        // Overlay topology changes where a gap or an interval crosses 2h.
        var splits = new List<double> { top, bottom };
        for (var index = 0; index < intervals.Count; index++)
        {
            var (left, right) = intervals[index];
            AddRoot(splits, top, bottom, right.XAt(top) - left.XAt(top) - 2 * h, right.XAt(bottom) - left.XAt(bottom) - 2 * h);
            if (index + 1 < intervals.Count)
            {
                var next = intervals[index + 1].Left;
                AddRoot(splits, top, bottom, next.XAt(top) - right.XAt(top) - 2 * h, next.XAt(bottom) - right.XAt(bottom) - 2 * h);
            }
        }
        splits.Sort();
        for (var piece = 0; piece + 1 < splits.Count; piece++)
        {
            var a = splits[piece];
            var b = splits[piece + 1];
            if (b - a < Epsilon) continue;
            var middle = (a + b) / 2;
            foreach (var (left, right) in Overlay(Intervals(active, middle), middle, h))
            {
                result.Add(new[]
                {
                    (left.XAt(a), a), (right.XAt(a), a), (right.XAt(b), b), (left.XAt(b), b)
                });
            }
        }
    }

    private static List<(Boundary Left, Boundary Right)> Overlay(List<(Edge Left, Edge Right)> intervals, double y, double h)
    {
        var pieces = new List<(Boundary Left, Boundary Right)>();
        var index = 0;
        while (index < intervals.Count)
        {
            // One group of intervals whose horizontal dilations touch.
            var last = index;
            while (last + 1 < intervals.Count &&
                   intervals[last + 1].Left.XAt(y) - intervals[last].Right.XAt(y) <= 2 * h) last++;
            var current = new Boundary(intervals[index].Left, -h);
            for (var member = index; member <= last; member++)
            {
                var (left, right) = intervals[member];
                if (right.XAt(y) - left.XAt(y) <= 2 * h) continue;
                pieces.Add((current, new Boundary(left, h)));
                current = new Boundary(right, -h);
            }
            pieces.Add((current, new Boundary(intervals[last].Right, h)));
            index = last + 1;
        }
        return pieces;
    }

    private static List<(Edge Left, Edge Right)> Intervals(List<Edge> active, double y)
    {
        var crossings = active.Where(edge => edge.Y0 < y && edge.Y1 > y)
            .Select(edge => (X: edge.XAt(y), Edge: edge)).OrderBy(item => item.X).ToList();
        var intervals = new List<(Edge Left, Edge Right)>();
        var winding = 0;
        Edge start = default;
        foreach (var (_, edge) in crossings)
        {
            var before = winding;
            winding += edge.Winding;
            if (before == 0 && winding != 0) start = edge;
            else if (before != 0 && winding == 0) intervals.Add((start, edge));
        }
        return intervals;
    }

    private static void AddRoot(List<double> splits, double top, double bottom, double atTop, double atBottom)
    {
        if (atTop * atBottom >= 0) return;
        var y = top + (bottom - top) * atTop / (atTop - atBottom);
        if (y > top + Epsilon && y < bottom - Epsilon) splits.Add(y);
    }
}
