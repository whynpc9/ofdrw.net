using System;
using System.Collections.Generic;
using System.Threading;
using PdfSharpCore.Drawing;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>Bounded boundary-only union for exactly orthogonal glyph regions.</summary>
internal static class RectilinearGlyphUnion
{
    internal const int MaximumInputEdges = 1024;
    private const int MaximumAxes = 64;
    private const int MaximumCells = 4096;
    private const int MaximumBoundaryEdges = 2048;
    private const int MaximumLoops = 128;
    private const int MaximumWork = 4_000_000;

    internal static bool TryBuild(IReadOnlyList<XPoint[]> fill, IReadOnlyList<XPoint[]> stroke,
        CancellationToken cancellationToken, out XPoint[][] loops, int maximumWork = MaximumWork)
    {
        loops = Array.Empty<XPoint[]>();
        var work = new Work(cancellationToken, maximumWork);
        work.CheckCancellation();
        try
        {
            var xs = new List<double>(); var ys = new List<double>(); var inputEdges = 0;
            bool Admit(IReadOnlyList<XPoint[]> paths)
            {
                foreach (var path in paths)
                {
                    if (path.Length < 3) return false;
                    for (var index = 0; index < path.Length; index++)
                    {
                        work.Visit();
                        var a = path[index]; var b = path[(index + 1) % path.Length];
                        if (!Finite(a.X) || !Finite(a.Y)) return false;
                        if (a.X == b.X && a.Y == b.Y) continue;
                        if (++inputEdges > MaximumInputEdges || a.X != b.X && a.Y != b.Y) return false;
                        xs.Add(a.X); ys.Add(a.Y);
                    }
                }
                return true;
            }
            if (!Admit(fill) || !Admit(stroke)) return false;
            SortUnique(xs); SortUnique(ys);
            if (xs.Count < 2 || ys.Count < 2 || xs.Count > MaximumAxes || ys.Count > MaximumAxes) return false;
            var width = xs.Count - 1; var height = ys.Count - 1;
            if (width * height > MaximumCells) return false;
            var occupied = new bool[width, height];
            for (var y = 0; y < height; y++)
            for (var x = 0; x < width; x++)
            {
                work.Visit();
                var px = xs[x] + (xs[x + 1] - xs[x]) * .5;
                var py = ys[y] + (ys[y + 1] - ys[y]) * .5;
                // Never snap nearly equal coordinates or classify on a boundary.
                if (!(xs[x] < px && px < xs[x + 1] && ys[y] < py && py < ys[y + 1])) return false;
                occupied[x, y] = Inside(fill, px, py, work) || Inside(stroke, px, py, work);
            }
            var edges = new Dictionary<int, int>(); var incoming = new HashSet<int>(); var rejected = false;
            int Node(int x, int y) => y * xs.Count + x;
            void Edge(int from, int to)
            {
                work.Visit();
                if (edges.Count >= MaximumBoundaryEdges || edges.ContainsKey(from) || !incoming.Add(to))
                { rejected = true; return; }
                edges.Add(from, to);
            }
            for (var y = 0; y < height; y++)
            for (var x = 0; x < width; x++)
            {
                work.Visit();
                if (!occupied[x, y]) continue;
                // Keep the occupied cell on the left of each directed edge.
                if (y == 0 || !occupied[x, y - 1]) Edge(Node(x, y), Node(x + 1, y));
                if (x == width - 1 || !occupied[x + 1, y]) Edge(Node(x + 1, y), Node(x + 1, y + 1));
                if (y == height - 1 || !occupied[x, y + 1]) Edge(Node(x + 1, y + 1), Node(x, y + 1));
                if (x == 0 || !occupied[x - 1, y]) Edge(Node(x, y + 1), Node(x, y));
                if (rejected) return false; // Touching branches retain the existing paint route.
            }
            foreach (var node in edges.Keys)
                if (!incoming.Contains(node)) return false;
            var remaining = new HashSet<int>(edges.Keys); var completed = new List<XPoint[]>();
            while (remaining.Count != 0)
            {
                work.Visit();
                if (completed.Count >= MaximumLoops) return false;
                var first = int.MaxValue;
                foreach (var node in remaining) { work.Visit(); first = Math.Min(first, node); }
                var current = first; var ring = new List<int>();
                do
                {
                    work.Visit();
                    if (!remaining.Remove(current) || !edges.TryGetValue(current, out var next)) return false;
                    ring.Add(current); current = next;
                } while (current != first);
                var changed = true;
                while (changed && ring.Count > 3)
                {
                    changed = false;
                    for (var i = 0; i < ring.Count; i++)
                    {
                        work.Visit();
                        var a = ring[(i + ring.Count - 1) % ring.Count]; var b = ring[i]; var c = ring[(i + 1) % ring.Count];
                        if (a % xs.Count == b % xs.Count && b % xs.Count == c % xs.Count ||
                            a / xs.Count == b / xs.Count && b / xs.Count == c / xs.Count)
                        { ring.RemoveAt(i); changed = true; break; }
                    }
                }
                if (ring.Count < 4) return false;
                var points = new XPoint[ring.Count];
                for (var i = 0; i < points.Length; i++)
                { work.Visit(); points[i] = new XPoint(xs[ring[i] % xs.Count], ys[ring[i] / xs.Count]); }
                completed.Add(points);
            }
            work.CheckCancellation();
            loops = completed.ToArray();
            return true;
        }
        catch (WorkLimitException) { return false; }
    }

    private static bool Inside(IReadOnlyList<XPoint[]> paths, double x, double y, Work work)
    {
        var winding = 0;
        foreach (var path in paths)
        for (var i = 0; i < path.Length; i++)
        {
            work.Visit(); var a = path[i]; var b = path[(i + 1) % path.Length];
            // Admission proved orthogonality; horizontal edges never cross this ray.
            if (a.X <= x || a.X != b.X) continue;
            if (a.Y <= y && y < b.Y) winding++;
            else if (b.Y <= y && y < a.Y) winding--;
        }
        return winding != 0;
    }

    private static void SortUnique(List<double> values)
    {
        values.Sort(); var count = 0;
        for (var i = 0; i < values.Count; i++)
        {
            var value = values[i];
            if (count == 0 || value != values[count - 1]) values[count++] = value;
        }
        if (count < values.Count) values.RemoveRange(count, values.Count - count);
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private sealed class WorkLimitException : Exception { }
    private sealed class Work
    {
        private readonly CancellationToken _token;
        private readonly int _maximum;
        private int _used;
        internal Work(CancellationToken token, int maximum) { _token = token; _maximum = maximum; }
        internal void CheckCancellation() => _token.ThrowIfCancellationRequested();
        internal void Visit()
        {
            CheckCancellation();
            if (++_used > _maximum) throw new WorkLimitException();
        }
    }
}
