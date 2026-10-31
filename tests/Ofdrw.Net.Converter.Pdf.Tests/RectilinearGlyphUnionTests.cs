using System.Collections;
using Ofdrw.Net.Converter.Pdf.Internal;
using PdfSharpCore.Drawing;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class RectilinearGlyphUnionTests
{
    [Fact]
    public void CanonicalBoundaryRetainsTheLegitimateStrokeAndRemainingCounter()
    {
        var fill = new[] { Rectangle(0, 0, 6, 6), Reverse(Rectangle(2, 2, 4, 4)) };
        var stroke = new[]
        {
            Rectangle(-.25, -.25, 6.25, 6.25), Reverse(Rectangle(.25, .25, 5.75, 5.75)),
            Rectangle(1.75, 1.75, 4.25, 4.25), Reverse(Rectangle(2.25, 2.25, 3.75, 3.75))
        };
        Assert.True(RectilinearGlyphUnion.TryBuild(fill, stroke, default, out var loops));
        Assert.Equal(2, loops.Length);
        Assert.All(loops, loop => Assert.Equal(4, loop.Length));
        Assert.Equal(40, loops.Sum(SignedArea));
        Assert.Contains(loops, loop => SignedArea(loop) < 0);
        Assert.False(Contains(loops, 3, 3));
        Assert.False(Contains(loops, -.3, 3));
        // This original-fill white point is valid stroke ink, not a white-pixel oracle.
        Assert.False(Contains(fill, 2.1, 3));
        Assert.True(Contains(loops, 2.1, 3));
    }

    [Fact]
    public void OverlappingComponentsHaveOneBoundaryWithoutInteriorMultiplicity()
    {
        Assert.True(RectilinearGlyphUnion.TryBuild(new[] { Rectangle(0, 0, 2, 2) },
            new[] { Rectangle(1, 0, 3, 2) }, default, out var loops));
        var loop = Assert.Single(loops);
        Assert.Equal(4, loop.Length);
        Assert.Equal(6, SignedArea(loop));
        Assert.All(loop, p => Assert.True(p.X is 0 or 3));
    }

    [Fact]
    public void InputNonzeroWindingAndIndependentStrokeAreNotConcatenated()
    {
        var a = Rectangle(0, 0, 2, 2); var b = Reverse(Rectangle(1, 0, 3, 2));
        Assert.True(RectilinearGlyphUnion.TryBuild(new[] { a, b }, Array.Empty<XPoint[]>(), default, out var cancelled));
        Assert.Equal(4, cancelled.Sum(SignedArea));
        Assert.False(Contains(cancelled, 1.5, 1));
        Assert.True(RectilinearGlyphUnion.TryBuild(new[] { a }, new[] { Reverse(a) }, default, out var independent));
        Assert.Equal(4, SignedArea(Assert.Single(independent)));
    }

    [Fact]
    public void UnsupportedGeometryAndBudgetsNeverReturnPartialLoops()
    {
        Assert.False(RectilinearGlyphUnion.TryBuild(new[] { Rectangle(0, 0, 1, 1), Rectangle(1, 1, 2, 2) },
            Array.Empty<XPoint[]>(), default, out var touching));
        Assert.Empty(touching);
        Assert.False(RectilinearGlyphUnion.TryBuild(new[] { new[] { new XPoint(0, 0), new XPoint(1, 1), new XPoint(0, 2) } },
            Array.Empty<XPoint[]>(), default, out var oblique));
        Assert.Empty(oblique);
        var manyAxes = Enumerable.Range(0, 65).Select(i => Rectangle(i * 3, 0, i * 3 + 1, 1)).ToArray();
        Assert.False(RectilinearGlyphUnion.TryBuild(manyAxes, Array.Empty<XPoint[]>(), default, out var axes));
        Assert.Empty(axes);
        var tooManyEdges = Enumerable.Repeat(Rectangle(0, 0, 1, 1), 257).ToArray();
        Assert.False(RectilinearGlyphUnion.TryBuild(tooManyEdges, Array.Empty<XPoint[]>(), default, out var edges));
        Assert.Empty(edges);
        Assert.False(RectilinearGlyphUnion.TryBuild(new[] { Rectangle(0, 0, 2, 2) },
            new[] { Rectangle(1, 0, 3, 2) }, default, out var work, maximumWork: 24));
        Assert.Empty(work);
    }

    [Fact]
    public void CancellationBeforeAndDuringClassificationPropagatesWithoutOutput()
    {
        using var before = new CancellationTokenSource(); before.Cancel();
        XPoint[][] loops = Array.Empty<XPoint[]>();
        Assert.Throws<OperationCanceledException>(() => RectilinearGlyphUnion.TryBuild(
            new[] { Rectangle(0, 0, 2, 2) }, Array.Empty<XPoint[]>(), before.Token, out loops));
        Assert.Empty(loops);
        using var during = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => RectilinearGlyphUnion.TryBuild(
            new CancelOnSecondEnumeration(during, Rectangle(0, 0, 2, 2)),
            new[] { Rectangle(1, 0, 3, 2) }, during.Token, out loops));
        Assert.Empty(loops);
    }

    [Fact]
    public void CurvedStrokeGeometryRetainsTheExistingRoute()
    {
        var geometry = new PdfGlyphStrokeGeometry(.001);
        geometry.BeginFigure(); geometry.MoveTo(new System.Numerics.Vector2(0, 0));
        geometry.CubicBezierTo(new(0, 0), new(0, 2), new(2, 2), new(2, 0));
        geometry.LineTo(new(0, 0)); geometry.EndFigure();
        Assert.False(geometry.TryBuildFillUnion(.5, default, out var result));
        Assert.Null(result);
        Assert.NotNull(geometry.Build(.5));
    }

    [Fact]
    public void EmptyFiguresAndUnrepresentableCellCentersRetainTheExistingRoute()
    {
        var empty = new PdfGlyphStrokeGeometry(.001);
        for (var i = 0; i < 200; i++) { empty.BeginFigure(); empty.EndFigure(); }
        Assert.False(empty.TryBuildFillUnion(.5, default, out var path));
        Assert.Null(path);
        Assert.False(RectilinearGlyphUnion.TryBuild(new[] { Rectangle(1, 0, Math.BitIncrement(1d), 1) },
            Array.Empty<XPoint[]>(), default, out var collapsed));
        Assert.Empty(collapsed);
    }

    private static XPoint[] Rectangle(double l, double b, double r, double t) =>
        [new(l, b), new(r, b), new(r, t), new(l, t)];
    private static XPoint[] Reverse(XPoint[] path) => path.Reverse().ToArray();
    private static double SignedArea(XPoint[] path) => path.Select((a, i) =>
        a.X * path[(i + 1) % path.Length].Y - a.Y * path[(i + 1) % path.Length].X).Sum() * .5;
    private static bool Contains(IEnumerable<XPoint[]> paths, double x, double y)
    {
        var winding = 0;
        foreach (var path in paths)
        for (var i = 0; i < path.Length; i++)
        {
            var a = path[i]; var b = path[(i + 1) % path.Length];
            var turn = (b.X - a.X) * (y - a.Y) - (b.Y - a.Y) * (x - a.X);
            if (a.Y <= y && y < b.Y && turn > 0) winding++;
            else if (b.Y <= y && y < a.Y && turn < 0) winding--;
        }
        return winding != 0;
    }

    private sealed class CancelOnSecondEnumeration(CancellationTokenSource cancellation, XPoint[] path)
        : IReadOnlyList<XPoint[]>
    {
        private int _enumerations;
        public int Count => 1;
        public XPoint[] this[int index] => index == 0 ? path : throw new IndexOutOfRangeException();
        public IEnumerator<XPoint[]> GetEnumerator()
        {
            if (++_enumerations == 2) cancellation.Cancel();
            return ((IEnumerable<XPoint[]>)new[] { path }).GetEnumerator();
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
