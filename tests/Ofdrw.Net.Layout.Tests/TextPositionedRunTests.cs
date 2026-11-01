using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextPositionedRunTests
{
    private static readonly Lazy<TextFontFace> Face = new(() => Load("DejaVuSans-R22.ttf"));

    [Fact]
    public void TwoDimensionalPenOffsetsAndInkInvertYOnceAndKeepFinalAdvance()
    {
        var shape = Controlled(new(36, 20, 1, 100, 20, 5, 10, new(-2, -5, 12, 15), false),
            new(57, 21, 1, 80, -10, -7, -3, null, true));
        var result = TextPositionedRun.Create(shape, 4, 10, 30);
        var scale = 4d / 2048;
        Assert.Same(shape, result.ShapedRun);
        Assert.Equal(4, result.FontSizeMillimeters); Assert.Equal(10, result.OriginXMillimeters);
        Assert.Equal(30, result.BaselineYMillimeters);
        Assert.Equal(10 + 5 * scale, result.Glyphs[0].OriginXMillimeters);
        Assert.Equal(30 - 10 * scale, result.Glyphs[0].OriginYMillimeters);
        Assert.Equal(10 + 93 * scale, result.Glyphs[1].OriginXMillimeters);
        Assert.Equal(30 - 17 * scale, result.Glyphs[1].OriginYMillimeters);
        Assert.Equal(-20 * scale, result.Glyphs[0].YAdvanceMillimeters);
        Assert.Equal(10 * scale, result.Glyphs[1].YAdvanceMillimeters);
        Assert.Equal(180 * scale, result.XAdvanceMillimeters);
        Assert.Equal(-10 * scale, result.YAdvanceMillimeters);
        Assert.Equal(88 * scale, result.Glyphs[1].OriginXMillimeters - result.Glyphs[0].OriginXMillimeters);
        Assert.Equal(-7 * scale, result.Glyphs[1].OriginYMillimeters - result.Glyphs[0].OriginYMillimeters);
        var ink = result.InkBoundsMillimeters!;
        Assert.Equal(10 + 3 * scale, ink.X); Assert.Equal(30 - 20 * scale, ink.Y);
        Assert.Equal(12 * scale, ink.Width); Assert.Equal(15 * scale, ink.Height);
        Assert.True(result.ShapedRun.Glyphs[1].UnsafeToBreak);
        Assert.Throws<NotSupportedException>(() => ((IList<OfdPositionedGlyph>)result.Glyphs).Clear());
    }

    [Theory]
    [InlineData("AV", 36)]
    [InlineData("ffi", 5044)]
    [InlineData("a\u0323\u0301", 2463)]
    public void NativeRunRetainsEveryOccurrenceAndClusterWithoutReshaping(string text, uint firstGid)
    {
        var shaped = new HarfBuzzTextShaper().Shape(new TextShapeRequest(text, 70, Face.Value));
        Assert.Equal(firstGid, shaped.Glyphs[0].GlyphId);
        var result = TextPositionedRun.Create(shaped, 5, -10, -20);
        Assert.Same(shaped, result.ShapedRun);
        Assert.Equal(shaped.Glyphs.Select(g => (g.GlyphId, g.ClusterStart, g.ClusterLength)),
            result.Glyphs.Select(g => (g.GlyphId, g.ClusterStart, g.ClusterLength)));
        Assert.Equal(shaped.XAdvance * 5d / Face.Value.UnitsPerEm, result.XAdvanceMillimeters);
        if (text.Length == 3 && text[0] == 'a')
        {
            Assert.Equal(2, result.Glyphs.Count);
            Assert.Equal(0, result.Glyphs[1].XAdvanceMillimeters);
            Assert.All(result.Glyphs, g => { Assert.Equal(70, g.ClusterStart); Assert.Equal(3, g.ClusterLength); });
        }
    }

    [Fact]
    public void ContextualArabicRetainsExactRequestAndNativeRtlOrder()
    {
        var whole = new TextShapeRequest("ببب", 100, Face.Value, TextDirection.RightToLeft, "Arab", "ar");
        var item = TextShapeRequest.CreateContextualItem(whole, new TextSourceSpan(101, 2));
        var shaped = new HarfBuzzTextShaper().ShapeContextualItem(item);
        var positioned = TextPositionedRun.Create(shaped, 4);
        Assert.Same(whole, positioned.ShapedRun.Request.Context);
        Assert.Same(item, positioned.ShapedRun.Request);
        Assert.Equal(new uint[] { 5258, 5260 }, positioned.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new[] { 102, 101 }, positioned.Glyphs.Select(g => g.ClusterStart));
        // This is a query result, not admission of Arabic to the prelayout writer.
        Assert.Equal(TextLayoutError.UnsupportedFontProfile, Assert.Throws<TextLayoutException>(() =>
            new TextPrelayoutRequest(item.Text, new[] { Face.Value }, 4, 50, script: "Arab", language: "ar")).Error);
    }

    [Fact]
    public void NonBmpMaximumSourceEndpointAndFaceSpecificSizesRemainSeparate()
    {
        var native = new HarfBuzzTextShaper();
        var original = native.Shape(new TextShapeRequest("\U00010300", int.MaxValue - 2, Face.Value));
        var glyph = Assert.Single(TextPositionedRun.Create(original, 4).Glyphs);
        Assert.Equal(int.MaxValue - 2, glyph.ClusterStart); Assert.Equal(2, glyph.ClusterLength);
        var noto = Load("NotoSans-Regular.ttf");
        var other = native.Shape(new TextShapeRequest("A", 12, noto));
        var a = TextPositionedRun.Create(other, 6);
        var b = TextPositionedRun.Create(native.Shape(new TextShapeRequest("A", 13, Face.Value)), 3);
        Assert.Equal(1000, a.ShapedRun.Request.Face.UnitsPerEm);
        Assert.Equal(2048, b.ShapedRun.Request.Face.UnitsPerEm);
        Assert.Equal(other.XAdvance * 6d / 1000, a.XAdvanceMillimeters);
        Assert.Equal(b.ShapedRun.XAdvance * 3d / 2048, b.XAdvanceMillimeters);
        Assert.NotEqual(a.ShapedRun.Request.Face.FaceSha256, b.ShapedRun.Request.Face.FaceSha256);
    }

    [Fact]
    public void EmptyRunHasNoInkOrGlyphsAndDoesNotInventAdvance()
    {
        var shaped = new TextShapedRun(new TextShapeRequest("", int.MaxValue, Face.Value),
            Array.Empty<TextShapedGlyph>(), "controlled-empty");
        var result = TextPositionedRun.Create(shaped, 4, 2, 3);
        Assert.Same(shaped, result.ShapedRun); Assert.Empty(result.Glyphs);
        Assert.Null(result.InkBoundsMillimeters); Assert.Equal(0, result.XAdvanceMillimeters);
        Assert.Equal(0, result.YAdvanceMillimeters);
    }

    [Fact]
    public void NegativePensAndAdvancesRemainSignedWithoutInventingInk()
    {
        var shaped = Controlled(new(36, 20, 1, -100, -20, -5, -10, null, false),
            new(57, 21, 1, 80, 10, 7, 3, null, false));
        var result = TextPositionedRun.Create(shaped, 4, -10, -30);
        var scale = 4d / 2048;
        Assert.Equal(-10 - 93 * scale, result.Glyphs[1].OriginXMillimeters);
        Assert.Equal(-30 + 17 * scale, result.Glyphs[1].OriginYMillimeters);
        Assert.Equal(-20 * scale, result.XAdvanceMillimeters);
        Assert.Equal(10 * scale, result.YAdvanceMillimeters);
        Assert.Null(result.InkBoundsMillimeters);
    }

    [Fact]
    public void InkUnionUsesTranslatedInkInsteadOfAdvances()
    {
        var shaped = Controlled(new(36, 20, 1, 100, 20, 5, 10, new(-2, -5, 12, 15), false),
            new(57, 21, 1, 80, -10, -7, -3, new(-100, -40, 8, 9), false));
        var ink = TextPositionedRun.Create(shaped, 4, 10, 30).InkBoundsMillimeters!;
        var scale = 4d / 2048;
        Assert.Equal(10 - 7 * scale, ink.X); Assert.Equal(30 - 20 * scale, ink.Y);
        Assert.Equal(22 * scale, ink.Width); Assert.Equal(43 * scale, ink.Height);
    }

    [Theory]
    [InlineData(0, 0, 0)]
    [InlineData(-1, 0, 0)]
    [InlineData(1001, 0, 0)]
    [InlineData(double.NaN, 0, 0)]
    [InlineData(double.PositiveInfinity, 0, 0)]
    [InlineData(double.Epsilon, 0, 0)]
    [InlineData(4, double.NegativeInfinity, 0)]
    [InlineData(4, 0, double.NaN)]
    [InlineData(4, 1000001, 0)]
    public void InvalidSizeScaleOrOriginRejects(double size, double x, double y)
    {
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            TextPositionedRun.Create(Simple(), size, x, y)).Error);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LargeIntermediatePenCannotHideBehindOppositeOffsetOrFinalAdvance(bool vertical)
    {
        var a = int.MaxValue;
        var shape = Controlled(new(36, 20, 1, vertical ? 0 : a, vertical ? a : 0, 0, 0, null, false),
            new(57, 21, 1, vertical ? 0 : -a, vertical ? -a : 0,
                vertical ? 0 : -a, vertical ? -a : 0, null, false));
        Assert.Equal(0, shape.XAdvance); Assert.Equal(0, shape.YAdvance);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            TextPositionedRun.Create(shape, 4)).Error);
    }

    [Fact]
    public void FiniteDesignInkOverflowAndTranslatedPageOverflowRejectAtomically()
    {
        var shape = Controlled(new(36, 20, 1, 100, 0, 0, 0,
            new(double.MaxValue, 0, double.MaxValue, 1), false), new(57, 21, 1, 0, 0, 0, 0, null, false));
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() => TextPositionedRun.Create(shape, 4)).Error);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            TextPositionedRun.Create(Simple(), 4, 1_000_000)).Error);
    }

    [Fact]
    public void PreCanceledCallerTokenIsPreserved()
    {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        var error = Assert.Throws<OperationCanceledException>(() =>
            TextPositionedRun.Create(Simple(), 4, cancellationToken: cancellation.Token));
        Assert.Equal(cancellation.Token, error.CancellationToken);
        Assert.Throws<ArgumentNullException>(() => TextPositionedRun.Create(null!, 4));
    }

    private static TextShapedRun Simple() => Controlled(new(36, 20, 1, 100, 0, 0, 0, null, false),
        new(57, 21, 1, 80, 0, 0, 0, null, false));
    private static TextShapedRun Controlled(params TextShapedGlyph[] glyphs) =>
        new(new TextShapeRequest("AV", 20, Face.Value), glyphs, "controlled-2d");
    private static TextFontFace Load(string name)
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", name);
            if (File.Exists(path)) return new TextFontFace(name, File.ReadAllBytes(path));
        }
        throw new FileNotFoundException("Materialize existing deterministic font: " + name);
    }
}
