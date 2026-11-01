using System.Collections;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextSizedFontSpanShapingTests
{
    private static readonly Lazy<TextFontFace> D = new(() => Load("DejaVuSans-R22.ttf"));
    private static readonly Lazy<TextFontFace> N = new(() => Load("NotoSans-Regular.ttf"));

    [Fact]
    public void SameFaceSizeAliasesCoalesceFfiAndRetainOriginalObjects()
    {
        var alias = Load("DejaVuSans-R22.ttf", "alias");
        var spans = new[] { S(20, 1, D.Value, 4, 7), S(21, 2, alias, 4, 7) };
        var request = new TextSizedFontSpanShapingRequest("ffi", 20, spans, maximumShapeCalls: 1);
        var meter = new Meter(); var result = new TextSizedFontSpanShapingEngine(meter).Shape(request);
        var run = Assert.Single(result.Runs);
        Assert.Same(request, result.Request); Assert.Same(spans[1], request.Spans[1]);
        Assert.Same(alias, request.Spans[1].Span.Face); Assert.Equal(1, meter.Calls);
        Assert.Equal(2, run.SpanCount); Assert.Equal(5044u, Assert.Single(run.Shaping.ShapedRun.Glyphs).GlyphId);
        Assert.Equal(1980, run.Shaping.ShapedRun.XAdvance); Assert.Equal(new[] { 7 }, run.GlyphStyleKeys);
        Assert.Same(run.Shaping.ShapedRun, run.PositionedRun.ShapedRun);
        Assert.Equal(1980 * 4d / 2048, run.PositionedRun.XAdvanceMillimeters);
        Assert.Equal(2, Assert.Single(run.Shaping.GlyphCoverage).Slices.Count);
        Assert.Throws<NotSupportedException>(() => ((IList<TextSizedFontSpanShapedRun>)result.Runs).Clear());
    }

    [Fact]
    public void DifferentSizesStartNativeItemsBeforeKerningInsteadOfScalingWholeGlyphs()
    {
        var meter = new Meter(); var request = R("AV", 20, S(20, 1, D.Value, 4, 7), S(21, 1, D.Value, 6, 9));
        var runs = new TextSizedFontSpanShapingEngine(meter).Shape(request).Runs;
        Assert.Equal(2, meter.Calls); Assert.Equal(2, runs.Count);
        Assert.Equal(new[] { 1401, 1401 }, runs.Select(r => r.Shaping.ShapedRun.XAdvance));
        Assert.Equal(new[] { 4d, 6d }, runs.Select(r => r.FontSizeMillimeters));
        Assert.Equal(6.8408203125, runs.Sum(r => r.PositionedRun.XAdvanceMillimeters));
        Assert.NotEqual(6.5849609375, runs.Sum(r => r.PositionedRun.XAdvanceMillimeters));
        Assert.All(meter.Requests, r => Assert.Equal("AV", r.Context!.Text));
        Assert.Equal(new[] { 0, 1 }, runs.Select(r => r.FirstSpanIndex));
        Assert.Equal(new[] { 7, 9 }, runs.Select(r => Assert.Single(r.GlyphStyleKeys)));
        Assert.All(runs, r => { Assert.Same(r.Shaping.ShapedRun, r.PositionedRun.ShapedRun);
            Assert.Equal(0, r.PositionedRun.OriginXMillimeters); Assert.Equal(0, r.PositionedRun.BaselineYMillimeters); });
    }

    [Fact]
    public void OpaqueKeysDoNotSplitCompatibleSizeKerning()
    {
        var meter = new Meter(); var run = Assert.Single(new TextSizedFontSpanShapingEngine(meter).Shape(
            R("AV", 20, S(20, 1, D.Value, 4, 7), S(21, 1, D.Value, 4, 9))).Runs);
        Assert.Equal(1, meter.Calls); Assert.Equal(2671, run.Shaping.ShapedRun.XAdvance);
        Assert.Equal(new[] { 7, 9 }, run.GlyphStyleKeys);
    }

    [Fact]
    public void SizeBoundaryReshapesFAndFiInsteadOfAssigningTwoSizesToWholeFfi()
    {
        var result = Engine().Shape(R("ffi", 20, S(20, 1, D.Value, 4, 7), S(21, 2, D.Value, 6, 9)));
        Assert.Equal(new uint[] { 73, 5042 }, result.Runs.Select(r => Assert.Single(r.Shaping.ShapedRun.Glyphs).GlyphId));
        Assert.Equal(new[] { 1, 2 }, result.Runs.Select(r => Assert.Single(r.Shaping.ShapedRun.Glyphs).ClusterLength));
        Assert.All(result.Runs, r => Assert.Equal("ffi", r.Shaping.ShapedRun.Request.Context!.Text));
        var meter = new Meter();
        var error = Assert.Throws<TextLayoutException>(() => new TextSizedFontSpanShapingEngine(meter).Shape(
            R("ffi", 20, S(20, 1, D.Value, 4, 7), S(21, 2, D.Value, 4, 9))));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error); Assert.Equal(1, meter.Calls);
        Assert.Equal(20, error.SourceStart); Assert.Equal(3, error.SourceLength);
    }

    [Fact]
    public void SharedMarkOccurrencesKeepOneSizeAndUniqueKeyGuard()
    {
        var request = R("a\u0323\u0301", 70, S(70, 1, D.Value, 4, 7), S(71, 2, D.Value, 4, 7));
        var run = Assert.Single(Engine().Shape(request).Runs);
        Assert.Equal(new uint[] { 2463, 690 }, run.PositionedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(0, run.PositionedRun.Glyphs[1].XAdvanceMillimeters);
        Assert.All(run.PositionedRun.Glyphs, g => { Assert.Equal(70, g.ClusterStart); Assert.Equal(3, g.ClusterLength); });
        var meter = new Meter();
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(meter).Shape(R(request.Text, 70,
                S(70, 1, D.Value, 4, 7), S(71, 2, D.Value, 4, 9)))).Error);
        Assert.Equal(1, meter.Calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void OriginalSizeSwitchInsideMarkRefusesBeforeAnyCallEvenOutsideSelection(bool selected)
    {
        var original = R("Xa\u0323V", 0, S(0, 1, D.Value, 4), S(1, 1, D.Value, 4),
            S(2, 1, D.Value, 6), S(3, 1, D.Value, 4));
        var request = selected ? TextSizedFontSpanShapingRequest.CreateContextualItem(original, new TextSourceSpan(0, 1)) : original;
        var meter = new Meter();
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(meter).Shape(request)).Error);
        Assert.Equal(0, meter.Calls);
    }

    [Fact]
    public void ArabicJoiningContextSurvivesSizeChangeWithoutGlobalVisualReordering()
    {
        var original = new TextSizedFontSpanShapingRequest("ببب", 100,
            new[] { S(100, 1, D.Value, 4, 7), S(101, 2, D.Value, 6, 9) }, TextDirection.RightToLeft, "Arab", "ar");
        var runs = Engine().Shape(original).Runs;
        Assert.Equal(new[] { 100, 101 }, runs.Select(r => r.Shaping.ShapedRun.Request.SourceStart));
        Assert.Equal(5259u, Assert.Single(runs[0].PositionedRun.Glyphs).GlyphId);
        Assert.Equal(new uint[] { 5258, 5260 }, runs[1].PositionedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new[] { 102, 101 }, runs[1].PositionedRun.Glyphs.Select(g => g.ClusterStart));
        Assert.All(runs, r => { Assert.Equal("ببب", r.Shaping.ShapedRun.Request.Context!.Text);
            Assert.Equal(TextDirection.RightToLeft, r.Shaping.ShapedRun.Request.Direction); });
    }

    [Fact]
    public void SelectedPartialSpansKeepActualSizeFaceKeyAndUnchangedSpanIdentity()
    {
        var alias = Load("DejaVuSans-R22.ttf", "alias");
        var original = new TextSizedFontSpanShapingRequest("بببAV", 100,
            new[] { S(100, 2, D.Value, 4, 7), S(102, 1, alias, 6, 9), S(103, 2, N.Value, 5, 11) },
            TextDirection.RightToLeft, "Arab", "ar");
        var item = TextSizedFontSpanShapingRequest.CreateContextualItem(original, new TextSourceSpan(101, 3));
        Assert.Same(original, item.Context); Assert.Same(original.Spans[1], item.Spans[1]);
        Assert.NotSame(original.Spans[0], item.Spans[0]); Assert.Same(D.Value, item.Spans[0].Span.Face);
        Assert.Same(alias, item.Spans[1].Span.Face); Assert.Same(N.Value, item.Spans[2].Span.Face);
        var result = Engine().Shape(item); Assert.Same(item, result.Request);
        Assert.Equal(new[] { 4d, 6d, 5d }, result.Runs.Select(r => r.FontSizeMillimeters));
        Assert.Equal(new uint[] { 5260, 5258, 36 }, result.Runs.Select(r => Assert.Single(r.PositionedRun.Glyphs).GlyphId));
        Assert.Equal(new[] { 7, 9, 11 }, result.Runs.Select(r => Assert.Single(r.GlyphStyleKeys)));
        Assert.All(result.Runs, r => Assert.Equal(original.Text, r.Shaping.ShapedRun.Request.Context!.Text));
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            TextSizedFontSpanShapingRequest.CreateContextualItem(item, new TextSourceSpan(101, 1))).Error);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ThreeSizeItemsAndSelectedOffsetsRetainOriginalFeatureOrder(bool lastDisables)
    {
        var features = new List<TextFeature> { new("kern", 0, 0, 7) };
        if (lastDisables) { features.Add(new("liga", 1, 2, 3)); features.Add(new("liga", 0, 0, 7)); }
        else { features.Add(new("liga", 0, 0, 7)); features.Add(new("liga", 1, 2, 3)); }
        features.Add(new("kern", 0, 3, 0));
        var original = new TextSizedFontSpanShapingRequest("XAffiVZ", 70,
            new[] { S(70, 2, D.Value, 4, 1), S(72, 3, D.Value, 6, 2), S(75, 2, D.Value, 4, 3) }, features: features);
        var item = TextSizedFontSpanShapingRequest.CreateContextualItem(original, new TextSourceSpan(71, 5));
        var meter = new Meter(); var runs = new TextSizedFontSpanShapingEngine(meter).Shape(item).Runs;
        Assert.Equal(3, meter.Calls); Assert.Equal(new[] { 71, 72, 75 }, meter.Requests.Select(r => r.SourceStart));
        Assert.Equal(new uint[] { 36 }, runs[0].PositionedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(lastDisables ? new uint[] { 73, 73, 76 } : new uint[] { 5044 }, runs[1].PositionedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new uint[] { 57 }, runs[2].PositionedRun.Glyphs.Select(g => g.GlyphId));
        Assert.All(meter.Requests, r => { Assert.Equal(original.Text, r.Context!.Text);
            Assert.Equal(features.Select(F), r.Context.Features.Select(F)); });
        Assert.Equal(lastDisables ? new[] { 2, 2, 2 } : new[] { 2 }, runs[1].GlyphStyleKeys);
    }

    [Fact]
    public void NonBmpMaximumEndpointUsesOriginalSizesAndDifferentFaceUpems()
    {
        var start = int.MaxValue - 5;
        var original = R("\U00010300a\u0323A", start, S(start, 2, D.Value, 4),
            S(start + 2, 2, D.Value, 6, 1), S(start + 4, 1, N.Value, 3, 2));
        var item = TextSizedFontSpanShapingRequest.CreateContextualItem(original, new TextSourceSpan(start + 2, 3));
        Assert.Same(original.Spans[1], item.Spans[0]); Assert.Same(original.Spans[2], item.Spans[1]);
        var runs = Engine().Shape(item).Runs;
        Assert.Equal(new[] { 2048, 1000 }, runs.Select(r => r.Shaping.ShapedRun.Request.Face.UnitsPerEm));
        Assert.Equal(new[] { start + 2, start + 4 }, runs.Select(r => Assert.Single(r.PositionedRun.Glyphs).ClusterStart));
        Assert.Equal(new uint[] { 2463, 36 }, runs.Select(r => Assert.Single(r.PositionedRun.Glyphs).GlyphId));
        Assert.Equal(new[] { 6d, 3d }, runs.Select(r => r.PositionedRun.FontSizeMillimeters));
        Assert.All(runs, r => Assert.Equal(original.Text, r.Shaping.ShapedRun.Request.Context!.Text));
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            TextSizedFontSpanShapingRequest.CreateContextualItem(original, new TextSourceSpan(start + 1, 1))).Error);
    }

    [Fact]
    public void SameDisplayNameDifferentProgramsKeepOwnFaceAndSize()
    {
        var a = Load("NotoSans-Regular.ttf", "same"); var b = Load("DejaVuSans-R22.ttf", "same");
        var runs = Engine().Shape(R("AV", 20, S(20, 1, a, 6), S(21, 1, b, 3))).Runs;
        Assert.Equal(2, runs.Count); Assert.Same(a, runs[0].Shaping.ShapedRun.Request.Face);
        Assert.Same(b, runs[1].Shaping.ShapedRun.Request.Face);
        Assert.Equal(6d / 1000 * runs[0].Shaping.ShapedRun.XAdvance, runs[0].PositionedRun.XAdvanceMillimeters);
        Assert.Equal(3d / 2048 * runs[1].Shaping.ShapedRun.XAdvance, runs[1].PositionedRun.XAdvanceMillimeters);
    }

    [Fact]
    public void CallerCollectionsAreSnapshotsAndIndexedBudgetsPrecedeReads()
    {
        var first = S(20, 1, D.Value, 4); var spans = new[] { first, S(21, 1, D.Value, 6) };
        var features = new List<TextFeature> { new("kern", 0, 0, 2) };
        var request = new TextSizedFontSpanShapingRequest("AV", 20, spans, features: features);
        spans[0] = S(20, 1, D.Value, 8); features.Clear();
        Assert.Same(first, request.Spans[0]); Assert.Single(request.Features);
        Assert.Equal(4, Engine().Shape(request).Runs[0].FontSizeMillimeters);
        Assert.Throws<NotSupportedException>(() => ((IList<TextSizedFontSpan>)request.Spans).Clear());
        var huge = new HugeSpans();
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingRequest("A", 0, huge)).Error);
        Assert.Equal(0, huge.Reads);
        Assert.Throws<ArgumentNullException>(() => new TextSizedFontSpan(null!, 4));
        Assert.Throws<ArgumentNullException>(() => new TextSizedFontSpanShapingRequest(null!, 0, new[] { S(0, 1, D.Value, 4) }));
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            R("", 0, S(0, 1, D.Value, 4))).Error);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    [InlineData(1001)]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.Epsilon)]
    public void InvalidExplicitSizeIsSourceAddressedBeforeShaping(double size)
    {
        var error = Assert.Throws<TextLayoutException>(() => S(20, 1, D.Value, size));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error); Assert.Equal(20, error.SourceStart); Assert.Equal(1, error.SourceLength);
    }

    [Fact]
    public void SizeItemsConsumeCallBudgetBeforeAnyNativeCall()
    {
        var meter = new Meter(); var request = new TextSizedFontSpanShapingRequest("AV", 20,
            new[] { S(20, 1, D.Value, 4), S(21, 1, D.Value, 6) }, maximumShapeCalls: 1);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(meter).Shape(request)).Error);
        Assert.Equal(0, meter.Calls);
    }

    [Fact]
    public void SelectedShortRangeStillChargesCompleteOriginalContextForEverySizeGroup()
    {
        var bytes = File.ReadAllBytes(Font("DejaVuSans-R22.ttf")); Array.Resize(ref bytes, 504 * 1024 * 1024 / 64 / 2);
        var face = new TextFontFace("padded existing program", bytes);
        var original = R(new string('A', 4096), 0, Enumerable.Range(0, 64)
            .Select(i => S(i, i == 63 ? 4033 : 1, face, i % 2 == 0 ? 4 : 6)).ToArray());
        var item = TextSizedFontSpanShapingRequest.CreateContextualItem(original, new TextSourceSpan(0, 64));
        var meter = new Meter(r => Many(r, 1)); var engine = new TextSizedFontSpanShapingEngine(meter);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() => engine.Shape(item)).Error);
        Assert.Equal(0, meter.Calls);
        Assert.Equal(64, engine.Shape(new TextSizedFontSpanShapingRequest(item.Text, 0, item.Spans)).Runs.Count);
        Assert.Equal(64, meter.Calls);
    }

    [Fact]
    public void GlobalGlyphBudgetCannotBeResetAtSizeBoundaries()
    {
        var meter = new Meter(r => Many(r, 8193));
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(meter).Shape(R(new string('A', 4096), 0,
                S(0, 2048, D.Value, 4), S(2048, 2048, D.Value, 6)))).Error);
        Assert.Equal(2, meter.Calls);
    }

    [Fact]
    public void GlobalSliceBudgetCannotBeResetAtSizeBoundaries()
    {
        var meter = new Meter(r => Many(r, 129));
        var request = R(new string('A', 256), 0, Enumerable.Range(0, 256)
            .Select(i => S(i, 1, D.Value, i < 128 ? 4 : 6)).ToArray());
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(meter).Shape(request)).Error);
        Assert.Equal(2, meter.Calls);
    }

    [Fact]
    public void LaterProjectionBudgetFailureReturnsNoPartialSizedResult()
    {
        var meter = new Meter(r => Many(r, 1, int.MaxValue));
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(meter).Shape(R("AA", 0, S(0, 1, D.Value, .5), S(1, 1, D.Value, 1000)))).Error);
        Assert.Equal(2, meter.Calls);
    }

    [Fact]
    public void CancellationAndForeignRequestKeepExistingAtomicContracts()
    {
        var request = R("AV", 20, S(20, 1, D.Value, 4), S(21, 1, D.Value, 6));
        using var canceled = new CancellationTokenSource(); canceled.Cancel(); var idle = new Meter();
        var error = Assert.Throws<OperationCanceledException>(() => new TextSizedFontSpanShapingEngine(idle).Shape(request, canceled.Token));
        Assert.Equal(canceled.Token, error.CancellationToken); Assert.Equal(0, idle.Calls);
        using var during = new CancellationTokenSource(); var mid = new Meter(r => { during.Cancel(); return Many(r, 1); });
        Assert.Equal(during.Token, Assert.Throws<OperationCanceledException>(() =>
            new TextSizedFontSpanShapingEngine(mid).Shape(request, during.Token)).CancellationToken);
        Assert.Equal(1, mid.Calls);
        var foreign = new Meter(r => Many(new TextShapeRequest(r.Text, r.SourceStart, r.Face), 1));
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            new TextSizedFontSpanShapingEngine(foreign).Shape(request)).Error);
        Assert.Equal(1, foreign.Calls);
    }

    private static object F(TextFeature f) => (f.Tag, f.Value, f.Start, f.Length);
    private static TextSizedFontSpan S(int start, int length, TextFontFace face, double size, int key = 0) => new(new TextFontSpan(start, length, face, key), size);
    private static TextSizedFontSpanShapingRequest R(string text, int start, params TextSizedFontSpan[] spans) => new(text, start, spans);
    private static TextSizedFontSpanShapingEngine Engine() => HarfBuzzTextPrelayout.CreateSizedFontSpanShaping();
    private static TextShapedRun Many(TextShapeRequest request, int count, int advance = 1) => new(request,
        Enumerable.Range(0, count).Select(_ => new TextShapedGlyph(36, request.SourceStart, request.Text.Length, advance, 0, 0, 0, null, false)).ToArray(), "controlled-size-budget");
    private sealed class Meter(Func<TextShapeRequest, TextShapedRun>? custom = null) : IContextualTextShaper
    {
        internal int Calls { get; private set; }
        internal List<TextShapeRequest> Requests { get; } = new();
        public TextShapedRun Shape(TextShapeRequest r, CancellationToken cancellationToken = default) => throw new InvalidOperationException("Explicit contextual entry required.");
        public TextShapedRun ShapeContextualItem(TextShapeRequest r, CancellationToken cancellationToken = default)
        { Calls++; Requests.Add(r); return custom is null ? new HarfBuzzTextShaper().ShapeContextualItem(r, cancellationToken) : custom(r); }
    }
    private sealed class HugeSpans : IReadOnlyList<TextSizedFontSpan>
    {
        internal int Reads;
        public int Count => int.MaxValue;
        public TextSizedFontSpan this[int index] { get { Reads++; throw new InvalidOperationException(); } }
        public IEnumerator<TextSizedFontSpan> GetEnumerator() => throw new InvalidOperationException("Do not enumerate hostile input.");
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
    private static TextFontFace Load(string name, string? display = null) => new(display ?? name, File.ReadAllBytes(Font(name)));
    private static string Font(string name)
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        { var p = Path.Combine(directory.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", name); if (File.Exists(p)) return p; }
        throw new FileNotFoundException("Materialize existing deterministic font: " + name);
    }
}
