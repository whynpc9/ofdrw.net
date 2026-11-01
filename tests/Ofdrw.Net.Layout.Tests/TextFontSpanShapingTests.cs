using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextFontSpanShapingTests
{
    private static readonly Lazy<TextFontFace> DejaVu = new(() => Load("DejaVuSans-R22.ttf"));
    private static readonly Lazy<TextFontFace> Noto = new(() => Load("NotoSans-Regular.ttf"));
    private static readonly Lazy<TextFontFace> WenKai = new(() => Load("LXGWWenKai-Regular.ttf"));

    [Fact]
    public void SameFaceCompatibleSpansKeepWholeFfiInsteadOfFragmentShaping()
    {
        var spans = new[] { Span(20, 1, DejaVu.Value, 7), Span(21, 2, DejaVu.Value, 7) };
        var request = new TextFontSpanShapingRequest("ffi", 20, spans);
        var recorder = NativeRecorder();
        var result = new TextFontSpanShapingEngine(recorder).Shape(request);
        var run = Assert.Single(result.Runs);
        Assert.Same(request, result.Request);
        Assert.Equal(1, recorder.Calls);
        Assert.Equal(2, run.SpanCount);
        Assert.Equal(5044u, Assert.Single(run.Shaping.ShapedRun.Glyphs).GlyphId);
        Assert.Equal(1980, run.Shaping.ShapedRun.XAdvance);
        Assert.Equal(new[] { 7 }, run.GlyphStyleKeys);
        var native = new HarfBuzzTextShaper();
        Assert.Equal(2011, native.Shape(new TextShapeRequest("f", 20, DejaVu.Value)).XAdvance +
            native.Shape(new TextShapeRequest("fi", 21, DejaVu.Value)).XAdvance);
    }

    [Fact]
    public void EqualProgramAliasesCoalesceWithoutLosingOriginalFaceAssociations()
    {
        var alias = Load("DejaVuSans-R22.ttf", "different display name");
        Assert.NotSame(DejaVu.Value, alias);
        Assert.Equal(DejaVu.Value.FaceSha256, alias.FaceSha256);
        var result = Engine().Shape(new TextFontSpanShapingRequest("ffi", 0,
            new[] { Span(0, 1, DejaVu.Value), Span(1, 2, alias) }));
        var run = Assert.Single(result.Runs);
        Assert.Equal(5044u, Assert.Single(run.Shaping.ShapedRun.Glyphs).GlyphId);
        Assert.Same(DejaVu.Value, result.Request.Spans[0].Face);
        Assert.Same(alias, result.Request.Spans[1].Face);
        Assert.Same(DejaVu.Value, run.Shaping.ShapedRun.Request.Face);
    }

    [Fact]
    public void ConflictingStyleKeysCannotChooseLeadingStyleForOneFfiOccurrence()
    {
        var request = new TextFontSpanShapingRequest("ffi", 20,
            new[] { Span(20, 1, DejaVu.Value, 7), Span(21, 2, DejaVu.Value, 9) });
        var error = Assert.Throws<TextLayoutException>(() => Engine().Shape(request));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
        Assert.Equal(20, error.SourceStart);
        Assert.Equal(3, error.SourceLength);
    }

    [Fact]
    public void ExplicitLigaOffPermitsStylesOnlyWhenActualClustersSplit()
    {
        var request = new TextFontSpanShapingRequest("ffi", 20,
            new[] { Span(20, 1, DejaVu.Value, 7), Span(21, 2, DejaVu.Value, 9) },
            features: new[] { new TextFeature("liga", 0, 0, 3) });
        var run = Assert.Single(Engine().Shape(request).Runs);
        Assert.Equal(new uint[] { 73, 73, 76 }, run.Shaping.ShapedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new[] { 7, 9, 9 }, run.GlyphStyleKeys);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CombiningSharedClusterKeepsOccurrencesAndRefusesConflictingKeysEvenWithLigaOff(bool conflict)
    {
        var request = new TextFontSpanShapingRequest("a\u0323\u0301", 30,
            new[] { Span(30, 1, DejaVu.Value, 7), Span(31, 2, DejaVu.Value, conflict ? 9 : 7) },
            features: new[] { new TextFeature("liga", 0, 0, 3) });
        if (conflict)
        {
            var error = Assert.Throws<TextLayoutException>(() => Engine().Shape(request));
            Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
            Assert.Equal(30, error.SourceStart); Assert.Equal(3, error.SourceLength);
            return;
        }
        var run = Assert.Single(Engine().Shape(request).Runs);
        Assert.True(run.Shaping.ShapedRun.Glyphs.Count > 1);
        Assert.Contains(run.Shaping.ShapedRun.Glyphs, g => g.XAdvance == 0);
        Assert.All(run.GlyphStyleKeys, key => Assert.Equal(7, key));
        Assert.All(run.Shaping.GlyphCoverage, c => Assert.Equal(run.Shaping.ShapedRun.Glyphs.Count, c.ClusterGlyphCount));
        var original = new HarfBuzzTextShaper().Shape(new TextShapeRequest(request.Text, request.SourceStart,
            DejaVu.Value, features: request.Features));
        Assert.Equal(original.Glyphs.Select(Identity), run.Shaping.ShapedRun.Glyphs.Select(Identity));
    }

    [Fact]
    public void ArabicStyleKeysFollowSourceCoverageWithoutBreakingJoiningOrNativeOrder()
    {
        var request = new TextFontSpanShapingRequest("ببب", 100,
            new[] { Span(100, 1, DejaVu.Value, 10), Span(101, 1, DejaVu.Value, 20), Span(102, 1, DejaVu.Value, 30) },
            TextDirection.RightToLeft, "Arab", "ar");
        var run = Assert.Single(Engine().Shape(request).Runs);
        Assert.Equal(new uint[] { 5258, 5260, 5259 }, run.Shaping.ShapedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new[] { 102, 101, 100 }, run.Shaping.ShapedRun.Glyphs.Select(g => g.ClusterStart));
        Assert.Equal(new[] { 30, 20, 10 }, run.GlyphStyleKeys);
    }

    [Fact]
    public void DifferentProgramsWithTheSameDisplayNameKeepOwnGidsUnitsAndLogicalItems()
    {
        var a = Load("NotoSans-Regular.ttf", "same name");
        var b = Load("DejaVuSans-R22.ttf", "same name");
        var result = Engine().Shape(new TextFontSpanShapingRequest("AV", 50,
            new[] { Span(50, 1, a, 1), Span(51, 1, b, 2) }));
        Assert.Equal(2, result.Runs.Count);
        Assert.Equal(new[] { 0, 1 }, result.Runs.Select(r => r.FirstSpanIndex));
        Assert.Equal(new[] { 1000, 2048 }, result.Runs.Select(r => r.Shaping.ShapedRun.Request.Face.UnitsPerEm));
        Assert.Equal(new uint[] { 36, 57 }, result.Runs.Select(r => Assert.Single(r.Shaping.ShapedRun.Glyphs).GlyphId));
        Assert.Equal(new[] { 50, 51 }, result.Runs.Select(r => Assert.Single(r.Shaping.ShapedRun.Glyphs).ClusterStart));
        Assert.Same(a, result.Runs[0].Shaping.ShapedRun.Request.Face);
        Assert.Same(b, result.Runs[1].Shaping.ShapedRun.Request.Face);
        Assert.All(result.Runs, r => Assert.Equal("AV", r.Shaping.Request.Run.Context!.Text));
    }

    [Fact]
    public void AFontNeedNotCoverExternalContextOwnedByAnotherFont()
    {
        Assert.False(Noto.Value.Covers('中'));
        var result = Engine().Shape(new TextFontSpanShapingRequest("中A", 10,
            new[] { Span(10, 1, WenKai.Value, 1), Span(11, 1, Noto.Value, 2) }, script: "Hani", language: "zh"));
        Assert.Equal(2, result.Runs.Count);
        Assert.Equal(36u, Assert.Single(result.Runs[1].Shaping.ShapedRun.Glyphs).GlyphId);
        Assert.Equal(11, Assert.Single(result.Runs[1].Shaping.ShapedRun.Glyphs).ClusterStart);
        Assert.Same(Noto.Value, result.Runs[1].Shaping.ShapedRun.Request.Face);
    }

    [Fact]
    public void RtlFontItemsStayLogicalInsteadOfClaimingFinalVisualLineOrder()
    {
        var result = Engine().Shape(new TextFontSpanShapingRequest("Aא", 20,
            new[] { Span(20, 1, Noto.Value, 1), Span(21, 1, DejaVu.Value, 2) }, TextDirection.RightToLeft, "Hebr", "he"));
        Assert.Equal(new[] { 0, 1 }, result.Runs.Select(r => r.FirstSpanIndex));
        Assert.Equal(new[] { 20, 21 }, result.Runs.Select(r => r.Shaping.ShapedRun.Request.SourceStart));
        Assert.All(result.Runs, r => Assert.Equal(TextDirection.RightToLeft, r.Shaping.ShapedRun.Request.Direction));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FeatureIntersectionsKeepFullSourceCoordinatesAcrossFontSwitch(bool outside)
    {
        var request = new TextFontSpanShapingRequest("Afi", 70,
            new[] { Span(70, 1, Noto.Value), Span(71, 2, DejaVu.Value) },
            features: new[] { new TextFeature("liga", 0, outside ? 0 : 1, outside ? 1 : 2) });
        var result = Engine().Shape(request);
        Assert.Equal(outside ? new uint[] { 5042 } : new uint[] { 73, 76 }, result.Runs[1].Shaping.ShapedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(71, result.Runs[1].Shaping.ShapedRun.Glyphs[0].ClusterStart);
    }

    [Fact]
    public void LaterFontBoundaryInsideBaseMarkRejectsBeforeAnyNativeCall()
    {
        var recorder = NativeRecorder();
        var request = new TextFontSpanShapingRequest("Xa\u0323", 0,
            new[] { Span(0, 1, Noto.Value), Span(1, 1, DejaVu.Value), Span(2, 1, Noto.Value) });
        var error = Assert.Throws<TextLayoutException>(() => new TextFontSpanShapingEngine(recorder).Shape(request));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
        Assert.Equal(0, recorder.Calls);
    }

    [Fact]
    public void CallerListMutationCannotChangeInputFacesFeaturesOrCompletedOwnership()
    {
        var spans = new List<TextFontSpan> { Span(0, 1, DejaVu.Value, 1), Span(1, 2, DejaVu.Value, 2) };
        var features = new List<TextFeature> { new("liga", 0, 0, 3) };
        var request = new TextFontSpanShapingRequest("ffi", 0, spans, features: features);
        spans.Clear(); features.Clear();
        var result = Engine().Shape(request);
        Assert.Equal(2, request.Spans.Count);
        Assert.Single(request.Features);
        Assert.Equal(new[] { 1, 2, 2 }, Assert.Single(result.Runs).GlyphStyleKeys);
        Assert.Throws<NotSupportedException>(() => ((IList<TextFontSpan>)request.Spans).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<int>)result.Runs[0].GlyphStyleKeys).Clear());
    }

    [Fact]
    public void ForeignContextualRequestIdentityCannotPassThroughFontQuery()
    {
        var recorder = new Recorder((r, token) => {
            var c = r.Context!;
            var foreignContext = new TextShapeRequest(c.Text, c.SourceStart, c.Face, c.Direction, c.Script, c.Language, c.Features);
            var foreignItem = TextShapeRequest.CreateContextualItem(foreignContext, new TextSourceSpan(r.SourceStart, r.Text.Length));
            return new HarfBuzzTextShaper().ShapeContextualItem(foreignItem, token);
        });
        var error = Assert.Throws<TextLayoutException>(() => new TextFontSpanShapingEngine(recorder).Shape(
            new TextFontSpanShapingRequest("A", 0, new[] { Span(0, 1, Noto.Value) })));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CancellationBeforeOrAfterFirstItemReturnsNoPartialQuery(bool before)
    {
        using var cancel = new CancellationTokenSource();
        var recorder = new Recorder((r, token) => {
            var shaped = new HarfBuzzTextShaper().ShapeContextualItem(r, token);
            cancel.Cancel(); return shaped;
        });
        if (before) cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => new TextFontSpanShapingEngine(recorder).Shape(
            new TextFontSpanShapingRequest("AV", 0, new[] { Span(0, 1, Noto.Value), Span(1, 1, DejaVu.Value) }), cancel.Token));
        Assert.Equal(before ? 0 : 1, recorder.Calls);
    }

    [Fact]
    public void AllShapeCallAndWorkBudgetsPreflightBeforeAnyNativeCall()
    {
        var recorder = NativeRecorder();
        var calls = new TextFontSpanShapingRequest("AAA", 0,
            new[] { Span(0, 1, Noto.Value), Span(1, 1, DejaVu.Value), Span(2, 1, Noto.Value) }, maximumShapeCalls: 2);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() => new TextFontSpanShapingEngine(recorder).Shape(calls)).Error);
        var large = PaddedFace(4 * 1024 * 1024);
        var request = new TextFontSpanShapingRequest(new string('A', 256), 0,
            Enumerable.Range(0, 256).Select(i => Span(i, 1, i % 2 == 0 ? large : DejaVu.Value)).ToArray());
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() => new TextFontSpanShapingEngine(recorder).Shape(request)).Error);
        Assert.Equal(0, recorder.Calls);
    }

    [Fact]
    public void InputSpanFeatureAndDistinctFaceBudgetsRejectBeforeShaping()
    {
        var spans = Enumerable.Range(0, 257).Select(i => Span(i, 1, DejaVu.Value)).ToArray();
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextFontSpanShapingRequest(new string('A', 257), 0, spans)).Error);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextFontSpanShapingRequest("A", 0, new[] { spans[0] }, features: Enumerable.Range(0, 257)
                .Select(_ => new TextFeature("kern", 1, 0, 1)).ToArray())).Error);
        // Appending aligned zero words preserves table contents and sfnt checksum,
        // but gives distinct immutable face-program byte identities. No new files.
        var variants = Enumerable.Range(1, 9).Select(i => PaddedFace(0, i)).ToArray();
        var selected = variants.Select((f, i) => Span(i, 1, f)).ToArray();
        _ = new TextFontSpanShapingRequest(new string('A', 8), 0, selected.Take(8).ToArray());
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextFontSpanShapingRequest(new string('A', 9), 0, selected)).Error);
    }

    [Fact]
    public void SelectedFaceByteBudgetAccepts64MiBAndRefusesNextAlignedWord()
    {
        var first = WenKai.Value;
        var second = PaddedFace(0, 1, "LXGWWenKai-Regular.ttf");
        var length = new FileInfo(Font("LXGWWenKai-Regular.ttf")).Length;
        var remaining = (int)(64L * 1024 * 1024 - length * 2 - 4);
        var boundary = PaddedFace(remaining);
        _ = new TextFontSpanShapingRequest("AAA", 0,
            new[] { Span(0, 1, first), Span(1, 1, second), Span(2, 1, boundary) });
        var overflow = PaddedFace(remaining + 4);
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextFontSpanShapingRequest("AAA", 0,
                new[] { Span(0, 1, first), Span(1, 1, second), Span(2, 1, overflow) })).Error);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AggregateGlyphBudgetAcceptsExactBoundaryAndRefusesNextOccurrence(bool overflow)
    {
        var request = new TextFontSpanShapingRequest(new string('A', 4096), 0,
            new[] { Span(0, 2048, Noto.Value), Span(2048, 2048, DejaVu.Value) });
        var fake = new Recorder((r, _) => Many(r, 8192 + (overflow && r.SourceStart == 2048 ? 1 : 0)));
        if (overflow) Assert.Equal(TextLayoutError.BudgetExceeded,
            Assert.Throws<TextLayoutException>(() => new TextFontSpanShapingEngine(fake).Shape(request)).Error);
        else Assert.Equal(16384, new TextFontSpanShapingEngine(fake).Shape(request).Runs.Sum(r => r.GlyphStyleKeys.Count));
        Assert.Equal(2, fake.Calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void AggregateSourceSliceBudgetCountsRepeatedSharedClusterOccurrences(bool overflow)
    {
        var request = new TextFontSpanShapingRequest(new string('A', 256), 0,
            Enumerable.Range(0, 256).Select(i => Span(i, 1, i < 128 ? Noto.Value : DejaVu.Value)).ToArray());
        var fake = new Recorder((r, _) => Many(r, 128 + (overflow && r.SourceStart == 128 ? 1 : 0)));
        if (overflow) Assert.Equal(TextLayoutError.BudgetExceeded,
            Assert.Throws<TextLayoutException>(() => new TextFontSpanShapingEngine(fake).Shape(request)).Error);
        else Assert.Equal(32768, new TextFontSpanShapingEngine(fake).Shape(request).Runs.Sum(r => r.Shaping.GlyphCoverage.Sum(g => g.Slices.Count)));
        Assert.Equal(2, fake.Calls);
    }

    [Fact]
    public void IntegerMaximumEndpointKeepsEveryGroupAbsoluteInterval()
    {
        var start = int.MaxValue - 2;
        var result = Engine().Shape(new TextFontSpanShapingRequest("AV", start,
            new[] { Span(start, 1, Noto.Value), Span(start + 1, 1, DejaVu.Value) }));
        Assert.Equal(new[] { start, start + 1 }, result.Runs.Select(r => Assert.Single(r.Shaping.ShapedRun.Glyphs).ClusterStart));
    }

    [Fact]
    public void NonBmpItemKeepsUtf16LengthAcrossAFontSwitch()
    {
        var result = Engine().Shape(new TextFontSpanShapingRequest("\U00010300A", 20,
            new[] { Span(20, 2, DejaVu.Value), Span(22, 1, Noto.Value) }));
        var first = Assert.Single(result.Runs[0].Shaping.ShapedRun.Glyphs);
        Assert.Equal(20, first.ClusterStart); Assert.Equal(2, first.ClusterLength);
        Assert.Equal(22, Assert.Single(result.Runs[1].Shaping.ShapedRun.Glyphs).ClusterStart);
    }

    private static TextFontSpan Span(int start, int length, TextFontFace face, int key = 0) => new(start, length, face, key);
    private static TextFontSpanShapingEngine Engine() => HarfBuzzTextPrelayout.CreateFontSpanShaping();
    private static Recorder NativeRecorder() => new((r, token) => new HarfBuzzTextShaper().ShapeContextualItem(r, token));
    private static TextShapedRun Many(TextShapeRequest r, int count) => new(r,
        Enumerable.Range(0, count).Select(_ => new TextShapedGlyph(1, r.SourceStart, r.Text.Length, 1, 0, 0, 0, null, false)).ToArray(), "controlled-MN");
    private static TextFontFace PaddedFace(int minimumBytes, int words = 0, string file = "NotoSans-Regular.ttf")
    {
        var bytes = File.ReadAllBytes(Font(file));
        Array.Resize(ref bytes, Math.Max(bytes.Length, minimumBytes) + words * 4);
        return new TextFontFace("same display name", bytes);
    }
    private static TextFontFace Load(string file, string? name = null) => new(name ?? file, File.ReadAllBytes(Font(file)));
    private static object Identity(TextShapedGlyph g) => (g.GlyphId, g.ClusterStart, g.ClusterLength,
        g.XAdvance, g.YAdvance, g.XOffset, g.YOffset, g.UnsafeToBreak);
    private sealed class Recorder(Func<TextShapeRequest, CancellationToken, TextShapedRun> shape) : IContextualTextShaper
    {
        internal int Calls { get; private set; }
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
            => throw new InvalidOperationException("Font partition query must use explicit contextual entry point.");
        public TextShapedRun ShapeContextualItem(TextShapeRequest request, CancellationToken cancellationToken = default)
        { Calls++; return shape(request, cancellationToken); }
    }
    private static string Font(string name)
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", name);
            if (File.Exists(path)) return path;
        }
        throw new FileNotFoundException("Materialize the existing deterministic fixture: " + name);
    }
}
