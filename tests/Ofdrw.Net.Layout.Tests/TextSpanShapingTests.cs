using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextSpanShapingTests
{
    private static readonly Lazy<TextFontFace> Face = new(() =>
        new TextFontFace("R22 fixed DejaVu", File.ReadAllBytes(Font())));

    [Fact]
    public void WholeRunKeepsFfiLigatureAcrossSourceSpansWithOneShapingCall()
    {
        var native = new HarfBuzzTextShaper();
        var shaper = new RecordingShaper(native.Shape);
        var run = Run("ffi", 20);
        var request = new TextSpanShapingRequest(run, new[] { new TextSourceSpan(20, 1), new TextSourceSpan(21, 2) });
        var result = new TextSpanShapingEngine(shaper).Shape(request);

        Assert.Equal(1, shaper.Calls);
        Assert.Same(run, shaper.LastRequest);
        var glyph = Assert.Single(result.ShapedRun.Glyphs);
        Assert.Equal(5044u, glyph.GlyphId); // Original DejaVu liga f+f+i.
        Assert.Equal(1980, glyph.XAdvance);
        var coverage = Assert.Single(result.GlyphCoverage);
        Assert.True(coverage.CrossesSpanBoundary);
        Assert.Equal(1, coverage.ClusterGlyphCount);
        Assert.Equal(new[] { (0, 20, 1), (1, 21, 2) }, Slices(coverage));
    }

    [Fact]
    public void RtlGlyphOrderDoesNotReorderLogicalSpanOwnership()
    {
        var request = new TextSpanShapingRequest(
            new TextShapeRequest("ببب", 10, Face.Value, TextDirection.RightToLeft, "Arab", "ar"),
            new[] { new TextSourceSpan(10, 1), new TextSourceSpan(11, 1), new TextSourceSpan(12, 1) });
        var result = HarfBuzzTextPrelayout.CreateSpanShaping().Shape(request);

        // Original GSUB final/medial/initial forms, in native visual order.
        Assert.Equal(new uint[] { 5258, 5260, 5259 }, result.ShapedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(3199, result.ShapedRun.XAdvance);
        Assert.Equal(new[] { 12, 11, 10 }, result.GlyphCoverage.Select(g => g.ClusterStart));
        Assert.Equal(new[] { 2, 1, 0 }, result.GlyphCoverage.Select(g => Assert.Single(g.Slices).SpanIndex));
        Assert.Equal(new[] { 0, 1, 2 }, result.GlyphCoverage.Select(g => g.GlyphIndex));
        Assert.All(result.GlyphCoverage, g => Assert.False(g.CrossesSpanBoundary));
    }

    [Fact]
    public void WholeRunRetainsCallerFeatureIntervalsAcrossTheSourcePartition()
    {
        var run = new TextShapeRequest("ffi", 20, Face.Value,
            features: new[] { new TextFeature("liga", 0, 0, 3) });
        var result = HarfBuzzTextPrelayout.CreateSpanShaping().Shape(new TextSpanShapingRequest(run,
            new[] { new TextSourceSpan(20, 1), new TextSourceSpan(21, 2) }));
        Assert.Same(run, result.ShapedRun.Request);
        Assert.Equal(new uint[] { 73, 73, 76 }, result.ShapedRun.Glyphs.Select(g => g.GlyphId));
        Assert.Equal(new[] { 0, 1, 1 }, result.GlyphCoverage.Select(g => Assert.Single(g.Slices).SpanIndex));
        Assert.All(result.GlyphCoverage, g => Assert.False(g.CrossesSpanBoundary));
    }

    [Fact]
    public void CombiningGlyphsRetainSharedClustersAndZeroAdvanceOccurrences()
    {
        var request = new TextSpanShapingRequest(Run("a\u0323\u0301", 30),
            new[] { new TextSourceSpan(30, 1), new TextSourceSpan(31, 1), new TextSourceSpan(32, 1) });
        var result = HarfBuzzTextPrelayout.CreateSpanShaping().Shape(request);

        Assert.True(result.ShapedRun.Glyphs.Count > 1);
        Assert.Contains(result.ShapedRun.Glyphs, glyph => glyph.XAdvance == 0);
        Assert.All(result.GlyphCoverage, coverage =>
        {
            Assert.Equal(30, coverage.ClusterStart);
            Assert.Equal(3, coverage.ClusterLength);
            Assert.Equal(result.ShapedRun.Glyphs.Count, coverage.ClusterGlyphCount);
            Assert.Equal(new[] { (0, 30, 1), (1, 31, 1), (2, 32, 1) }, Slices(coverage));
        });
    }

    [Fact]
    public void ManyGlyphsToManySpansKeepEachOccurrenceAndOriginalPosition()
    {
        var run = Run("ab", 100);
        var shaped = new TextShapedRun(run, new[] {
            Glyph(1, 100, 2, 100), new TextShapedGlyph(2, 100, 2, 0, 0, -13, 15, null, true)
        }, "controlled-shaper");
        var result = new TextSpanShapingEngine(new RecordingShaper((_, _) => shaped)).Shape(
            new TextSpanShapingRequest(run, new[] { new TextSourceSpan(100, 1), new TextSourceSpan(101, 1) }));

        Assert.Same(shaped, result.ShapedRun);
        Assert.Equal(0, result.ShapedRun.Glyphs[1].XAdvance);
        Assert.Equal(-13, result.ShapedRun.Glyphs[1].XOffset);
        Assert.Equal(15, result.ShapedRun.Glyphs[1].YOffset);
        Assert.Equal(new[] { 0, 1 }, result.GlyphCoverage.Select(g => g.GlyphIndex));
        Assert.All(result.GlyphCoverage, coverage =>
        {
            Assert.Equal(2, coverage.ClusterGlyphCount);
            Assert.True(coverage.CrossesSpanBoundary);
            Assert.Equal(new[] { (0, 100, 1), (1, 101, 1) }, Slices(coverage));
        });
    }

    [Theory]
    [InlineData("gap")]
    [InlineData("overlap")]
    [InlineData("before-run")]
    [InlineData("after-run")]
    [InlineData("missing-tail")]
    [InlineData("reordered")]
    public void InvalidPartitionsAreRejected(string defect)
    {
        var spans = defect switch {
            "gap" => new[] { new TextSourceSpan(10, 1), new TextSourceSpan(12, 2) },
            "overlap" => new[] { new TextSourceSpan(10, 3), new TextSourceSpan(12, 2) },
            "before-run" => new[] { new TextSourceSpan(9, 1), new TextSourceSpan(10, 4) },
            "after-run" => new[] { new TextSourceSpan(10, 5) },
            "missing-tail" => new[] { new TextSourceSpan(10, 3) },
            _ => new[] { new TextSourceSpan(12, 2), new TextSourceSpan(10, 2) }
        };
        var error = Assert.Throws<TextLayoutException>(() => new TextSpanShapingRequest(Run("abcd", 10), spans));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Theory]
    [InlineData(-1, 1)]
    [InlineData(0, 0)]
    [InlineData(1, -1)]
    [InlineData(int.MaxValue, 1)]
    public void InvalidSpanArithmeticCannotWrap(int start, int length)
        => Assert.Throws<ArgumentOutOfRangeException>(() => new TextSourceSpan(start, length));

    [Fact]
    public void SourcePartitionCannotSplitSurrogatesButMayKeepTheirCompleteRange()
    {
        var run = Run("A\U00010400B", 10);
        var error = Assert.Throws<TextLayoutException>(() => new TextSpanShapingRequest(run,
            new[] { new TextSourceSpan(10, 2), new TextSourceSpan(12, 2) }));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
        var valid = new TextSpanShapingRequest(run,
            new[] { new TextSourceSpan(10, 1), new TextSourceSpan(11, 2), new TextSourceSpan(13, 1) });
        Assert.Equal(2, valid.Spans[1].SourceLength);
    }

    [Fact]
    public void CallerListMutationCannotChangeTheInputOrCompletedCoverage()
    {
        var run = Run("ab");
        var original = new List<TextSourceSpan> { new(0, 1), new(1, 1) };
        var request = new TextSpanShapingRequest(run, original);
        original.Clear();
        var result = new TextSpanShapingEngine(new RecordingShaper((r, _) =>
            new TextShapedRun(r, new[] { Glyph(1, 0, 2) }, "controlled"))).Shape(request);
        Assert.Equal(2, request.Spans.Count);
        Assert.Equal(new[] { (0, 0, 1), (1, 1, 1) }, Slices(Assert.Single(result.GlyphCoverage)));
        Assert.Throws<NotSupportedException>(() => ((IList<TextSourceSpan>)request.Spans).Clear());
        Assert.Throws<NotSupportedException>(() => ((IList<TextSpanSourceSlice>)result.GlyphCoverage[0].Slices).Clear());
    }

    [Fact]
    public void EmptyRunHasAnEmptyPartitionAndNoGlyphAssociations()
    {
        var run = Run(string.Empty, 50);
        var shaper = new RecordingShaper((r, _) => new TextShapedRun(r, Array.Empty<TextShapedGlyph>(), "empty"));
        var result = new TextSpanShapingEngine(shaper).Shape(new TextSpanShapingRequest(run, Array.Empty<TextSourceSpan>()));
        Assert.Equal(1, shaper.Calls);
        Assert.Empty(result.GlyphCoverage);
        Assert.Empty(result.ShapedRun.Glyphs);
    }

    [Fact]
    public void SpanCountBudgetIsCheckedBeforeAnyShaping()
    {
        var spans = Enumerable.Range(0, 257).Select(i => new TextSourceSpan(i, 1)).ToArray();
        var error = Assert.Throws<TextLayoutException>(() => new TextSpanShapingRequest(Run(new string('a', 257)), spans));
        Assert.Equal(TextLayoutError.BudgetExceeded, error.Error);
    }

    [Fact]
    public void ManyToManyAssociationBudgetAcceptsBoundaryAndRefusesNextOccurrence()
    {
        var run = Run(new string('a', 256));
        var request = new TextSpanShapingRequest(run,
            Enumerable.Range(0, 256).Select(i => new TextSourceSpan(i, 1)).ToArray());
        TextShapedRun Shape(int count) => new(run,
            Enumerable.Range(0, count).Select(_ => Glyph(1, 0, 256)).ToArray(), "controlled");
        var accepted = new TextSpanShapingEngine(new RecordingShaper((_, _) => Shape(128))).Shape(request);
        Assert.Equal(32768, accepted.GlyphCoverage.Sum(g => g.Slices.Count));
        var error = Assert.Throws<TextLayoutException>(() =>
            new TextSpanShapingEngine(new RecordingShaper((_, _) => Shape(129))).Shape(request));
        Assert.Equal(TextLayoutError.BudgetExceeded, error.Error);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void CancellationBeforeAndAfterShapingReturnsNoAssociationResult(bool before)
    {
        using var cancelled = new CancellationTokenSource();
        var run = Run("a");
        var request = new TextSpanShapingRequest(run, new[] { new TextSourceSpan(0, 1) });
        var shaper = new RecordingShaper((r, _) => {
            cancelled.Cancel();
            return new TextShapedRun(r, new[] { Glyph(1, 0, 1) }, "controlled");
        });
        if (before) cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => new TextSpanShapingEngine(shaper).Shape(request, cancelled.Token));
        Assert.Equal(before ? 0 : 1, shaper.Calls);
    }

    [Fact]
    public void ReplacementShaperCannotReturnAnotherOtherwiseValidRequest()
    {
        var run = Run("a");
        var shaper = new RecordingShaper((_, _) => new TextShapedRun(Run("a"), new[] { Glyph(1, 0, 1) }, "foreign-request"));
        var error = Assert.Throws<TextLayoutException>(() => new TextSpanShapingEngine(shaper).Shape(
            new TextSpanShapingRequest(run, new[] { new TextSourceSpan(0, 1) })));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Fact]
    public void SourceEndpointAtIntMaximumDoesNotOverflowAssociation()
    {
        var run = Run("ab", int.MaxValue - 2);
        var request = new TextSpanShapingRequest(run, new[] {
            new TextSourceSpan(int.MaxValue - 2, 1), new TextSourceSpan(int.MaxValue - 1, 1) });
        var result = new TextSpanShapingEngine(new RecordingShaper((r, _) => new TextShapedRun(r,
            new[] { Glyph(1, int.MaxValue - 2, 2) }, "controlled"))).Shape(request);
        Assert.Equal(new[] { (0, int.MaxValue - 2, 1), (1, int.MaxValue - 1, 1) },
            Slices(Assert.Single(result.GlyphCoverage)));
    }

    private static TextShapeRequest Run(string text, int start = 0) => new(text, start, Face.Value);
    private static TextShapedGlyph Glyph(uint id, int start, int length, int advance = 1)
        => new(id, start, length, advance, 0, 0, 0, null, false);
    private static (int, int, int)[] Slices(TextSpanGlyphCoverage coverage)
        => coverage.Slices.Select(s => (s.SpanIndex, s.SourceStart, s.SourceLength)).ToArray();

    private sealed class RecordingShaper : ITextShaper
    {
        private readonly Func<TextShapeRequest, CancellationToken, TextShapedRun> _shape;
        internal RecordingShaper(Func<TextShapeRequest, CancellationToken, TextShapedRun> shape) => _shape = shape;
        internal int Calls { get; private set; }
        internal TextShapeRequest? LastRequest { get; private set; }
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
        { Calls++; LastRequest = request; return _shape(request, cancellationToken); }
    }

    private static string Font()
    {
        for (var directory = new DirectoryInfo(AppContext.BaseDirectory); directory is not null; directory = directory.Parent)
        {
            var path = Path.Combine(directory.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/DejaVuSans-R22.ttf");
            if (File.Exists(path)) return path;
        }
        throw new FileNotFoundException("Materialize the existing DejaVu R22 fixture.");
    }
}
