using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextContextShapingTests
{
    private static readonly Lazy<TextFontFace> Face = new(() =>
        new TextFontFace("R22 fixed DejaVu", File.ReadAllBytes(Font())));

    [Theory]
    [InlineData("ببب", 5260u)]
    [InlineData(" بب", 5259u)]
    [InlineData("بب ", 5258u)]
    [InlineData(" ب ", 1366u)]
    public void ActualArabicItemUsesBothExternalContextsWithoutEmittingThem(string text, uint gid)
    {
        var context = Arabic(text, 100);
        var item = TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(101, 1));
        var shaped = new HarfBuzzTextShaper().ShapeContextualItem(item);
        Assert.Same(context, item.Context);
        Assert.Same(context.Face, item.Face);
        Assert.Same(item, shaped.Request);
        Assert.Equal("ب", item.Text);
        var glyph = Assert.Single(shaped.Glyphs);
        Assert.Equal(gid, glyph.GlyphId);
        Assert.Equal(101, glyph.ClusterStart);
        Assert.Equal(1, glyph.ClusterLength);
        Assert.Equal(1366u, Assert.Single(new HarfBuzzTextShaper().Shape(Arabic("ب", 101)).Glyphs).GlyphId);
    }

    [Theory]
    [InlineData(0, 5259u)]
    [InlineData(2, 5258u)]
    public void FirstAndLastRtlItemsKeepAbsoluteEndpoints(int offset, uint gid)
    {
        var item = TextShapeRequest.CreateContextualItem(Arabic("ببب", 20), new TextSourceSpan(20 + offset, 1));
        var glyph = Assert.Single(new HarfBuzzTextShaper().ShapeContextualItem(item).Glyphs);
        Assert.Equal(gid, glyph.GlyphId);
        Assert.Equal(20 + offset, glyph.ClusterStart);
        Assert.Equal(1, glyph.ClusterLength);
    }

    [Fact]
    public void NativeUtf16OffsetsAreNotScalarOrAbsoluteSourceOffsets()
    {
        var item = TextShapeRequest.CreateContextualItem(Arabic("\U00010400ببب", 200), new TextSourceSpan(203, 1));
        var glyph = Assert.Single(new HarfBuzzTextShaper().ShapeContextualItem(item).Glyphs);
        Assert.Equal(5260u, glyph.GlyphId);
        Assert.Equal(203, glyph.ClusterStart);
        Assert.Equal(1, glyph.ClusterLength);
    }

    [Fact]
    public void ContextIsNotEquivalentToCroppingAWholeRunLigature()
    {
        var context = new TextShapeRequest("ffi", 20, Face.Value);
        var item = TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(21, 2));
        Assert.Equal(5044u, Assert.Single(new HarfBuzzTextShaper().Shape(context).Glyphs).GlyphId);
        var shaped = new HarfBuzzTextShaper().ShapeContextualItem(item);
        var glyph = Assert.Single(shaped.Glyphs);
        Assert.Equal(5042u, glyph.GlyphId); // fi, not a clipped/shared ffi.
        Assert.Equal(21, glyph.ClusterStart);
        Assert.Equal(2, glyph.ClusterLength);
        var query = HarfBuzzTextPrelayout.CreateSpanShaping().Shape(new TextSpanShapingRequest(item,
            new[] { new TextSourceSpan(21, 1), new TextSourceSpan(22, 1) }));
        Assert.True(Assert.Single(query.GlyphCoverage).CrossesSpanBoundary);
        Assert.Same(context, query.Request.Run.Context);
    }

    [Theory]
    [InlineData(0, 1, true)]
    [InlineData(1, 2, false)]
    [InlineData(0, 3, false)]
    [InlineData(2, 1, false)]
    public void ContextFeatureIntersectionsAreRebasedBeforeNativeClusterTranslation(int start, int length, bool ligature)
    {
        var feature = new TextFeature("liga", 0, start, length);
        var context = new TextShapeRequest("ffi", 50, Face.Value, features: new[] { feature });
        var item = TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(51, 2));
        Assert.Equal(ligature ? new uint[] { 5042 } : new uint[] { 73, 76 },
            new HarfBuzzTextShaper().ShapeContextualItem(item).Glyphs.Select(g => g.GlyphId));
        Assert.Same(feature, Assert.Single(context.Features));
        if (ligature) Assert.Empty(item.Features);
        else
        {
            var mapped = Assert.Single(item.Features);
            Assert.Equal(Math.Max(start, 1) - 1, mapped.Start);
            Assert.Equal(Math.Min(start + length, 3) - Math.Max(start, 1), mapped.Length);
        }
    }

    [Fact]
    public void FullRangeItemKeepsEveryOriginalGlyphAndPosition()
    {
        var context = Arabic("ببب", 10);
        var shaper = new HarfBuzzTextShaper();
        var whole = shaper.Shape(context);
        var item = shaper.ShapeContextualItem(TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(10, 3)));
        Assert.Equal(whole.Glyphs.Select(Identity), item.Glyphs.Select(Identity));
        Assert.Equal(whole.XAdvance, item.XAdvance);
    }

    [Theory]
    [InlineData("a\u0323\u0301", 0, 1)]
    [InlineData("a\u0323\u0301", 1, 2)]
    [InlineData("A\U00010400B", 1, 1)]
    [InlineData("A\U00010400B", 2, 1)]
    public void ContextItemsCannotSplitTextElementsOrSurrogatePairs(string text, int start, int length)
    {
        var error = Assert.Throws<TextLayoutException>(() => TextShapeRequest.CreateContextualItem(
            new TextShapeRequest(text, 10, Face.Value), new TextSourceSpan(10 + start, length)));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
        Assert.Equal(10 + start, error.SourceStart);
        Assert.Equal(length, error.SourceLength);
    }

    [Fact]
    public void WholeCombiningSequenceRemainsOneSourceIntervalWithAllOccurrences()
    {
        var item = TextShapeRequest.CreateContextualItem(new TextShapeRequest("Xa\u0323\u0301Y", 30, Face.Value),
            new TextSourceSpan(31, 3));
        var shaped = new HarfBuzzTextShaper().ShapeContextualItem(item);
        Assert.True(shaped.Glyphs.Count > 1);
        Assert.Contains(shaped.Glyphs, g => g.XAdvance == 0);
        Assert.All(shaped.Glyphs, g => { Assert.Equal(31, g.ClusterStart); Assert.Equal(3, g.ClusterLength); });
    }

    [Theory]
    [InlineData(9, 1)]
    [InlineData(12, 2)]
    public void OutsideItemsRejectWithoutReinterpretingOffsets(int start, int length)
    {
        var error = Assert.Throws<TextLayoutException>(() => TextShapeRequest.CreateContextualItem(
            new TextShapeRequest("abc", 10, Face.Value), new TextSourceSpan(start, length)));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Fact]
    public void NestedItemsRejectAndIntegerMaximumEndpointIsSafe()
    {
        var context = Arabic("ببب", int.MaxValue - 3);
        var item = TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(int.MaxValue - 1, 1));
        var glyph = Assert.Single(new HarfBuzzTextShaper().ShapeContextualItem(item).Glyphs);
        Assert.Equal(int.MaxValue - 1, glyph.ClusterStart);
        var error = Assert.Throws<TextLayoutException>(() => TextShapeRequest.CreateContextualItem(item,
            new TextSourceSpan(item.SourceStart, 1)));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Fact]
    public void CancelledContextItemDoesNotEnterNativeShaping()
    {
        var item = TextShapeRequest.CreateContextualItem(Arabic("ببب", 20), new TextSourceSpan(21, 1));
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Throws<OperationCanceledException>(() => new HarfBuzzTextShaper().ShapeContextualItem(item, cancel.Token));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExistingLayoutEnginesRejectContextualReplacementRequestsBeforeWriterAdmission(bool mixed)
    {
        var shaper = new ContextReplacingShaper();
        var error = Assert.Throws<TextLayoutException>(() =>
        {
            if (mixed) new TextMixedParagraphEngine(shaper).Layout(
                new TextMixedParagraphRequest("abc שלום def", Face.Value, 4, 100));
            else new TextPrelayoutEngine(shaper).Layout(new TextPrelayoutRequest("abc", new[] { Face.Value }, 4, 100));
        });
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Fact]
    public void CompleteContextMustAlreadyPassSourceBudgetAndControlValidation()
    {
        Assert.Equal(TextLayoutError.BudgetExceeded, Assert.Throws<TextLayoutException>(() =>
            new TextShapeRequest(new string('a', 4097), 0, Face.Value)).Error);
        Assert.Equal(TextLayoutError.InvalidText, Assert.Throws<TextLayoutException>(() =>
            new TextShapeRequest("a\nb", 0, Face.Value)).Error);
        var context = new TextShapeRequest(new string('a', 4096), 0, Face.Value);
        var item = TextShapeRequest.CreateContextualItem(context, new TextSourceSpan(4095, 1));
        Assert.Same(context, item.Context);
        Assert.Equal(4096, item.Context!.Text.Length);
        Assert.Equal(4095, Assert.Single(new HarfBuzzTextShaper().ShapeContextualItem(item).Glyphs).ClusterStart);
    }

    [Fact]
    public void OrdinaryOnlyReplacementShaperCannotSilentlyDiscardExternalContextInSpanQuery()
    {
        var item = TextShapeRequest.CreateContextualItem(Arabic("ببب", 10), new TextSourceSpan(11, 1));
        var legacy = new LegacySubstringShaper();
        // This previously conforming implementation knows only request.Text.
        // It even returns the exact request; reference identity cannot detect
        // its lost context. The new engine must refuse before invoking it.
        Assert.Equal(1366u, Assert.Single(legacy.Shape(item).Glyphs).GlyphId);
        Assert.Equal(1, legacy.Calls);
        var engine = new TextSpanShapingEngine(legacy);
        var error = Assert.Throws<TextLayoutException>(() => engine.Shape(new TextSpanShapingRequest(item,
            new[] { new TextSourceSpan(11, 1) })));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
        Assert.Equal(1, legacy.Calls);
        var ordinary = Arabic("ب", 11);
        Assert.Equal(1366u, Assert.Single(engine.Shape(new TextSpanShapingRequest(ordinary,
            new[] { new TextSourceSpan(11, 1) })).ShapedRun.Glyphs).GlyphId);
        Assert.Equal(2, legacy.Calls);
    }

    [Fact]
    public void AdapterEntryPointsRequireExplicitWholeRunOrContextualIntent()
    {
        var whole = Arabic("ببب", 10);
        var item = TextShapeRequest.CreateContextualItem(whole, new TextSourceSpan(11, 1));
        var shaper = new HarfBuzzTextShaper();
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() => shaper.Shape(item)).Error);
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() => shaper.ShapeContextualItem(whole)).Error);
    }

    private sealed class LegacySubstringShaper : ITextShaper
    {
        internal int Calls { get; private set; }
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
        {
            Calls++;
            var isolated = new TextShapeRequest(request.Text, request.SourceStart, request.Face,
                request.Direction, request.Script, request.Language, request.Features);
            var shaped = new HarfBuzzTextShaper().Shape(isolated, cancellationToken);
            return new TextShapedRun(request, shaped.Glyphs, shaped.EngineVersion);
        }
    }

    private sealed class ContextReplacingShaper : ITextShaper
    {
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
        {
            var context = new TextShapeRequest(request.Text + "x", request.SourceStart, request.Face,
                request.Direction, request.Script, request.Language);
            return new HarfBuzzTextShaper().ShapeContextualItem(TextShapeRequest.CreateContextualItem(context,
                new TextSourceSpan(request.SourceStart, request.Text.Length)), cancellationToken);
        }
    }

    private static TextShapeRequest Arabic(string text, int start) => new(text, start, Face.Value, TextDirection.RightToLeft, "Arab", "ar");
    private static object Identity(TextShapedGlyph g) => (g.GlyphId, g.ClusterStart, g.ClusterLength,
        g.XAdvance, g.YAdvance, g.XOffset, g.YOffset, g.UnsafeToBreak);
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
