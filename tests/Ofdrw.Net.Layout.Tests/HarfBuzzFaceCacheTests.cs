using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;

namespace Ofdrw.Net.Layout.Tests;

public sealed class HarfBuzzFaceCacheTests
{
    [Fact]
    public void RepeatedShapingDoesNotCopyTheFacePerCall()
    {
        const string text = "敏捷的棕色狐狸";
        var face = ParagraphLayoutTests.Face("LXGW WenKai", "LXGWWenKai-Regular.ttf");
        var faceBytes = face.GetFaceBytes().Length;
        var shaper = new HarfBuzzTextShaper();
        var shared = new TextShapeRequest(text, 0, face, TextDirection.LeftToRight, "Hani", "zh");
        AssertSmallAllocations(shared, 200);

        // Together with "zh", these fill the 32-entry segment cache.
        for (var i = 0; i < 31; i++)
            shaper.Shape(new TextShapeRequest(text, 0, face, TextDirection.LeftToRight, "Hani", $"zh-x-cache{i}"));
        var overflow = new TextShapeRequest(text, 0, face, TextDirection.LeftToRight, "Hani", "zh-x-overflow");
        AssertSmallAllocations(overflow, 16);
        var featured = new TextShapeRequest(text, 0, face, TextDirection.LeftToRight, "Hani", "zh",
            new[] { new TextFeature("kern", 0, 0, text.Length) });
        AssertSmallAllocations(featured, 16);

        void AssertSmallAllocations(TextShapeRequest request, int repetitions)
        {
            shaper.Shape(request);
            var start = GC.GetAllocatedBytesForCurrentThread();
            for (var i = 0; i < repetitions; i++)
                shaper.Shape(request);
            var allocated = GC.GetAllocatedBytesForCurrentThread() - start;
            Assert.True(allocated < faceBytes, $"Shaping {repetitions} runs allocated {allocated} bytes.");
        }
    }

    [Fact]
    public void CachedFaceGivesIdenticalRuns()
    {
        var face = ParagraphLayoutTests.Face("Noto Sans", "NotoSans-Regular.ttf");
        var sameBytes = new TextFontFace(face.Name, face.GetFaceBytes());
        var shaper = new HarfBuzzTextShaper();
        const string text = "AV office e\u0301";
        var request = new TextShapeRequest(text, 10, face);
        var first = shaper.Shape(request);

        AssertSameGlyphs(first, new HarfBuzzTextShaper().Shape(request));
        AssertSameGlyphs(first, shaper.Shape(new TextShapeRequest(text, 10, sameBytes)));
        var kernOff = new[] { new TextFeature("kern", 0, 0, 2) };
        var featured = shaper.Shape(new TextShapeRequest(text, 10, face, features: kernOff));
        AssertSameGlyphs(featured, shaper.Shape(new TextShapeRequest(text, 10, sameBytes, features: kernOff)));
        Assert.NotEqual(first.Glyphs[0].XAdvance, featured.Glyphs[0].XAdvance);
        AssertSameGlyphs(first, shaper.Shape(request));
    }

    [Fact]
    public void ConcurrentShapingOfOneFaceIsConsistent()
    {
        var face = ParagraphLayoutTests.Face("Noto Sans", "NotoSans-Regular.ttf");
        var referenceFace = new TextFontFace(face.Name, face.GetFaceBytes());
        const string text = "JAV To office e\u0301";
        var expected = new HarfBuzzTextShaper().Shape(new TextShapeRequest(text, 0, referenceFace));

        // The workers race on a cold cache entry, using independent adapters.
        var request = new TextShapeRequest(text, 0, face);
        var actual = new TextShapedRun[64];
        Parallel.For(0, actual.Length, i => actual[i] = new HarfBuzzTextShaper().Shape(request));
        Assert.All(actual, run => AssertSameGlyphs(expected, run));
    }

    private static void AssertSameGlyphs(TextShapedRun expected, TextShapedRun actual)
    {
        var snapshot = (TextShapedGlyph glyph) => (glyph.GlyphId, glyph.ClusterStart, glyph.ClusterLength,
            glyph.XAdvance, glyph.YAdvance, glyph.XOffset, glyph.YOffset, glyph.UnsafeToBreak,
            glyph.Ink?.X, glyph.Ink?.Y, glyph.Ink?.Width, glyph.Ink?.Height);
        Assert.Equal(expected.Glyphs.Select(snapshot), actual.Glyphs.Select(snapshot));
    }
}
