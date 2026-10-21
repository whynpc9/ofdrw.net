using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextPrelayoutTests
{
    [Fact]
    public void FixedFaceSnapshotAndNativeShapeUseDesignUnitsAndUtf16Clusters()
    {
        var data = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var face = new TextFontFace("Noto Sans", data);
        data[0] = 0xff;
        var copied = face.GetFaceBytes();
        copied[0] = 0xff;
        Assert.NotEqual(0xff, face.GetFaceBytes()[0]);
        Assert.Equal("b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5", face.SourceSha256);
        var shaper = new HarfBuzzTextShaper();
        var result = shaper.Shape(new TextShapeRequest("AV fi", 12, face));
        Assert.Equal(face, result.Request.Face);
        Assert.NotEmpty(result.Glyphs);
        Assert.Contains("libHarfBuzzSharp/14.2.1", result.EngineVersion);
        Assert.True(result.XAdvance > 0);
        Assert.All(result.Glyphs, glyph =>
        {
            Assert.InRange(glyph.ClusterStart, 12, 16);
            Assert.InRange(glyph.ClusterLength, 1, 5);
            Assert.NotEqual(0u, glyph.GlyphId);
        });
        Assert.Contains(result.Glyphs, glyph => glyph.Ink is not null);
    }

    [Fact]
    public void FullRunFeaturesChangeLigatureAndDoNotSplitIntoGraphemeShaping()
    {
        var face = Noto();
        var shaper = new HarfBuzzTextShaper();
        var normal = shaper.Shape(new TextShapeRequest("ffi", 0, face));
        var disabled = shaper.Shape(new TextShapeRequest("ffi", 0, face,
            features: new[] { new TextFeature("liga", 0, 0, 3) }));
        Assert.True(normal.Glyphs.Count < disabled.Glyphs.Count);
        Assert.Equal(3, normal.Glyphs[0].ClusterLength);
        Assert.Equal("ffi", normal.Request.Text);
    }

    [Fact]
    public void SameDisplayNameWithDifferentFaceBytesChangesMeasuredWidth()
    {
        var noto = new TextFontFace("Alias", File.ReadAllBytes(Font("NotoSans-Regular.ttf")));
        var wenkai = new TextFontFace("Alias", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var engine = HarfBuzzTextPrelayout.Create();
        var first = engine.Layout(new TextPrelayoutRequest("AV To", new[] { noto }, 4, 100));
        var second = engine.Layout(new TextPrelayoutRequest("AV To", new[] { wenkai }, 4, 100));
        Assert.NotEqual(first.Face.FaceSha256, second.Face.FaceSha256);
        Assert.NotEqual(first.WidthMillimeters, second.WidthMillimeters);
        Assert.NotEqual(first.Lines[0].ShapedRun.XAdvance, second.Lines[0].ShapedRun.XAdvance);
    }

    [Fact]
    public void RtlOutputOrderDoesNotCorruptAbsoluteSourceClusterIntervals()
    {
        var run = new HarfBuzzTextShaper().Shape(new TextShapeRequest("ABC", 8, Noto(),
            TextDirection.RightToLeft));
        Assert.Equal(new[] { 10, 9, 8 }, run.Glyphs.Select(glyph => glyph.ClusterStart));
        Assert.All(run.Glyphs, glyph => Assert.Equal(1, glyph.ClusterLength));
    }

    [Fact]
    public void FinalLinesExposeHeightBaselineSourceAndAdvanceSeparatelyFromInk()
    {
        var face = Noto();
        var engine = HarfBuzzTextPrelayout.Create();
        var wide = engine.Layout(new TextPrelayoutRequest("AV fi\nCafe\u0301", new[] { face }, 4, 100));
        Assert.Equal(2, wide.Lines.Count);
        Assert.Equal((0, 5, 1), (wide.Lines[0].SourceStart, wide.Lines[0].SourceLength, wide.Lines[0].TerminatorLength));
        Assert.Equal((6, 5, 0), (wide.Lines[1].SourceStart, wide.Lines[1].SourceLength, wide.Lines[1].TerminatorLength));
        Assert.True(wide.Lines[0].BaselineMillimeters > wide.Lines[0].TopMillimeters);
        Assert.Equal(wide.Lines[0].HeightMillimeters, wide.Lines[1].TopMillimeters, 6);
        Assert.NotNull(wide.Lines[0].InkBoundsMillimeters);
        Assert.True(wide.WidthMillimeters > 0);
        Assert.False(wide.Overflows);
        Assert.Same(wide.Lines[0].ShapedRun.Request.Face, wide.Face);
        var narrow = engine.Layout(new TextPrelayoutRequest("AV fi", new[] { face }, 4, 5));
        Assert.True(narrow.Lines.Count >= 2);
        Assert.Equal("AV fi", string.Concat(narrow.Lines.Select(line =>
            narrow.Request.Text.Substring(line.SourceStart, line.SourceLength))));
    }

    [Fact]
    public void ExplicitWholeRunFallbackAndFailureRemainObservable()
    {
        var noto = Noto();
        var wenkai = new TextFontFace("WenKai", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var request = new TextPrelayoutRequest("中文", new[] { noto, wenkai }, 4, 30, script: "Hani", language: "zh");
        var result = HarfBuzzTextPrelayout.Create().Layout(request);
        Assert.Same(wenkai, result.Face);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "TEXT_FACE_FALLBACK");
        Assert.Equal(2, result.Lines[0].ShapedRun.Glyphs.Count);
        var missing = Assert.Throws<TextLayoutException>(() => HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("中文", new[] { noto }, 4, 30, script: "Hani", language: "zh")));
        Assert.Equal(TextLayoutError.MissingGlyph, missing.Error);
        Assert.Equal(0, missing.SourceStart);
    }

    [Fact]
    public void SameDisplayNameStyleAndTtcIndexNeverAliasTheSelectedFace()
    {
        var regular = new TextFontFace("Alias", File.ReadAllBytes(OtherFont("regular.ttf")));
        var bold = new TextFontFace("Alias", File.ReadAllBytes(OtherFont("bold.ttf")));
        var italic = new TextFontFace("Alias", File.ReadAllBytes(OtherFont("italic.ttf")));
        Assert.NotEqual(regular.FaceSha256, bold.FaceSha256);
        Assert.NotEqual(regular.FaceSha256, italic.FaceSha256);
        Assert.False(regular.Bold);
        Assert.True(bold.Bold);
        Assert.True(italic.Italic);
        var collection = File.ReadAllBytes(CollectionFont());
        var first = new TextFontFace("Alias", collection, 0);
        Assert.Equal(0, first.CollectionFaceIndex);
        Assert.NotEqual(first.SourceSha256, first.FaceSha256);
        Assert.Equal(TextLayoutError.InvalidFont, Assert.Throws<TextLayoutException>(() =>
            new TextFontFace("Alias", collection, 1)).Error);
    }

    [Fact]
    public void InvalidInputAndLegacyFlowBehaviorStaySeparate()
    {
        var face = Noto();
        Assert.Equal(TextLayoutError.InvalidText, Assert.Throws<TextLayoutException>(() =>
            new TextShapeRequest("A\ud800", 0, face)).Error);
        Assert.Equal(TextLayoutError.InvalidText, Assert.Throws<TextLayoutException>(() =>
            new TextPrelayoutRequest("A\rB", new[] { face }, 4, 20)).Error);
        var old = new FlowDocument();
        old.Blocks.Add(new Paragraph("AV fi"));
        Assert.Single(old.Render().Pages);
        Assert.Single(old.Render().Pages[0].Elements);
    }

    [Fact]
    public void CancellationAndShapeBudgetFailWithoutAnyPartialLayout()
    {
        var face = Noto();
        var engine = HarfBuzzTextPrelayout.Create();
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => engine.Layout(
            new TextPrelayoutRequest("AV To", new[] { face }, 4, 3), canceled.Token));
        var error = Assert.Throws<TextLayoutException>(() => engine.Layout(
            new TextPrelayoutRequest("AV To", new[] { face }, 4, 3, maximumShapeCalls: 1)));
        Assert.Equal(TextLayoutError.BudgetExceeded, error.Error);
    }

    [Fact]
    public async Task RepresentableAsciiWriterReusesMeasuredRunAndExactEmbeddedFace()
    {
        var face = Noto();
        var layout = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV A\nTo", new[] { face }, 4, 50));
        var package = TextPrelayoutOfdWriter.CreateSinglePage(layout, 100, 100, 10, 15);
        var text = package.Pages[0].Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "AV A", "To" }, text.Select(element => element.Text));
        for (var i = 0; i < text.Length; i++)
        {
            Assert.Equal(15 + layout.Lines[i].TopMillimeters, text[i].YMillimeters, 6);
            Assert.Equal(layout.Lines[i].BaselineMillimeters - layout.Lines[i].TopMillimeters,
                text[i].Runs[0].YMillimeters, 6);
            var deltas = text[i].Runs[0].DeltaX!.Split(' ').Select(value =>
                double.Parse(value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            Assert.Equal(layout.Lines[i].ShapedRun.Glyphs.Count - 1, deltas.Length);
            for (var j = 0; j < deltas.Length; j++) Assert.Equal(
                layout.Lines[i].ShapedRun.Glyphs[j].XAdvance * 4d / face.UnitsPerEm, deltas[j], 10);
        }
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var read = await new OfdReader().ReadAsync(stream);
        Assert.Equal(new[] { "AV A", "To" }, read.Pages[0].Elements.OfType<OfdTextElement>().Select(element => element.Text));
        var embedded = Assert.Single(read.Fonts).Data;
        Assert.Equal(face.FaceSha256, Convert.ToHexString(SHA256.HashData(embedded)).ToLowerInvariant());
        var ligature = HarfBuzzTextPrelayout.Create().Layout(new TextPrelayoutRequest("ffi", new[] { face }, 4, 50));
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePage(ligature, 100, 100, 10, 15)).Error);
    }

    private static TextFontFace Noto() => new("Noto Sans", File.ReadAllBytes(Font("NotoSans-Regular.ttf")));

    private static string Font(string file) => Fixture("Ofdrw.Net.FontSubset.E2E", file);
    private static string OtherFont(string file) => Fixture("Ofdrw.Net.SkiaSharp.E2E", file);
    private static string CollectionFont() => Fixture("Ofdrw.Net.Converter.Pdf.E2E", "style-collection.ttc");

    private static string Fixture(string project, string file)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "e2e", project, "testdata", "fonts", file);
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        throw new FileNotFoundException("Materialize the licensed font fixture through Git LFS.", file);
    }
}
