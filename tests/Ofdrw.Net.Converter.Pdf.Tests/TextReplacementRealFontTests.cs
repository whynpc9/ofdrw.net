using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Geometry;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Trusted font geometry")]
public sealed class TextReplacementRealFontTests(Xunit.Abstractions.ITestOutputHelper output)
{
    [RealTrustedFontsTheory]
    [InlineData("latin")]
    public async Task ActualOriginalOflFontVariableReplacementRetainsIndependentAnchorsAndGlyphs(string face)
    {
        var bytes = TrustedFontFixtures.Read(face); var source = await Read(bytes, 2.6, "BBB");
        var before = new OfdKeywordLocator(Provider(bytes)).Find(source, "BBB"); var original = Assert.Single(before.Hits);
        // Explicitly select the middle genuine scalar hit from one operation.
        var all = new OfdKeywordLocator(Provider(bytes)).Find(source, "B");
        var hit = all.Hits[1]; var originalXml = Text(source).SourceXml;
        using var result = new OfdTextReplacer(Provider(bytes)).Replace(source, new[] { new OfdTextReplacement(hit, "--") });
        var package = result.CreatePackageCopy();
        Assert.Equal("B--B", Text(package).Text); Assert.Equal(originalXml, Text(source).SourceXml);
        Assert.NotSame(source.Fonts[0].Data, package.Fonts[0].Data); Assert.Equal(Digest(bytes), Digest(package.Fonts[0].Data));
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, saved); saved.Position = 0;
        var readback = await new OfdReader().ReadAsync(saved); var after = new OfdKeywordLocator(Provider(bytes)).Find(readback, "B");
        Assert.True(after.IsComplete); Assert.Equal(2, after.Hits.Count);
        Assert.Equal(original.Fragments[0].BaselineMillimeters.X, after.Hits[0].Fragments[0].BaselineMillimeters.X);
        Assert.Equal(original.Fragments[2].BaselineMillimeters.X, after.Hits[1].Fragments[0].BaselineMillimeters.X);
        SameEvidence(original.Fragments[0], after.Hits[0].Fragments[0]); SameEvidence(original.Fragments[2], after.Hits[1].Fragments[0]);
        Assert.Equal(22.2, after.Hits[1].Fragments[0].BaselineMillimeters.X, 12);
        Assert.Equal(37u, after.Hits[0].Fragments[0].GlyphId); Assert.Equal("30", Text(readback).ObjectId);
        Assert.Equal(Digest(bytes), Digest(readback.Fonts[0].Data));
        output.WriteLine(JsonSerializer.Serialize(new { face, stage = "actual-reader-hit-replace-transfer-writer-reader-locator",
            result.Statistics, suffix = after.Hits[1].Fragments[0].BaselineMillimeters.X, size = saved.Length }));
    }

    [RealTrustedFontsTheory]
    [InlineData("latin", 2.556)]
    [InlineData("cjk", 2.432)]
    public async Task OriginalOflFontLibraryTransferAndSeparatelyBudgetedWriterPassActualReadback(string face, double advance)
    {
        var bytes = TrustedFontFixtures.Read(face); var source = await Read(bytes, advance);
        var originals = new OfdKeywordLocator(Provider(bytes)).Find(source, "A").Hits;
        var hit = originals[1]; var before = Text(source).SourceXml;
        using var result = new OfdTextReplacer(Provider(bytes)).Replace(source, new[] { new OfdTextReplacement(hit, "-") });
        var package = result.CreatePackageCopy(); Assert.Equal("A-A", Text(package).Text);
        Assert.Equal(before, Text(source).SourceXml);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, saved); saved.Position = 0;
        var readback = await new OfdReader().ReadAsync(saved); var anchors = new OfdKeywordLocator(Provider(bytes)).Find(readback, "A");
        Assert.True(anchors.IsComplete); Assert.Equal(2, anchors.Hits.Count);
        Assert.Equal(17, anchors.Hits[0].Fragments[0].BaselineMillimeters.X);
        Assert.Equal(17 + 2 * advance, anchors.Hits[1].Fragments[0].BaselineMillimeters.X, 12);
        SameEvidence(originals[0].Fragments[0], anchors.Hits[0].Fragments[0]); SameEvidence(originals[2].Fragments[0], anchors.Hits[1].Fragments[0]);
        Assert.Equal(Digest(bytes), Digest(readback.Fonts[0].Data));
        output.WriteLine(JsonSerializer.Serialize(new { face, outcome = "library-transfer-separate-public-writer-readback-passed", result.Statistics,
            sourceFontAndEntryAlias = source.PreservedEntries.Values.Any(value => ReferenceEquals(value, source.Fonts[0].Data)),
            note = "Public Writer is a separate operation; this does not establish aggregate CLI save accounting." }));
    }

    private static void SameEvidence(OfdKeywordFragment expected, OfdKeywordFragment actual)
    {
        Assert.Equal((expected.BaselineMillimeters.X, expected.BaselineMillimeters.Y), (actual.BaselineMillimeters.X, actual.BaselineMillimeters.Y));
        Assert.Equal((expected.AdvanceXMillimeters, expected.AdvanceYMillimeters, expected.GlyphId), (actual.AdvanceXMillimeters, actual.AdvanceYMillimeters, actual.GlyphId));
        Assert.Equal((expected.FontResourceId, expected.FontSha256, expected.FontByteCount), (actual.FontResourceId, actual.FontSha256, actual.FontByteCount));
        Assert.Equal((expected.FontSizeMillimeters, expected.Weight, expected.Italic, expected.FontBold, expected.FontItalic),
            (actual.FontSizeMillimeters, actual.Weight, actual.Italic, actual.FontBold, actual.FontItalic));
        Assert.Equal(expected.InkQuadrilateralMillimeters!.Corners.Select(point => (point.X, point.Y)),
            actual.InkQuadrilateralMillimeters!.Corners.Select(point => (point.X, point.Y)));
    }

    [RealTrustedFontsTheory]
    [InlineData("latin", 2.556)]
    [InlineData("cjk", 2.432)]
    public async Task OriginalOflKeywordReplacementTransfersUnderDefaultBudgetAndSeparateWriterReadback(string face, double advance)
    {
        var bytes = TrustedFontFixtures.Read(face); var source = await Read(bytes, advance); var before = Text(source).SourceXml;
        var originals = new OfdKeywordLocator(Provider(bytes)).Find(source, "A").Hits;
        using var result = new OfdTextReplacer(Provider(bytes)).Replace(source, "A", "-");
        Assert.Single(result.Statistics.Analyses); Assert.Equal(3, result.ChangedCount);
        var package = result.CreatePackageCopy(); Assert.Equal("---", Text(package).Text);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, saved); saved.Position = 0;
        var readback = await new OfdReader().ReadAsync(saved);
        var found = new OfdKeywordLocator(Provider(bytes)).Find(readback, "-");
        Assert.True(found.IsComplete); Assert.Equal(3, found.Hits.Count);
        // Independent original-font measurement supplies expected replacement
        // metrics; output coordinates come from the original source anchors.
        using var metricOperation = new OfdKeywordOperation();
        var fontSnapshot = metricOperation.CaptureFont(source.Fonts[0], "10");
        using var metricSession = Provider(bytes).OpenSession(new());
        var metric = Assert.Single(metricSession.Measure(metricOperation.Request(fontSnapshot, "-", 4, 400, false)).Clusters);
        for (var i = 0; i < 3; i++)
        {
            var actual = found.Hits[i].Fragments[0]; var original = originals[i].Fragments[0];
            Assert.Equal((originals[i].Fragments[0].BaselineMillimeters.X, originals[i].Fragments[0].BaselineMillimeters.Y),
                (actual.BaselineMillimeters.X, actual.BaselineMillimeters.Y));
            Assert.Equal((metric.GlyphId, metric.AdvanceXMillimeters, metric.AdvanceYMillimeters),
                (actual.GlyphId, actual.AdvanceXMillimeters, actual.AdvanceYMillimeters));
            Assert.Equal((original.FontResourceId, original.FontSha256, original.FontByteCount, original.FontSizeMillimeters,
                original.Weight, original.Italic, original.FontBold, original.FontItalic),
                (actual.FontResourceId, actual.FontSha256, actual.FontByteCount, actual.FontSizeMillimeters,
                actual.Weight, actual.Italic, actual.FontBold, actual.FontItalic));
            var box = metric.FontBoundsMillimeters!; var x = original.BaselineMillimeters.X; var y = original.BaselineMillimeters.Y;
            var expectedCorners = new[] { (x + box.Left, y + box.Top), (x + box.Right, y + box.Top), (x + box.Right, y + box.Bottom), (x + box.Left, y + box.Bottom) };
            for (var corner = 0; corner < 4; corner++)
            {
                // Independent coordinate composition can differ in the last
                // floating bit; source baseline preservation above is exact.
                Assert.Equal(expectedCorners[corner].Item1, actual.InkQuadrilateralMillimeters!.Corners[corner].X, 12);
                Assert.Equal(expectedCorners[corner].Item2, actual.InkQuadrilateralMillimeters.Corners[corner].Y, 12);
            }
        }
        Assert.Equal("30", Text(readback).ObjectId); Assert.Equal("10", Text(readback).FontResourceId);
        Assert.Single(Text(readback).Runs); Assert.Equal(before, Text(source).SourceXml);
        Assert.Equal(Digest(bytes), Digest(readback.Fonts[0].Data)); Assert.Equal(Digest(bytes), Digest(source.Fonts[0].Data));
        output.WriteLine(JsonSerializer.Serialize(new { face, outcome = "fused-keyword-default-budget-transfer-separate-writer-readback-passed",
            result.Statistics, size = saved.Length, sourceFontBytes = bytes.Length,
            note = "Separate public Writer operation; aggregate CLI accounting remains a separate gate." }));
    }

    [RealTrustedFontsTheory]
    [InlineData("cjk")]
    public async Task OriginalCjkKeywordLowerCapacityIsAnExplicitOutputlessRefusal(string face)
    {
        var bytes = TrustedFontFixtures.Read(face); var source = await Read(bytes, 2.432); var before = Text(source).SourceXml;
        var error = Assert.Throws<InvalidDataException>(() => new OfdTextReplacer(Provider(bytes)).Replace(source, "A", "-",
            new OfdTextReplacementOptions(new(maxMaterializedBytes: 256L * 1024 * 1024))));
        Assert.Contains("budget", error.Message, StringComparison.OrdinalIgnoreCase);
        var counters = Assert.IsType<OfdTextReplacementPhase>(error.Data["Ofdrw.TextReplacement.Counters"]);
        Assert.Equal(before, Text(source).SourceXml); Assert.Equal(Digest(bytes), Digest(source.Fonts[0].Data));
        output.WriteLine(JsonSerializer.Serialize(new { face, outcome = "keyword-driven-capacity-refusal-no-result", error.Message,
            phase = error.Data["Ofdrw.TextReplacement.Phase"], counters,
            note = "Deliberate lower-capacity negative, separate from default-budget positive." }));
    }

    private static IOfdGlyphGeometryProvider Provider(byte[] bytes)
        => new OfdTrustedEmbeddedFontGeometryProvider(new[] { new OfdTrustedEmbeddedFontIdentity(Digest(bytes), bytes.Length) });
    private static string Digest(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static OfdTextElement Text(OfdDocumentPackage package) => (OfdTextElement)package.Pages[0].Elements[0];
    private static async Task<OfdDocumentPackage> Read(byte[] bytes, double advance, string value = "AAA")
    {
        var package = new OfdDocumentPackage();
        // The trust profile below authenticates the original full OFL bytes, not a subset.
        package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new() { Id = "10", FontName = "OriginalOFL", FileName = "font.ttf", Data = bytes });
        var page = new OfdPage { Id = "1", XMillimeters = 5, YMillimeters = 7, WidthMillimeters = 200, HeightMillimeters = 200 };
        var text = new OfdTextElement { ObjectId = "30", LayerId = "20", FontResourceId = "10", FontName = "OriginalOFL", Text = value,
            XMillimeters = 20, YMillimeters = 30, WidthMillimeters = 100, HeightMillimeters = 50, FontSizeMillimeters = 4 };
        var delta = advance.ToString("R", CultureInfo.InvariantCulture);
        text.Runs.Add(new() { Text = value, XMillimeters = 2, YMillimeters = 8, DeltaX = delta + " " + delta });
        page.Elements.Add(text); package.Pages.Add(page);
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        return await new OfdReader().ReadAsync(input);
    }
}
