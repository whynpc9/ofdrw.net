using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class VectorTests
{
    private static byte[] Font(string name = "regular") => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", name + ".ttf"));
    private static PdfVectorToOfdOptions Options(PdfVectorUnsupportedPagePolicy policy = PdfVectorUnsupportedPagePolicy.Fail) => new()
    { UnsupportedPagePolicy = policy, Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
    [Fact]
    public async Task ActualPdfPaintOrderOriginalSpacesAffineAndFontBytesSurviveReadback()
    {
        using var fixture = new PdfFixture(Font());
        var body = "q 1 .2 .15 1 12 8 cm 0 0 1 rg 10 70 180 35 re f 0 0 0 rg\n" + fixture.Text("A  B中文", 20, 85, 14) + "1 0 0 rg 40 65 14 50 re f Q\n";
        using var input = new MemoryStream(fixture.Create(new[] { body }, inheritedResources: true)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output);
        Assert.True(result.Pages.Single().IsNative); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output); var page = package.Pages.Single();
        Assert.Collection(page.Elements, element => Assert.IsType<OfdPathElement>(element), element => Assert.IsType<OfdTextElement>(element), element => Assert.IsType<OfdPathElement>(element));
        var text = page.Elements.OfType<OfdTextElement>().Single();
        Assert.Equal("A  B中文", text.Text); Assert.Equal("A  B中文", string.Concat(text.Runs.Select(run => run.Text)));
        Assert.Equal(Font(), package.Fonts.Single().Data); Assert.False(package.Fonts.Single().Bold);
        Assert.Equal(420 * 25.4 / 72, page.WidthMillimeters, 3); Assert.Equal(595 * 25.4 / 72, page.HeightMillimeters, 3);
        Assert.Contains("DeltaX", text.SourceXml); Assert.DoesNotContain(page.Elements, element => element is OfdImageElement);
    }
    [Theory]
    [InlineData("q 0 0 m W n Q", "CLIP_PROFILE")]
    [InlineData("/GS1 gs", "gs")]
    [InlineData("q 420 0 0 595 0 0 cm /Im1 Do Q", "IMAGE_ONLY")]
    [InlineData("23 42 XYZUnsupported", "XYZUnsupported")]
    [InlineData("BT 7 Tr ET", "Tr")]
    [InlineData("q 10 10 m 1 0 0 1 20 0 cm 20 20 l S Q", "PATH_STATE_CHANGE")]
    [InlineData("q Q 99", "TRAILING_OPERANDS")]
    [InlineData("99 q Q", "OPERATOR_ARITY")]
    public async Task UnsupportedAtFirstOperationFailsWithoutWriting(string content, string diagnostic)
    {
        using var fixture = new PdfFixture(Font()); using var input = new MemoryStream(fixture.Create(new[] { content }));
        using var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains(diagnostic, exception.Message); Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task WholePageFallbackDiscardsPriorNativePaintAndThenNextPageIsNative()
    {
        using var fixture = new PdfFixture(Font());
        var native = "0 0 0 rg\n" + fixture.Text("A  B中文", 20, 85) + "0 0 1 RG 10 10 m 40 40 l S";
        using var input = new MemoryStream(fixture.Create(new[] { native + " /GS1 gs", native + " /GS1 gs", native })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.RasterizePage)).ConvertWithResultAsync(input, output);
        Assert.False(result.Pages[0].IsNative); Assert.Equal(0, result.Pages[0].PathObjects); Assert.Equal(1, result.Pages[0].ImageObjects);
        Assert.False(result.Pages[1].IsNative); Assert.True(result.Pages[2].IsNative); output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.All(package.Pages[0].Elements.OfType<OfdTextElement>(), text => Assert.Equal(0, text.FillColor.Alpha));
        Assert.Equal("A  B中文", package.Pages[2].Elements.OfType<OfdTextElement>().Single().Text);
    }
    [Fact]
    public async Task GlyphMismatchAndNonRepresentablePageFailBeforeOutput()
    {
        using var fixture = new PdfFixture(Font()); var text = fixture.Text("AB", 20, 85);
        foreach (var pdf in new[] { fixture.Create(new[] { text }, wrongUnicode: true), fixture.Create(new[] { text }, pageEntries: "/Rotate 45"), fixture.Create(new[] { text }, pageEntries: "/Annots []"), fixture.Create(new[] { "" }) })
        {
            using var input = new MemoryStream(pdf); using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output)); Assert.Equal(0, output.Length);
        }
    }
    [Theory]
    [InlineData("/Rotate 90", "0 0 420 595", 595, 420)]
    [InlineData("/Rotate 180", "0 0 420 595", 420, 595)]
    [InlineData("/Rotate 270", "0 0 420 595", 595, 420)]
    [InlineData("/Rotate -90", "0 0 420 595", 595, 420)]
    [InlineData("/CropBox [10 20 300 400]", "0 0 420 595", 290, 380)]
    [InlineData("/CropBox [10 20 300 400] /Rotate 90", "0 0 420 595", 380, 290)]
    [InlineData("/CropBox [110 220 400 600] /Rotate 270", "100 200 520 795", 380, 290)]
    public async Task SupportedDisplayedPageFrameRetainsNativePaintAndText(string entries, string media, double width, double height)
    {
        using var fixture = new PdfFixture(Font());
        var content = "0 0 1 rg 140 270 40 25 re f " + fixture.Text("A  B中文", 160, 290);
        using var input = new MemoryStream(fixture.Create(new[] { content }, pageEntries: entries, mediaBox: media));
        using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages.Single().IsNative);
        output.Position = 0; var page = (await new OfdReader().ReadAsync(output)).Pages.Single();
        Assert.Equal(Math.Round(width * 25.4 / 72, 3), page.WidthMillimeters, 3);
        Assert.Equal(Math.Round(height * 25.4 / 72, 3), page.HeightMillimeters, 3);
        var path = Assert.Single(page.Elements.OfType<OfdPathElement>());
        Assert.InRange(path.XMillimeters, 0, page.WidthMillimeters);
        Assert.InRange(path.YMillimeters, 0, page.HeightMillimeters);
        Assert.Equal("A  B中文", page.Elements.OfType<OfdTextElement>().Single().Text);
    }
    [Theory]
    [InlineData("/CropBox [0 0 430 595]", "CROP_BOX")]
    [InlineData("/CropBox [20 20 20 100]", "CROP_BOX")]
    [InlineData("/Rotate 45", "PAGE_ROTATION")]
    public async Task InvalidDisplayedPageFrameRemainsWholePageFailure(string entries, string diagnostic)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("A", 20, 85) }, pageEntries: entries));
        using var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains(diagnostic, error.Message); Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray());
    }
    [Theory]
    [InlineData("W", "NonZero")]
    [InlineData("W*", "Even-Odd")]
    public async Task BoundedPathClipIntersectsOnlyItsGraphicsState(string operation, string rule)
    {
        using var fixture = new PdfFixture(Font());
        var content = "q 20 60 100 70 re " + operation + " n 0 0 1 rg 10 50 140 90 re f " +
            fixture.Text("A  B中文", 30, 90) + " Q 1 0 0 rg 150 50 20 20 re f";
        using var input = new MemoryStream(fixture.Create(new[] { content })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages.Single().IsNative); output.Position = 0;
        var page = (await new OfdReader().ReadAsync(output)).Pages.Single();
        Assert.Collection(page.Elements,
            path => { Assert.IsType<OfdPathElement>(path); Assert.Contains("Clips", path.ClippingXml); Assert.Contains(rule, path.ClippingXml); },
            text => { Assert.IsType<OfdTextElement>(text); Assert.Contains("Clips", text.ClippingXml); },
            path => { Assert.IsType<OfdPathElement>(path); Assert.Null(path.ClippingXml); });
    }
    [Theory]
    [InlineData("1 J 1 j [6 3] 2 d", "Round", "Round", "6 3", "2")]
    [InlineData("2 J 2 j [5 2 1] 1 d", "Square", "Bevel", "5 2 1", "1")]
    [InlineData("0 J 0 j 4 M [] 0 d", "Butt", "Miter", null, null)]
    public async Task BoundedStrokeStyleSurvivesWriteAndRead(string state, string cap, string join, string? dash, string? phase)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(fixture.Create(new[] { state + " 3 w 30 70 m 160 100 l S" }));
        using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages.Single().IsNative); output.Position = 0;
        var path = (await new OfdReader().ReadAsync(output)).Pages.Single().Elements.OfType<OfdPathElement>().Single();
        Assert.Contains("Cap=\"" + cap + "\"", path.SourceXml);
        Assert.Contains("Join=\"" + join + "\"", path.SourceXml);
        if (dash is null) Assert.DoesNotContain("DashPattern", path.SourceXml);
        else { Assert.Contains("DashPattern=\"" + dash + "\"", path.SourceXml); Assert.Contains("DashOffset=\"" + phase + "\"", path.SourceXml); }
    }
    [Theory]
    [InlineData("[0 3] 0 d", "STROKE_PROFILE")]
    [InlineData("1 J 5 M", "STROKE_PROFILE")]
    [InlineData("11 M", "STROKE_PROFILE")]
    public async Task UnrepresentableStrokeStyleStillFailsWholePage(string state, string diagnostic)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(fixture.Create(new[] { "0 0 1 rg 10 10 20 20 re f " + state + " 3 w 30 70 m 160 100 l S" }));
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains(diagnostic, error.Message); Assert.Empty(output.ToArray());
    }
    [Theory]
    [InlineData("regular", false, false)]
    [InlineData("semibold", false, false)]
    [InlineData("oblique", false, false)]
    [InlineData("bold", true, false)]
    [InlineData("italic", false, true)]
    public async Task ExactPhysicalFaceFlagsUseExistingResourceContract(string name, bool bold, bool italic)
    {
        using var fixture = new PdfFixture(Font(name)); using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("AB", 20, 85) })); using var output = new MemoryStream();
        await new PdfVectorToOfdConverter().ConvertAsync(input, output); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output); var font = package.Fonts.Single();
        Assert.Equal(Font(name), font.Data); Assert.Equal(bold, font.Bold); Assert.Equal(italic, font.Italic);
        var text = package.Pages[0].Elements.OfType<OfdTextElement>().Single(); Assert.Equal(bold ? 700 : 400, text.Weight); Assert.Equal(italic, text.Italic);
    }
    [Fact]
    public async Task PageSelectionPreservesDuplicatesAndTextOnlyPagesAreNative()
    {
        using var fixture = new PdfFixture(Font()); var pdf = fixture.Create(new[] { fixture.Text("AA", 20, 85), fixture.Text("BB", 20, 85) });
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output, new[] { 1, -1, 0, 1 });
        Assert.Equal(new[] { 1, 0, 1 }, result.Pages.Select(page => page.SourcePageIndex)); Assert.All(result.Pages, page => Assert.True(page.IsNative));
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(new[] { "BB", "AA", "BB" }, package.Pages.Select(page => page.Elements.OfType<OfdTextElement>().Single().Text));
    }
    [Theory]
    [InlineData("input")][InlineData("content")][InlineData("operations")][InlineData("events")][InlineData("path")][InlineData("text")][InlineData("font")]
    public async Task LimitsDoNotBecomeRasterFallback(string budget)
    {
        using var fixture = new PdfFixture(Font()); var pdf = fixture.Create(new[] { "0 0 1 rg 10 10 20 20 re f " + fixture.Text("AB", 20, 85) });
        var options = Options(PdfVectorUnsupportedPagePolicy.RasterizePage);
        switch (budget) { case "input": options.Compatibility.MaxInputBytes = 10; break; case "content": options.MaxContentBytesPerPage = 10; break;
            case "operations": options.MaxOperationsPerPage = 1; break; case "events": options.MaxEventsPerPage = 1; break; case "path": options.MaxPathCommandsPerPage = 1; break;
            case "text": options.MaxTextCharactersPerPage = 1; break; case "font": options.MaxFontBytes = 10; break; }
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        await Assert.ThrowsAnyAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output)); Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task DefaultConverterStillProducesOnlyImageAndTransparentText()
    {
        using var fixture = new PdfFixture(Font()); using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("AB", 20, 85) })); using var output = new MemoryStream();
        await new PdfToOfdConverter(new PdfToOfdOptions { PreferExternalPdfToPpm = true }).ConvertAsync(input, output); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output); Assert.Single(package.Pages[0].Elements.OfType<OfdImageElement>());
        Assert.Empty(package.Pages[0].Elements.OfType<OfdPathElement>()); Assert.All(package.Pages[0].Elements.OfType<OfdTextElement>(), text => Assert.Equal(0, text.FillColor.Alpha));
    }
    [Fact]
    public async Task PrecancellationWritesNothing()
    {
        using var input = new MemoryStream(new byte[100]); using var output = new MemoryStream();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output, cancellationToken: new CancellationToken(true)));
        Assert.Equal(0, output.Length);
    }
}
