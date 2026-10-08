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
    [InlineData("q 0 0 50 50 re W n Q", "W")]
    [InlineData("/GS1 gs", "gs")]
    [InlineData("q 420 0 0 595 0 0 cm /Im1 Do Q", "Do")]
    [InlineData("23 42 XYZUnsupported", "XYZUnsupported")]
    [InlineData("BT 7 Tr ET", "Tr")]
    [InlineData("q 10 10 m 1 0 0 1 20 0 cm 20 20 l S Q", "PATH_STATE_CHANGE")]
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
        foreach (var pdf in new[] { fixture.Create(new[] { text }, wrongUnicode: true), fixture.Create(new[] { text }, pageEntries: "/Rotate 90"), fixture.Create(new[] { text }, pageEntries: "/Annots []"), fixture.Create(new[] { "" }) })
        {
            using var input = new MemoryStream(pdf); using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output)); Assert.Equal(0, output.Length);
        }
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
