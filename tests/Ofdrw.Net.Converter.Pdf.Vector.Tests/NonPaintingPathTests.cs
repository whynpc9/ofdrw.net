using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class NonPaintingPathTests
{
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static PdfVectorToOfdOptions Options(PdfVectorUnsupportedPagePolicy policy) => new()
    { UnsupportedPagePolicy = policy, Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };

    [Theory]
    [InlineData("20 20 m S")][InlineData("20 20 m h S")][InlineData("20 20 m s")]
    [InlineData("20 20 m 40 40 m f")][InlineData("20 20 m 40 40 m F")][InlineData("20 20 m 40 40 m f*")]
    [InlineData("20 20 m B")][InlineData("20 20 m B*")]
    public async Task SafeNoOpsPreserveMixedNativeContentAndNoContentUsesPagePolicy(string noOp)
    {
        using var fixture = Fixture();
        var mixed = noOp + " q 1 0 0 1 2 0 cm " + fixture.Text("A  B中文", 20, 85) +
            " Q 0 0 1 RG 10 10 m 30 30 l S " + noOp;
        foreach (var policy in new[] { PdfVectorUnsupportedPagePolicy.Fail, PdfVectorUnsupportedPagePolicy.RasterizePage })
        {
            var options = Options(policy); options.MaxEventsPerPage = 2;
            using var input = new MemoryStream(fixture.Create(new[] { mixed })); using var output = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
            Assert.True(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].PathObjects); Assert.Equal(1, result.Pages[0].TextObjects);
            Assert.Equal(0, result.Pages[0].ImageObjects); output.Position = 0; var package = await new OfdReader().ReadAsync(output);
            Assert.Equal("A  B中文", package.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
        }
        using var failInput = new MemoryStream(fixture.Create(new[] { mixed, noOp })); using var failOutput = new MemoryStream();
        failOutput.Write(new byte[] { 1, 2, 3 });
        var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(failInput, failOutput));
        Assert.Contains("NO_NATIVE_CONTENT", exception.Message); Assert.Equal(new byte[] { 1, 2, 3 }, failOutput.ToArray());
        using var rasterInput = new MemoryStream(fixture.Create(new[] { noOp })); using var rasterOutput = new MemoryStream();
        var raster = await new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.RasterizePage)).ConvertWithResultAsync(rasterInput, rasterOutput, new[] { 0, 0 });
        Assert.Equal(2, raster.Pages.Count);
        Assert.All(raster.Pages, page => { Assert.False(page.IsNative); Assert.Equal(1, page.ImageObjects); Assert.Equal(0, page.PathObjects); Assert.Contains("NO_NATIVE_CONTENT", page.Diagnostic); });
        rasterOutput.Position = 0; var readback = await new OfdReader().ReadAsync(rasterOutput);
        Assert.All(readback.Pages, page => Assert.Single(page.Elements.OfType<OfdImageElement>()));
    }

    [Theory]
    [InlineData("20 20 m h f")][InlineData("20 20 m h F")][InlineData("20 20 m h f*")]
    [InlineData("20 20 m h B")][InlineData("20 20 m h B*")][InlineData("20 20 m b")][InlineData("20 20 m b*")]
    [InlineData("20 20 m h 40 40 m f")][InlineData("20 20 m h 40 40 m 60 60 l f")]
    public async Task ClosedSingletonFillDoesNotDisappearAmongOtherNativeContent(string pointFill)
    {
        using var fixture = Fixture(); var text = fixture.Text("A  B中文", 20, 85);
        var pdf = fixture.Create(new[] { text, text + " 0 0 20 20 re f " + pointFill, text });
        using var failInput = new MemoryStream(pdf); using var failOutput = new MemoryStream();
        var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(failInput, failOutput));
        Assert.Contains("DEGENERATE_POINT_FILL", exception.Message); Assert.Equal(0, failOutput.Length);
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.RasterizePage)).ConvertWithResultAsync(input, output);
        Assert.Equal(new[] { true, false, true }, result.Pages.Select(page => page.IsNative));
        Assert.Equal(0, result.Pages[1].PathObjects); Assert.Equal(1, result.Pages[1].ImageObjects);
        Assert.Contains("DEGENERATE_POINT_FILL", result.Pages[1].Diagnostic); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.All(package.Pages[1].Elements.OfType<OfdTextElement>(), element => Assert.Equal(0, element.FillColor.Alpha));
        Assert.Equal("A  B中文", package.Pages[2].Elements.OfType<OfdTextElement>().Single().Text);
    }

    [Theory]
    [InlineData("20 20 m 20 20 l S")][InlineData("20 20 m 20 20 20 20 20 20 c S")]
    [InlineData("20 20 m 20 20 20 20 v S")][InlineData("20 20 m 20 20 20 20 y S")]
    [InlineData("20 20 0 0 re S")][InlineData("20 20 m 50 40 -10 40 20 20 c S")]
    [InlineData("20 20 m 40 40 l 60 60 m S")][InlineData("20 20 m h 40 40 l S")]
    public async Task ExplicitSegmentsRemainNativeRegardlessOfBoundsOrEndpointEquality(string content)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { content })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].PathObjects); output.Position = 0;
        var page = (await new OfdReader().ReadAsync(output)).Pages[0];
        Assert.NotEmpty(page.Elements.OfType<OfdPathElement>().Single().AbbreviatedData);
    }

    [Fact]
    public async Task NoOpAndDiscardedPathResetStateWithoutErasingOtherPaints()
    {
        using var fixture = Fixture(); var content = "10 10 m 30 30 l S 10 10 m h S q 1 0 0 1 2 0 cm 40 40 m 50 50 l n Q 20 20 m S 50 50 m 70 70 l S S";
        using var input = new MemoryStream(fixture.Create(new[] { content })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages[0].IsNative); Assert.Equal(2, result.Pages[0].PathObjects);
    }

    [Theory]
    [InlineData("20 20 m S 30 30 m S", false)]
    [InlineData("20 20 m h S", false)][InlineData("20 20 m s", false)]
    [InlineData("20 20 m S", true)]
    public async Task IgnoredPaintStillChargesCommandsAndOperations(string content, bool operations)
    {
        using var fixture = Fixture(); var options = Options(PdfVectorUnsupportedPagePolicy.RasterizePage);
        if (operations) options.MaxOperationsPerPage = 1; else options.MaxPathCommandsPerPage = 1;
        using var input = new MemoryStream(fixture.Create(new[] { content + " " + fixture.Text("AB", 20, 85) })); using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task NoOpsConsumeNoEventsButRealFillAndStrokeStillConsumeTwo()
    {
        using var fixture = Fixture(); var options = Options(PdfVectorUnsupportedPagePolicy.RasterizePage); options.MaxEventsPerPage = 1;
        using var input = new MemoryStream(fixture.Create(new[] { "20 20 m S 20 20 m h S 10 10 m 30 30 l S" })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output); Assert.Equal(1, result.Pages[0].PathObjects);
        foreach (var real in new[] { "10 10 20 20 re B", "10 10 m 30 30 l S 40 40 m 60 60 l S" })
        {
            using var failInput = new MemoryStream(fixture.Create(new[] { "20 20 m S " + real })); using var failOutput = new MemoryStream();
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(failInput, failOutput)); Assert.Equal(0, failOutput.Length);
        }
    }

    [Fact]
    public async Task RoundCapSingletonCannotBeSilentlyElided()
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { "1 J 20 20 m h S " + fixture.Text("AB", 20, 85) })); using var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("STROKE_PROFILE", exception.Message); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("0 0 0 0 0 0")][InlineData("1 2 2 4 0 0")][InlineData("1 1 1 1.000000001 0 0")]
    public async Task SingularAndFloatCollapsedTransformsUsePagePolicyBeforeAdapterValidation(string matrix)
    {
        using var fixture = Fixture(); var text = fixture.Text("AB", 20, 85);
        foreach (var invalid in new[] { "q " + matrix + " cm 20 20 m 40 40 l S Q", "BT /F1 12 Tf " + matrix + " Tm <" + fixture.Hex("A") + "> Tj ET" })
        {
            var pdf = fixture.Create(new[] { text, invalid, text });
            using var failInput = new MemoryStream(pdf); using var failOutput = new MemoryStream();
            var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(failInput, failOutput));
            Assert.Contains("SINGULAR_SERIALIZED_MATRIX", exception.Message); Assert.Equal(0, failOutput.Length);
            using var input = new MemoryStream(pdf); using var output = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(Options(PdfVectorUnsupportedPagePolicy.RasterizePage)).ConvertWithResultAsync(input, output);
            Assert.Equal(new[] { true, false, true }, result.Pages.Select(page => page.IsNative)); Assert.Equal(1, result.Pages[1].ImageObjects);
            Assert.Equal(0, result.Pages[1].PathObjects); Assert.Contains("SINGULAR_SERIALIZED_MATRIX", result.Pages[1].Diagnostic);
            output.Position = 0; var package = await new OfdReader().ReadAsync(output);
            Assert.All(package.Pages[1].Elements.OfType<OfdTextElement>(), element => Assert.Equal(0, element.FillColor.Alpha));
        }
    }
}
