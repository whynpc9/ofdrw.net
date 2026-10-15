using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class PathPrecisionTests
{
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static PdfVectorToOfdOptions Raster() => new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
        Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
    [Theory]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777217 100 l S Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777219 100 m 16777223 100 l S Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777217 120 16777219 80 16777220 100 c S Q")]
    [InlineData("q 1 0 0 1 -16777216 -16777216 cm 16777216 16777216 m 16777217 16777217 16777217 16777217 16777216 16777216 c S Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777217 120 16777220 100 v S Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777217 120 16777220 100 y S Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777217 100 m 16777220 120 l h S Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777216 100 1 40 re f Q")]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777217 100 -1 40 re B* Q")]
    [InlineData("q 1 0 0 1 -9007199254740992 0 cm 9007199254740992 100 1 40 re f Q")]
    [InlineData("q 1 0 0 1 -16777217 0 cm 16777218 100 m 16777220 100 l S Q")]
    [InlineData("q 1.00000003 0 0 1 -16777216 0 cm 16777216 100 m 16777220 100 l S Q")]
    [InlineData("q 1000000 0 0 1 -1000000 0 cm 1.00000001 100 m 1.00000002 100 l S Q")]
    [InlineData("0.0000000000000000000000000000000000000000000001 100 m 0 100 l S")]
    [InlineData("q 1 1 1 1.00001 0 0 cm 0 0 m 30 0 l S Q")]
    [InlineData("q 1.00004 0 0 1 0 0 cm 4000 w 0 0 m 0 30 l S Q")]
    [InlineData("0.0000000000000000000000000000000000000000000001 w 10 100 m 30 100 l S")]
    [InlineData("60.1 350.2 m 61.3 351.4 l S")]
    public async Task LossyPaintUsesWholePagePolicyWithoutNativeLeakage(string paint)
    {
        using var fixture = Fixture(); var text = fixture.Text("A  B中文", 20, 85);
        var pdf = fixture.Create(new[] { text, text + " 10 10 20 20 re f " + paint, text });
        using var failInput = new MemoryStream(pdf); using var failOutput = new MemoryStream(); failOutput.Write(new byte[] { 4, 5, 6 });
        var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(failInput, failOutput));
        Assert.Contains("PATH_FLOAT_PRECISION", exception.Message); Assert.Equal(new byte[] { 4, 5, 6 }, failOutput.ToArray());
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Raster()).ConvertWithResultAsync(input, output);
        Assert.Equal(new[] { true, false, true }, result.Pages.Select(page => page.IsNative));
        Assert.Equal(1, result.Pages[1].ImageObjects); Assert.Equal(0, result.Pages[1].PathObjects); Assert.Contains("PATH_FLOAT_PRECISION", result.Pages[1].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.All(package.Pages[1].Elements.OfType<OfdTextElement>(), text => Assert.Equal(0, text.FillColor.Alpha));
        Assert.Equal("A  B中文", package.Pages[2].Elements.OfType<OfdTextElement>().Single().Text);
        var nativeFontIds = package.Pages.Where((_, index) => index != 1).SelectMany(page => page.Elements.OfType<OfdTextElement>()).Select(text => text.FontResourceId).ToHashSet();
        Assert.All(package.Fonts.Where(font => font.Data.Length > 0), font => Assert.Contains(font.Id, nativeFontIds));
    }
    [Theory]
    [InlineData("q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777218 100 l S Q")]
    [InlineData("10.1 20.2 m 30.3 40.4 l S")]
    [InlineData("20 20 m 20 20 l S")]
    [InlineData("20 20 m 20 20 20 20 20 20 c S")]
    [InlineData("q 1 .2 .15 1 12 8 cm 10 70 180 35 re f Q")]
    [InlineData("q .978 -.208 .208 .978 306 142 cm 1.8 w -65 -18 130 40 re S Q")]
    public async Task RepresentableBoundariesDecimalsAndDeliberateCoincidenceStayNative(string paint)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { paint })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].PathObjects); Assert.Equal(0, result.Pages[0].ImageObjects);
    }
    [Fact]
    public async Task DiscardedAndTrailingMovesDoNotTurnSupportedTextIntoFallback()
    {
        using var fixture = Fixture(); var text = fixture.Text("A  B中文", 20, 85);
        var body = "q 1 0 0 1 -16777216 0 cm 16777216 100 m 16777217 100 l n 16777217 100 m S Q " +
            text + " 10 10 m 20 20 l 16777217 100 m S";
        using var input = new MemoryStream(fixture.Create(new[] { body })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].PathObjects); Assert.Equal(1, result.Pages[0].TextObjects);
    }
    [Fact]
    public async Task PrecisionFailureCannotBypassCommandLimit()
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { "16777216 100 m 16777217 100 l S" })); using var output = new MemoryStream();
        var options = Raster(); options.MaxPathCommandsPerPage = 1;
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output)); Assert.Equal(0, output.Length);
    }
}
