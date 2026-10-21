using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class MixedImageTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));

    [Fact]
    public async Task MixedNativePagePreservesPathImageTextOrderAndImagePixels()
    {
        byte[] pixels = { 255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0 };
        using var fixture = new PdfFixture(Font);
        var body = "1 0 0 rg 20 20 35 25 re f\nq 60 0 0 50 70 90 cm /Im1 Do Q\n" + fixture.Text("A", 20, 150);
        using var input = new MemoryStream(fixture.Create(new[] { body }, imageBytes: pixels));
        using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        var decision = Assert.Single(result.Pages);
        Assert.True(decision.IsNative); Assert.Equal(1, decision.PathObjects); Assert.Equal(1, decision.ImageObjects); Assert.Equal(1, decision.TextObjects);
        output.Position = 0;
        var page = Assert.Single((await new OfdReader().ReadAsync(output)).Pages);
        Assert.Collection(page.Elements, element => Assert.IsType<OfdPathElement>(element), element => Assert.IsType<OfdImageElement>(element),
            element => Assert.IsType<OfdTextElement>(element));
        var image = (OfdImageElement)page.Elements[1];
        using var decoded = Image.Load<Rgb24>(image.Data);
        Assert.Equal(new Rgb24(255, 0, 0), decoded[0, 0]);
        Assert.Equal(new Rgb24(255, 255, 0), decoded[1, 1]);
        Assert.InRange(image.XMillimeters, 24.69, 24.70);
        Assert.InRange(image.YMillimeters, 160.51, 160.52);
    }

    [Theory]
    [InlineData("/SMask 8 0 R")]
    [InlineData("/Decode [1 0 1 0 1 0]")]
    [InlineData("/Filter /FlateDecode")]
    public async Task ImageOutsideOpaqueRawProfileRejectsWholePage(string imageEntries)
    {
        using var fixture = new PdfFixture(Font);
        var body = "20 20 35 25 re f\nq 60 0 0 50 70 90 cm /Im1 Do Q";
        using var input = new MemoryStream(fixture.Create(new[] { body }, imageEntries: imageEntries));
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("IMAGE_PROFILE", error.Message);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task LaterUnsupportedOperationDiscardsEarlierNativePathAndImage()
    {
        using var fixture = new PdfFixture(Font);
        var body = "20 20 35 25 re f q 60 0 0 50 70 90 cm /Im1 Do Q /GS1 gs";
        using var input = new MemoryStream(fixture.Create(new[] { body }));
        using var output = new MemoryStream();
        var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
        var page = Assert.Single(result.Pages);
        Assert.False(page.IsNative); Assert.Equal(0, page.PathObjects); Assert.Equal(1, page.ImageObjects); Assert.Contains("gs", page.Diagnostic);
        output.Position = 0;
        Assert.IsType<OfdImageElement>(Assert.Single((await new OfdReader().ReadAsync(output)).Pages.Single().Elements));
    }
}
