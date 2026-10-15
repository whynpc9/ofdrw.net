using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class WriterPrecisionTests
{
    [Fact]
    public async Task WriterMismatchCannotBypassTheRemainingCharactersInItsRun()
    {
        using var fixture = Fixture();
        using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("AA", 20, 85, .0004) }));
        using var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 });
        var options = new PdfVectorToOfdOptions { MaxTextCharactersPerPage = 1, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage };
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WriterMismatchCannotConvertAnExhaustedEventBudgetIntoFallback(bool text)
    {
        using var fixture = Fixture();
        var paint = text ? Paint(fixture, true, ".0004", "1") : ".0004 w 30 100 30 20 re B";
        using var input = new MemoryStream(fixture.Create(new[] { "10 10 20 20 re f " + paint }));
        using var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 });
        var options = new PdfVectorToOfdOptions { MaxEventsPerPage = 1, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage };
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray());
    }

    [Fact]
    public async Task CombinedFillStrokeWithWriterZeroUsesWholePageFallback()
    {
        using var fixture = Fixture();
        using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("A  B中文", 20, 85) + " .0004 w 30 100 30 20 re B" }));
        using var output = new MemoryStream();
        var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
        Assert.Contains("STROKE_WRITER_PRECISION", result.Pages[0].Diagnostic);
        Assert.Equal(1, result.Pages[0].ImageObjects); Assert.Equal(0, result.Pages[0].PathObjects);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.All(package.Pages[0].Elements.OfType<OfdTextElement>(), value => Assert.Equal(0, value.FillColor.Alpha));
    }

    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static string Paint(PdfFixture fixture, bool text, string size, string scale)
    {
        var content = text ? $"BT /F1 {size} Tf 1 0 0 1 30 100 Tm <{fixture.Hex("A")}> Tj ET"
            : $"{size} w 30 100 m 60 100 l S";
        return $"q {scale} 0 0 {scale} 0 0 cm {content} Q";
    }

    [Theory]
    [InlineData(false, ".0004", "1", "STROKE_WRITER_PRECISION")]
    [InlineData(false, ".00049999", "1", "STROKE_WRITER_PRECISION")]
    [InlineData(false, ".0004", "128", "STROKE_WRITER_PRECISION")]
    [InlineData(true, ".0004", "1", "TEXT_WRITER_PRECISION")]
    [InlineData(true, ".00049999", "1", "TEXT_WRITER_PRECISION")]
    [InlineData(true, ".0004", "128", "TEXT_WRITER_PRECISION")]
    [InlineData(false, ".000499999999", "1", "STROKE_WRITER_PRECISION")]
    [InlineData(false, ".0005", "1", "STROKE_WRITER_PRECISION")]
    [InlineData(false, ".00050001", "1", "STROKE_WRITER_PRECISION")]
    [InlineData(true, ".000499999999", "1", "TEXT_WRITER_PRECISION")]
    [InlineData(true, ".0005", "1", "TEXT_WRITER_PRECISION")]
    [InlineData(true, ".00050001", "1", "TEXT_WRITER_PRECISION")]
    [InlineData(false, "0", "1", "STROKE_PROFILE")]
    public async Task RejectedWrittenScalarUsesWholePagePolicyBeforeOutput(bool text, string size, string scale, string reason)
    {
        using var fixture = Fixture();
        var normal = fixture.Text("A  B中文", 20, 85);
        var pdf = fixture.Create(new[] { normal, normal + " 10 10 20 20 re f " + Paint(fixture, text, size, scale), normal });
        using var failInput = new MemoryStream(pdf); using var sentinel = new MemoryStream();
        sentinel.Write(new byte[] { 4, 5, 6 });
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(failInput, sentinel));
        Assert.Contains(reason, error.Message); Assert.Equal(new byte[] { 4, 5, 6 }, sentinel.ToArray());
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, output);
        Assert.Equal(new[] { true, false, true }, result.Pages.Select(page => page.IsNative));
        Assert.Contains(reason, result.Pages[1].Diagnostic);
        Assert.Equal(1, result.Pages[1].ImageObjects); Assert.Equal(0, result.Pages[1].PathObjects);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.All(package.Pages[1].Elements.OfType<OfdTextElement>(), value => Assert.Equal(0, value.FillColor.Alpha));
        Assert.Equal("A  B中文", package.Pages[2].Elements.OfType<OfdTextElement>().Single().Text);
        Assert.Single(package.Pages[1].Elements.OfType<OfdImageElement>());
    }

    [Theory]
    [InlineData(false, ".001", ".001")]
    [InlineData(true, ".001", ".001")]
    public async Task SmallPhysicalSizeWithAccurateWrittenScalarStaysNative(bool text, string size, string scale)
    {
        // A positive written scalar must also meet the page-space accuracy
        // certificate. These exact local values retain their small CTM control.
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Paint(fixture, text, size, scale) }));
        using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.True(result.Pages[0].IsNative); Assert.Equal(0, result.Pages[0].ImageObjects);
        Assert.Equal(text ? 1 : 0, result.Pages[0].TextObjects); Assert.Equal(text ? 0 : 1, result.Pages[0].PathObjects);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        if (text) Assert.True(package.Pages[0].Elements.OfType<OfdTextElement>().Single().FontSizeMillimeters > 0);
        else Assert.True(package.Pages[0].Elements.OfType<OfdPathElement>().Single().LineWidthMillimeters > 0);
    }
}
