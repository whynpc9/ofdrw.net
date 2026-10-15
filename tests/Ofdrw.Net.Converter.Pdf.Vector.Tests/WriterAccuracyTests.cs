using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class WriterAccuracyTests
{
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static readonly PdfVectorUnsupportedPagePolicy[] Policies = { PdfVectorUnsupportedPagePolicy.Fail, PdfVectorUnsupportedPagePolicy.RasterizePage };
    private static PdfVectorToOfdOptions Options(PdfVectorUnsupportedPagePolicy policy) => new()
    { UnsupportedPagePolicy = policy, Compatibility = new() { PreferExternalPdfToPpm = true } };
    private static string Stroke(string scalar, string matrix, string paint = "S") => $"q {matrix} cm {scalar} w 0 0 m .001 .001 l {paint} Q";
    private static string Text(PdfFixture f, string scalar, string matrix, string text = "AA") =>
        $"q {matrix} cm BT /F1 {scalar} Tf 1 0 0 1 0 0 Tm <{f.Hex(text)}> Tj ET Q";
    private static void Sentinel(MemoryStream output) { Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray()); Assert.Equal(0, output.Position); }

    [Theory]
    [InlineData("0 0 .001 100")][InlineData("0 0 100 .001")]
    [InlineData("0 0 .0014173228336456696 100")][InlineData("0 0 100 .0014173228336456696")]
    public async Task UnrepresentablePageFailsAtomicallyUnderBothPolicies(string box)
    {
        using var f = Fixture(); var bytes = f.Create(new[] { "0 0 .0005 10 re f" }, mediaBox: box);
        foreach (var policy in Policies)
        {
            using var input = new MemoryStream(bytes); using var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
            var error = await Assert.ThrowsAnyAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(Options(policy)).ConvertAsync(input, output));
            Assert.Contains("PAGE_WRITER_PRECISION", error.Message); Sentinel(output);
        }
    }

    [Theory]
    [InlineData("0 0 .003 100")][InlineData("0 0 100 .003")]
    [InlineData("0 0 .0014173228356456696 100")][InlineData("0 0 100 .0014173228356456696")]
    public async Task RepresentablePageBoundaryStillUsesNative(string box)
    {
        using var f = Fixture(); var bytes = f.Create(new[] { "0 0 .0005 10 re f" }, mediaBox: box);
        foreach (var policy in Policies)
        {
            using var input = new MemoryStream(bytes); using var output = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(Options(policy)).ConvertWithResultAsync(input, output);
            Assert.True(result.Pages[0].IsNative); output.Position = 0; var page = (await new OfdReader().ReadAsync(output)).Pages[0];
            Assert.True(page.WidthMillimeters > 0); Assert.True(page.HeightMillimeters > 0);
        }
    }

    [Theory]
    [InlineData(false, ".0006", "100000 0 0 100000 0 0", false)]
    [InlineData(true, ".0006", "100000 0 0 100000 0 0", false)]
    [InlineData(false, ".0006", "1 0 0 1 0 0", false)]
    [InlineData(true, ".0006", "1 0 0 1 0 0", false)]
    [InlineData(false, ".0006", ".001 0 0 .001 0 0", true)]
    [InlineData(true, ".0006", ".001 0 0 .001 0 0", false)]
    [InlineData(false, ".0006", "100000 10000 50000 200000 0 0", false)]
    [InlineData(true, ".0006", "100000 10000 50000 200000 0 0", false)]
    [InlineData(false, "1.000019", "1 0 0 1 0 0", true)]
    [InlineData(false, "1.000021", "1 0 0 1 0 0", false)]
    [InlineData(true, ".00099999", "1 0 0 1 0 0", false)]
    [InlineData(true, ".00100001", "1 0 0 1 0 0", true)]
    [InlineData(false, ".001", "100000 0 0 100000 0 0", true)]
    [InlineData(true, ".001", "100000 0 0 100000 0 0", true)]
    [InlineData(false, ".125", "1 0 0 1 0 0", true)]
    [InlineData(true, "12", "1 0 0 1 0 0", true)]
    public async Task ActualWrittenScalarUsesExistingPageSpaceBounds(bool text, string scalar, string matrix, bool native)
    {
        using var f = Fixture(); var bytes = f.Create(new[] { text ? Text(f, scalar, matrix) : Stroke(scalar, matrix) });
        foreach (var policy in Policies)
        {
            using var input = new MemoryStream(bytes); using var output = new MemoryStream();
            if (!native && policy == PdfVectorUnsupportedPagePolicy.Fail)
            {
                output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
                var error = await Assert.ThrowsAnyAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(Options(policy)).ConvertAsync(input, output));
                Assert.Contains(text ? "TEXT_WRITER_PRECISION" : "STROKE_WRITER_PRECISION", error.Message); Sentinel(output);
            }
            else
            {
                var result = await new PdfVectorToOfdConverter(Options(policy)).ConvertWithResultAsync(input, output);
                Assert.Equal(native, result.Pages[0].IsNative); output.Position = 0; var package = await new OfdReader().ReadAsync(output);
                if (native)
                {
                    Assert.Equal(0, result.Pages[0].ImageObjects);
                    if (text) Assert.Equal("AA", package.Pages[0].Elements.OfType<OfdTextElement>().Single().Text);
                }
                else { Assert.Equal(1, result.Pages[0].ImageObjects); Assert.Empty(package.Pages[0].Elements.OfType<OfdPathElement>()); }
            }
        }
    }

    [Fact]
    public async Task ANewWriterFailureCannotHideLaterCharactersOrPaintCapacity()
    {
        using var f = Fixture();
        foreach (var policy in Policies)
        {
            var o = Options(policy); o.MaxTextCharactersPerPage = 1;
            await Hard(f.Create(new[] { Text(f, ".0006", "100000 0 0 100000 0 0") }), o, "original text character limit");
            o = Options(policy); o.MaxEventsPerPage = 1;
            await Hard(f.Create(new[] { Stroke(".0006", "100000 0 0 100000 0 0", "B") }), o, "draw event limit");
            await Hard(f.Create(new[] { "10 10 20 20 re f " + Text(f, ".0006", "100000 0 0 100000 0 0") }), o, "draw event limit");
            await Hard(f.Create(new[] { "0 0 .0005 10 re f 0 0 .0005 10 re f" }, mediaBox: "0 0 .001 100"), o, "draw event limit");
            o = Options(policy); o.MaxFontReadBytesPerPage = 1;
            await Hard(f.Create(new[] { f.Text("A", 0, 0) }, mediaBox: "0 0 .001 100"), o, "font read");
        }
    }

    [Fact]
    public async Task EarlierNoContentAndUnsupportedOperationsRetainTheirDiagnostics()
    {
        using var f = Fixture();
        foreach (var body in new[] { "", "0 0 m S" })
        {
            using var input = new MemoryStream(f.Create(new[] { body }, mediaBox: "0 0 .001 100")); using var output = new MemoryStream();
            var e = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output)); Assert.Contains("NO_NATIVE_CONTENT", e.Message);
        }
        foreach (var policy in Policies)
        {
            using var input = new MemoryStream(f.Create(new[] { "10 10 20 20 re f .0006 w 0 0 m S BT /F1 .0006 Tf () Tj ET" })); using var output = new MemoryStream();
            var o = Options(policy); o.MaxEventsPerPage = 1; var result = await new PdfVectorToOfdConverter(o).ConvertWithResultAsync(input, output); Assert.True(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].PathObjects);
        }
    }

    private static async Task Hard(byte[] bytes, PdfVectorToOfdOptions options, string message)
    {
        using var input = new MemoryStream(bytes); using var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 }); output.Position = 0;
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output)); Assert.Contains(message, error.Message); Sentinel(output);
    }
}
