using System.Text;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class ResourceAlternativeTests
{
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    internal static byte[] EncodingCMap() => Encoding.ASCII.GetBytes("/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> def /CMapName /ProbeIdentityH def /CMapType 1 def /WMode 0 def 1 begincodespacerange <0000> <FFFF> endcodespacerange 1 begincidrange <0000> <FFFF> 0 endcidrange endcmap CMapName currentdict /CMap defineresource pop end end");
    internal static byte[] IdentityCidMap() => Enumerable.Range(0, 65536).SelectMany(value => new[] { (byte)(value >> 8), (byte)value }).ToArray();
    private static PdfVectorToOfdOptions Raster() => new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage, Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
    [Theory]
    [InlineData("encoding", "FONT_PROFILE")][InlineData("cid", "CID_MAPPING")]
    public async Task LegalFontStreamAlternativesUsePagePolicyBeforeLoadingFonts(string kind, string code)
    {
        using var fixture = Fixture(); var text = fixture.Text("A  B中文", 20, 85);
        var pdf = fixture.Create(new[] { "10 10 20 20 re f", text, "10 10 20 20 re f" },
            encodingStream: kind == "encoding" ? EncodingCMap() : null, cidMapStream: kind == "cid" ? IdentityCidMap() : null);
        using var failInput = new MemoryStream(pdf); using var failOutput = new MemoryStream(); failOutput.Write(new byte[] { 1, 2 });
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(failInput, failOutput));
        Assert.Contains(code, error.Message); Assert.Equal(new byte[] { 1, 2 }, failOutput.ToArray());
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Raster()).ConvertWithResultAsync(input, output);
        Assert.Equal(new[] { true, false, true }, result.Pages.Select(page => page.IsNative));
        Assert.Equal(1, result.Pages[1].ImageObjects); Assert.Equal(0, result.Pages[1].PathObjects); Assert.Contains(code, result.Pages[1].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.DoesNotContain(package.Fonts, font => font.Data.Length > 0);
        Assert.All(package.Pages[1].Elements.OfType<OfdTextElement>(), text => Assert.Equal(0, text.FillColor.Alpha));
        Assert.Contains("中文", string.Concat(package.Pages[1].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
    }
    [Theory]
    [InlineData(false)][InlineData(true)]
    public async Task LiteralIndexedArrayPreservesPaletteSamples(bool indirect)
    {
        using var fixture = Fixture();
        var pdf = fixture.Create(new[] { "q 420 0 0 595 0 0 cm /Im1 Do Q" }, imageBytes: new byte[] { 0, 1, 1, 0 },
            imageColorSpace: indirect ? "11 0 R" : "[/Indexed /DeviceRGB 1 <d2283c1e82be>]",
            trailingObjects: indirect ? new[] { "[/Indexed /DeviceRGB 1 <d2283c1e82be>]" } : null);
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Raster()).ConvertWithResultAsync(input, output);
        Assert.False(result.Pages[0].IsNative); Assert.Equal(1, result.Pages[0].ImageObjects);
        Assert.StartsWith("PDFV_ORIGINAL_IMAGE_PAGE", result.Pages[0].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        var image = package.Pages[0].Elements.OfType<OfdImageElement>().Single();
        Assert.Contains("PdfInterpolateV1", image.SourceXml ?? "");
        using var decoded = SixLabors.ImageSharp.Image.Load<SixLabors.ImageSharp.PixelFormats.Rgb24>(image.Data);
        Assert.NotEqual(decoded[0, 0], decoded[decoded.Width - 1, 0]); // Actual colored raster, not blank success.
    }
    [Theory]
    [InlineData("99", null)][InlineData(null, "99")]
    public async Task MalformedPrimitiveFontTokensAreFatalRatherThanRasterProfileMiss(string? encoding, string? cid)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("AB", 20, 85) }, encodingOverride: encoding, cidMapOverride: cid));
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(Raster()).ConvertAsync(input, output)); Assert.Equal(0, output.Length);
    }
    [Theory]
    [InlineData("100 0 R", null)]
    [InlineData("11 0 R", new[] { "12 0 R", "11 0 R" })]
    public async Task MissingAndCyclicProfileReferencesCannotEscapeThroughFallback(string encoding, string[]? objects)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("AB", 20, 85) }, encodingOverride: encoding, trailingObjects: objects));
        using var output = new MemoryStream();
        var error = await Record.ExceptionAsync(() => new PdfVectorToOfdConverter(Raster()).ConvertAsync(input, output));
        Assert.NotNull(error); Assert.IsNotType<NotSupportedException>(error); Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task DeepProfileReferenceChainIsInputFailureAndDoesNotRasterize()
    {
        using var fixture = Fixture(); var chain = Enumerable.Range(12, 9).Select(number => number + " 0 R").Append("/Identity-H").ToArray();
        using var input = new MemoryStream(fixture.Create(new[] { fixture.Text("AB", 20, 85) }, encodingOverride: "11 0 R", trailingObjects: chain)); using var output = new MemoryStream();
        var options = Raster(); options.MaxStackDepth = 8;
        var error = await Record.ExceptionAsync(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
        Assert.NotNull(error); Assert.IsNotType<NotSupportedException>(error); Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task NonBooleanInterpolationDeclinesHintRatherThanCoercingFalse()
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { "q 420 0 0 595 0 0 cm /Im1 Do Q" }, interpolateOverride: "0")); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Raster()).ConvertWithResultAsync(input, output);
        Assert.False(result.Pages[0].IsNative); Assert.StartsWith("PDFV_RASTER_PAGE", result.Pages[0].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.DoesNotContain("PdfInterpolateV1", package.Pages[0].Elements.OfType<OfdImageElement>().Single().SourceXml ?? "");
    }
}
