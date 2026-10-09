using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class IndexedImageTests
{
    private const string Body = "q 420 0 0 595 0 0 cm /Im1 Do Q";
    private const string Space = "[/Indexed /DeviceRGB 1 <1f5da6e9f1f9>]";
    private static PdfFixture Fixture() => new(File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf")));
    private static PdfVectorToOfdOptions Options() => new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
        Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };

    [Theory]
    [InlineData(null, false)][InlineData(false, false)][InlineData(true, true)]
    public async Task IndexedGridKeepsEveryPalettePixelAndPdfInterpolation(bool? declared, bool expected)
    {
        byte[] indices = { 0, 1, 0, 1, 1, 0 };
        using var fixture = Fixture();
        var source = fixture.Create(new[] { Body }, imageWidth: 3, imageHeight: 2, imageBytes: indices, imageColorSpace: Space, interpolate: declared);
        using var input = new MemoryStream(source); using var ofd = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, ofd);
        var page = Assert.Single(result.Pages); Assert.False(page.IsNative); Assert.StartsWith("PDFV_ORIGINAL_IMAGE_PAGE", page.Diagnostic);
        ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[0].Elements));
        Assert.Equal(expected ? "true" : "false", XElement.Parse(image.SourceXml!).Attribute(XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints"))!.Value);
        var rgb = indices.SelectMany(index => index == 0 ? new byte[] { 31, 93, 166 } : new byte[] { 233, 241, 249 }).ToArray();
        using (var decoded = Image.Load<Rgb24>(image.Data))
        {
            Assert.Equal(3, decoded.Width); Assert.Equal(2, decoded.Height);
            for (var i = 0; i < indices.Length; i++) Assert.Equal(new Rgb24(rgb[i * 3], rgb[i * 3 + 1], rgb[i * 3 + 2]), decoded[i % 3, i / 3]);
        }
        ofd.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var document = PdfDocument.Open(pdf.ToArray()); var exported = Assert.Single(document.GetPage(1).GetImages());
        Assert.Equal(expected, exported.Interpolate); Assert.True(exported.TryGetBytesAsMemory(out var bytes)); Assert.Equal(rgb, bytes.ToArray());
        using var strict = new MemoryStream(source); using var unchanged = new MemoryStream(); unchanged.Write(new byte[] { 7, 8 });
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(strict, unchanged));
        Assert.Equal(new byte[] { 7, 8 }, unchanged.ToArray());
    }

    [Theory]
    [InlineData("[/Indexed /DeviceRGB 1 (\\037\\135\\246\\351\\361\\371)]")]
    [InlineData("[/Indexed /DeviceRGB 0 <1f5da6>]")]
    [InlineData("[/Indexed /DeviceRGB 255 <PALETTE>]")]
    public async Task LiteralAndBoundaryPalettesDecodeTheirOriginalBytes(string space)
    {
        // Maximum hival exercises index 255, rather than truncating to an 8-bit palette length.
        space = space.Replace("PALETTE", string.Concat(Enumerable.Range(0, 256).Select(i => $"{i:X2}5DA6")));
        var last = space.Contains("255") ? (byte)255 : space.Contains(" 0 ") ? (byte)0 : (byte)1;
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageWidth: 3, imageHeight: 2,
            imageBytes: new byte[] { 0, last, 0, last, last, 0 }, imageColorSpace: space)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output); Assert.StartsWith("PDFV_ORIGINAL_IMAGE_PAGE", result.Pages[0].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        using var pixels = Image.Load<Rgb24>(package.Pages[0].Elements.OfType<OfdImageElement>().Single().Data);
        Assert.Equal(space.Contains("255") ? new Rgb24(255, 93, 166) : last == 0 ? new Rgb24(31, 93, 166) : new Rgb24(233, 241, 249), pixels[1, 0]);
    }

    [Theory]
    [InlineData("[/Indexed /DeviceRGB -1 <1f5da6e9f1f9>]", 0)]
    [InlineData("[/Indexed /DeviceRGB 256 <1f5da6e9f1f9>]", 0)]
    [InlineData("[/Indexed /DeviceRGB 1.5 <1f5da6e9f1f9>]", 0)]
    [InlineData("[/Indexed /DeviceRGB 1 <1f5da6>]", 0)]
    [InlineData("[/Indexed /DeviceRGB 0 <1f5da6e9f1f9>]", 0)]
    [InlineData(Space, 2)]
    public async Task MalformedPaletteOrIndexFailsAtomicallyRatherThanRasterizing(string space, byte sample)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageBytes: new[] { sample, (byte)0, (byte)1, (byte)0 }, imageColorSpace: space));
        using var output = new MemoryStream(); output.Write(new byte[] { 7, 8 });
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(Options()).ConvertAsync(input, output));
        Assert.Equal(new byte[] { 7, 8 }, output.ToArray());
    }

    [Theory]
    [InlineData("/Decode [1 0]")][InlineData("/Intent /RelativeColorimetric")][InlineData("/Mask [0 0]")]
    public async Task AdditionalIndexedSemanticsUseExistingFallback(string entries)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageBytes: new byte[] { 0, 1, 1, 0 }, imageColorSpace: Space, imageEntries: entries)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output);
        Assert.StartsWith("PDFV_RASTER_PAGE", result.Pages[0].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.DoesNotContain("PdfInterpolateV1", package.Pages[0].Elements.OfType<OfdImageElement>().Single().SourceXml ?? "");
    }

    [Fact]
    public async Task DifferentIndexedBaseCannotBeReinterpretedAsRgb()
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageBytes: new byte[] { 0, 1, 1, 0 }, imageColorSpace: "[/Indexed /DeviceGray 1 <20e0>]")); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output); Assert.StartsWith("PDFV_RASTER_PAGE", result.Pages[0].Diagnostic);
    }

    [Theory]
    [InlineData("stream-lookup")][InlineData("filter")][InlineData("default-rgb")]
    public async Task ValidIndexedAlternativesDoNotBypassStrictOriginalSampleProfile(string alternative)
    {
        using var fixture = Fixture();
        var source = fixture.Create(new[] { Body },
            imageBytes: alternative == "filter" ? System.Text.Encoding.ASCII.GetBytes("00010100>") : new byte[] { 0, 1, 1, 0 },
            imageColorSpace: alternative == "stream-lookup" ? "[/Indexed /DeviceRGB 1 11 0 R]" : Space,
            trailingObjects: alternative == "stream-lookup" ? new[] { "<< /Length 6 >>\nstream\nabcdef\nendstream" } : null,
            imageEntries: alternative == "filter" ? "/Filter /ASCIIHexDecode" : "",
            resourceEntries: alternative == "default-rgb" ? "/ColorSpace << /DefaultRGB /DeviceRGB >> " : "");
        using var input = new MemoryStream(source); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output);
        Assert.StartsWith("PDFV_RASTER_PAGE", result.Pages[0].Diagnostic);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.DoesNotContain("PdfInterpolateV1", package.Pages[0].Elements.OfType<OfdImageElement>().Single().SourceXml ?? "");
    }

    [Fact]
    public async Task IccBasedImageRemainsOutsideOriginalSampleProfileUnderStrictFail()
    {
        // The strict native parser must reject Do before loading an unsupported ICC resource.
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageColorSpace: "[/ICCBased 11 0 R]", trailingObjects: new[] { "<< /N 3 /Length 0 >>\nstream\n\nendstream" })); using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter().ConvertAsync(input, output)); Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("raw")][InlineData("encoded")][InlineData("pixels")][InlineData("cancel")]
    public async Task IndexedCannotEscapeBudgetsOrCancellationThroughFallback(string gate)
    {
        using var fixture = Fixture(); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageBytes: new byte[] { 0, 1, 1, 0 }, imageColorSpace: Space)); using var output = new MemoryStream();
        var options = Options(); if (gate == "raw") options.Compatibility.MaxTotalImageBytes = 11;
        if (gate == "encoded") options.Compatibility.MaxTotalImageBytes = 12;
        if (gate == "pixels") options.Compatibility.MaxRasterizedPixelsPerPage = 4;
        if (gate == "cancel") await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output, cancellationToken: new CancellationToken(true)));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }
}
