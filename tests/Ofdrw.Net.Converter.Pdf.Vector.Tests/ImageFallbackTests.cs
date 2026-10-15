using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using SkiaSharp;
using UglyToad.PdfPig;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class ImageFallbackTests
{
    private static readonly XName Hint = XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints");
    private const string Body = "q 420 0 0 595 0 0 cm /Im1 Do Q";
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));
    private static PdfVectorToOfdOptions Options() => new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
        Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
    [Theory]
    [InlineData(null, false)][InlineData(false, false)][InlineData(true, true)]
    public async Task PureImageKeepsAsymmetricRgbGridAndEffectiveSourceInterpolation(bool? declared, bool expected)
    {
        byte[] original = { 255, 0, 0, 0, 255, 0, 0, 0, 255, 255, 255, 0, 0, 0, 0, 255, 255, 255 };
        using var fixture = new PdfFixture(Font); using var input = new MemoryStream(fixture.Create(new[] { Body }, interpolate: declared,
            imageWidth: 3, imageHeight: 2, imageBytes: original)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output);
        var decision = result.Pages.Single(); Assert.False(decision.IsNative); Assert.Contains("PDFV_ORIGINAL_IMAGE_PAGE", decision.Diagnostic);
        Assert.Equal(1, decision.ImageObjects); Assert.Equal(0, decision.PathObjects); Assert.Equal(0, decision.TextObjects);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output); var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[0].Elements));
        Assert.Equal(expected ? "true" : "false", XElement.Parse(image.SourceXml!).Attribute(Hint)!.Value);
        using (var pixels = Image.Load<Rgb24>(image.Data))
        {
            Assert.Equal(3, pixels.Width); Assert.Equal(2, pixels.Height);
            for (var index = 0; index < 6; index++)
            { var pixel = pixels[index % 3, index / 3]; Assert.Equal(original[index * 3], pixel.R); Assert.Equal(original[index * 3 + 1], pixel.G); Assert.Equal(original[index * 3 + 2], pixel.B); }
        }
        await AssertExport(package, expected);
    }
    [Theory]
    [InlineData("q 419 0 0 595 0 0 cm /Im1 Do Q")]
    [InlineData("q -420 0 0 595 420 0 cm /Im1 Do Q")]
    [InlineData("q 420 0 0 595 0 0 cm /Im1 Do /Im1 Do Q")]
    [InlineData("q 420 0 0 595 0 0 cm /Im1 Do Q 99")]
    [InlineData("q 420 0 0 595 0 0 cm /Im1 Do Q /dangling")]
    [InlineData("99 q 420 0 0 595 0 0 cm /Im1 Do Q")]
    [InlineData("q 0 0 100 100 re W n 420 0 0 595 0 0 cm /Im1 Do Q")]
    [InlineData("/GS1 gs q 420 0 0 595 0 0 cm /Im1 Do Q")]
    public async Task ExtraTokensStateOrInexactShapeNeverEnterOriginalImageBranch(string body)
    {
        using var fixture = new PdfFixture(Font); using var input = new MemoryStream(fixture.Create(new[] { body })); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output);
        Assert.DoesNotContain("PDFV_ORIGINAL_IMAGE_PAGE", result.Pages.Single().Diagnostic);
        Assert.Contains("PDFV_RASTER_PAGE", result.Pages.Single().Diagnostic);
    }
    [Theory]
    [InlineData("/Decode [1 0 1 0 1 0]")][InlineData("/Intent /RelativeColorimetric")][InlineData("/OC << /Type /OCG /Name (layer) >>")]
    public async Task ImageSemanticsBeyondRawRgbProfileUseExistingRasterFallback(string entries)
    {
        using var fixture = new PdfFixture(Font); using var input = new MemoryStream(fixture.Create(new[] { Body }, imageEntries: entries)); using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(Options()).ConvertWithResultAsync(input, output);
        Assert.Contains("PDFV_RASTER_PAGE", result.Pages.Single().Diagnostic);
    }
    [Theory]
    [InlineData("operations")][InlineData("pixels")][InlineData("bytes")][InlineData("repeat")]
    public async Task OriginalImageCannotBypassOperationPixelOrAccumulatedByteLimits(string limit)
    {
        using var fixture = new PdfFixture(Font); using var input = new MemoryStream(fixture.Create(new[] { Body })); using var output = new MemoryStream(); var options = Options();
        if (limit == "operations") options.MaxOperationsPerPage = 3;
        if (limit == "pixels") options.Compatibility.MaxRasterizedPixelsPerPage = 4; // Source fits, physical target viewport does not.
        if (limit == "bytes") options.Compatibility.MaxTotalImageBytes = 11;
        if (limit == "repeat") options.Compatibility.MaxTotalImageBytes = 150;
        await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(options).ConvertAsync(input, output, limit == "repeat" ? new[] { 0, 0 } : null));
        Assert.Equal(0, output.Length);
    }
    [Fact]
    public async Task ValidHintSurvivesRewriteMergeMixSplitAndResourcePruning()
    {
        using var fixture = new PdfFixture(Font); using var input = new MemoryStream(fixture.Create(new[] { Body, Body }, interpolate: false)); using var ofd = new MemoryStream();
        await new PdfVectorToOfdConverter(Options()).ConvertAsync(input, ofd); ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        var merge = OfdDocumentMerger.Merge(new[] { package });
        var mix = OfdDocumentMixer.Mix(new[] { new OfdMixSource(package, 0), new OfdMixSource(package, 1) });
        var split = OfdDocumentSplitter.Split(package, new[] { 1 });
        package.Pages.RemoveAt(0); // Writer pruning must keep the selected image and private hint.
        foreach (var candidate in new[] { merge, mix, split, package })
        {
            using var rewritten = new MemoryStream(); await new OfdPackageWriter().WriteAsync(candidate, rewritten); rewritten.Position = 0;
            var read = await new OfdReader().ReadAsync(rewritten);
            Assert.All(read.Pages.SelectMany(page => page.Elements.OfType<OfdImageElement>()), image => Assert.Equal("false", XElement.Parse(image.SourceXml!).Attribute(Hint)!.Value));
            await AssertExport(read, false);
        }
    }
    [Theory]
    [InlineData("True")][InlineData(" false ")][InlineData("1")][InlineData("bad")]
    public async Task MalformedExactHintFailsExportAndIsNotMergeWhitelisted(string value)
    {
        var package = Package(value, Hint);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge(new[] { package }));
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0; using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf)); Assert.Equal(0, pdf.Length);
        ofd.Position = 0; using var svg = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToSvgConverter().ConvertAsync(ofd, svg)); Assert.Equal(0, svg.Length);
    }
    [Fact]
    public async Task UntaggedOrWrongNamespaceKeepHistoricExportDefaultAndUnknownVendorStaysRejected()
    {
        await AssertExport(Package(null, Hint), true, expectSvgHint: false);
        var unknown = Package("false", XName.Get("PdfInterpolateV1", "urn:unknown-vendor"));
        await AssertExport(unknown, true, expectSvgHint: false);
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge(new[] { unknown }));
        var unknownChild = Package(null, Hint); var image = (OfdImageElement)unknownChild.Pages[0].Elements[0];
        image.SourceXml = new XElement(XName.Get("ImageObject", unknownChild.Options.Namespace), new XElement(XName.Get("Unknown", "urn:vendor"))).ToString();
        Assert.Throws<NotSupportedException>(() => OfdDocumentMerger.Merge(new[] { unknownChild }));
    }
    private static OfdDocumentPackage Package(string? hint, XName name)
    {
        using var bitmap = new SKBitmap(2, 2); bitmap.Erase(SKColors.Red); using var data = bitmap.Encode(SKEncodedImageFormat.Png, 100);
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 50, HeightMillimeters = 70 }; package.Pages.Add(page);
        var xml = new XElement(XName.Get("ImageObject", package.Options.Namespace)); if (hint is not null) xml.SetAttributeValue(name, hint);
        page.Elements.Add(new OfdImageElement { Data = data.ToArray(), WidthMillimeters = 50, HeightMillimeters = 70, FileName = "image.png", MediaType = "image/png", SourceXml = xml.ToString() }); return package;
    }
    private static async Task AssertExport(OfdDocumentPackage package, bool interpolate, bool expectSvgHint = true)
    {
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0; using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdf); using var document = PdfDocument.Open(pdf.ToArray());
        for (var index = 1; index <= document.NumberOfPages; index++) Assert.All(document.GetPage(index).GetImages(), image => Assert.Equal(interpolate, image.Interpolate));
        ofd.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(ofd, svg);
        svg.Position = 0; var svgXml = XDocument.Load(svg); var images = svgXml.Descendants().Where(node => node.Name.LocalName == "image");
        Assert.All(images, image => Assert.Equal(expectSvgHint ? interpolate ? "smooth" : "crisp-edges" : null, image.Attribute("image-rendering")?.Value));
    }
}
