using System.Security.Cryptography;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Fonts;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Font registry")]
public sealed class NameOnlyFontBindingTests
{
    private const string Name = "R08 Pinned Noto Sans Regular";
    private const string Latin = "AV office ffi é";

    [Fact]
    public async Task BoundNameOnlyFacePreservesOfdAndBindsSvgPdfToOneSnapshot()
    {
        var original = LatinFace();
        var expected = (byte[])original.Clone();
        var identity = Convert.ToHexString(SHA256.HashData(expected)).ToLowerInvariant();
        var bindings = new OfdNameOnlyFontBindings(new Dictionary<string, byte[]> { [Name] = original });
        original[100] ^= 0x7f;
        var ofd = await Source(Latin);

        using (var source = new MemoryStream(ofd))
        {
            var package = await new OfdReader().ReadAsync(source);
            Assert.Equal("10", Assert.Single(package.Fonts).Id);
            Assert.Equal(Name, package.Fonts[0].FontName);
            Assert.Empty(package.Fonts[0].Data);
            Assert.Equal(Latin, Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>()).Text);
        }

        using var svg = new MemoryStream();
        await new OfdToSvgConverter(new OfdPackageLoadOptions(), bindings).ConvertAsync(new MemoryStream(ofd), svg);
        var xml = XDocument.Parse(System.Text.Encoding.UTF8.GetString(svg.ToArray()));
        var css = Assert.Single(xml.Descendants().Where(node => node.Name.LocalName == "style")).Value;
        Assert.Contains("ofd-font-" + identity, css);
        Assert.Contains(Convert.ToBase64String(expected), css);
        Assert.Equal("ofd-font-" + identity,
            Assert.Single(xml.Descendants().Where(node => node.Name.LocalName == "text")).Attribute("font-family")?.Value);

        var context = new DocumentFontContext([new OfdFontResource { Id = "10", FontName = Name }],
            nameOnlyBindings: bindings);
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", FontName = Name, Text = Latin }, out _);
        Assert.Equal("ofd-font-" + identity + "-rn", family);
        var selected = GlobalFontSettings.FontResolver.ResolveTypeface(family, false, false);
        Assert.Equal(expected, PdfFontRegistry.GetOriginalFont(selected.FaceName));

        using var pdf = new MemoryStream();
        await new OfdToPdfConverter(new OfdToPdfOptions { NameOnlyFontBindings = bindings })
            .ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.Contains("Ofdrw-" + identity, System.Text.Encoding.ASCII.GetString(pdf.ToArray()));
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal(Latin, parsed.GetPage(1).Text);
        using var fontInspection = PdfReader.Open(new MemoryStream(pdf.ToArray()), PdfDocumentOpenMode.Import);
        var descriptor = Assert.Single(fontInspection.Internals.GetAllObjects().OfType<PdfDictionary>()
            .Where(value => value.Elements.ContainsKey("/FontFile2") &&
                value.Elements.GetName("/FontName").Contains("Ofdrw-" + identity, StringComparison.OrdinalIgnoreCase)));
        // U+00E9 uses a nested composite in this source face. PDFsharpCore's
        // subset omits the nested accent outline while retaining text mapping.
        Assert.Equal(OpenTypeFontIdentity.WithUniqueNames(expected, identity),
            descriptor.Elements.GetDictionary("/FontFile2").Stream.UnfilteredValue);
    }

    [Fact]
    public async Task BoundNameOnlyStrokeAndSimulatedBoldKeepOneSemanticTextRun()
    {
        var face = LatinFace();
        var bindings = new OfdNameOnlyFontBindings(new Dictionary<string, byte[]> { [Name] = face });
        var ofd = await Source("AV", outlined: true);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter(new OfdToPdfOptions { NameOnlyFontBindings = bindings })
            .ConvertAsync(new MemoryStream(ofd), pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray());
        Assert.Equal("AV", parsed.GetPage(1).Text);

        var context = new DocumentFontContext([new OfdFontResource { Id = "10", FontName = Name }],
            nameOnlyBindings: bindings);
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", FontName = Name,
            Text = "AV", Weight = OfdTextElement.BoldWeight }, out _);
        var selected = GlobalFontSettings.FontResolver.ResolveTypeface(family, true, false);
        Assert.True(selected.MustSimulateBold);
        Assert.Equal(face, PdfFontRegistry.GetOriginalFont(selected.FaceName));
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("uncovered")]
    [InlineData("invalid")]
    [InlineData("supplementary")]
    [InlineData("resource-bold")]
    public async Task StrictBindingsRejectUnsupportedTextBeforeWritingEitherOutput(string failure)
    {
        var text = failure switch { "uncovered" => "医院", "supplementary" => "𠮷", _ => Latin };
        var bindings = failure switch
        {
            "missing" => new OfdNameOnlyFontBindings(new Dictionary<string, byte[]>()),
            "invalid" => new OfdNameOnlyFontBindings(new Dictionary<string, byte[]> { [Name] = [1, 2, 3] }),
            _ => new OfdNameOnlyFontBindings(new Dictionary<string, byte[]> { [Name] = LatinFace() })
        };
        var ofd = await Source(text, resourceBold: failure == "resource-bold");
        using var svg = new MemoryStream();
        var svgFailure = await Record.ExceptionAsync(() => new OfdToSvgConverter(new OfdPackageLoadOptions(), bindings)
            .ConvertAsync(new MemoryStream(ofd), svg));
        Assert.True(svgFailure is InvalidDataException or NotSupportedException, svgFailure?.ToString());
        Assert.Equal(0, svg.Length);
        using var pdf = new MemoryStream();
        var pdfFailure = await Record.ExceptionAsync(() => new OfdToPdfConverter(new OfdToPdfOptions { NameOnlyFontBindings = bindings })
            .ConvertAsync(new MemoryStream(ofd), pdf));
        Assert.True(pdfFailure is InvalidDataException or NotSupportedException, pdfFailure?.ToString());
        Assert.Equal(0, pdf.Length);
    }

    private static byte[] LatinFace()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        return File.ReadAllBytes(Path.Combine(directory!.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
    }

    private static async Task<byte[]> Source(string value, bool outlined = false, bool resourceBold = false)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = Name, Bold = resourceBold });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 60 };
        page.Elements.Add(new OfdTextElement { ObjectId = "12", FontResourceId = "10", FontName = Name,
            Text = value, FontSizeMillimeters = 6, XMillimeters = 10, YMillimeters = 12,
            WidthMillimeters = 80, HeightMillimeters = 20,
            Weight = outlined ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
            SourceXml = outlined ? "<TextObject xmlns='http://www.ofdspec.org' Boundary='10 12 80 20' Size='6' Weight='700' " +
                "Stroke='true' Fill='false' LineWidth='0.1'><StrokeColor Value='190 20 30'/>" +
                "<TextCode X='0' Y='6'>AV</TextCode></TextObject>" : null });
        package.Pages.Add(page);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }
}
