using System.Globalization;
using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Font registry")]
public sealed class Audit014RegressionTests
{
    private static byte[] Font => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts/narrow.ttf"));
    private static async Task<byte[]> Input(string codes, byte[]? font = null, bool composite = false)
    {
        var package = new OfdDocumentPackage(); package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit014", Data = font ?? Font });
        var page = new OfdPage { WidthMillimeters = composite ? 200 : 100, HeightMillimeters = composite ? 160 : 80 }; package.Pages.Add(page);
        var textElement = new OfdTextElement { Text = "A", FontResourceId = "10", FontSizeMillimeters = 4,
            XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 60 };
        if (composite)
        {
            package.CompositeResources = new();
            var resource = new OfdCompositeResource { WidthMillimeters = 100, HeightMillimeters = 80 };
            resource.Elements.Add(textElement); package.CompositeResources.Items.Add(resource);
            foreach (var x in new[] { 0d, 80d }) page.Elements.Add(new OfdCompositeElement { Resource = resource,
                XMillimeters = x, YMillimeters = 20, WidthMillimeters = 100, HeightMillimeters = 80 });
        }
        else page.Elements.Add(textElement);
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, true))
        {
            var entry = archive.Entries.Where(e => e.FullName.EndsWith(".xml", StringComparison.Ordinal)).Single(e =>
            { using var stream = e.Open(); return XDocument.Load(stream).Descendants().Any(n => n.Name.LocalName == "TextObject"); });
            XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            var text = xml.Descendants().Single(e => e.Name.LocalName == "TextObject");
            text.Elements().Where(e => e.Name.LocalName == "TextCode").Remove();
            var wrapper = XElement.Parse("<root xmlns=\"" + text.Name.NamespaceName + "\">" + codes + "</root>");
            text.Add(wrapper.Elements()); var entryName = entry.FullName; entry.Delete();
            using var replacement = archive.CreateEntry(entryName).Open(); xml.Save(replacement);
        }
        return output.ToArray();
    }
    private static async Task<byte[]> Pdf(byte[] bytes)
    { using var input = new MemoryStream(bytes); using var output = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, output); return output.ToArray(); }
    private static async Task<XDocument> Svg(byte[] bytes)
    { using var input = new MemoryStream(bytes); using var output = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(input, output); output.Position = 0; return XDocument.Load(output); }
    private static XElement[] Nodes(XDocument svg) => svg.Descendants().Where(e => e.Name.LocalName == "text").ToArray();
    private static double Axis(XElement node, string axis) => double.Parse(node.Attribute(axis)!.Value, CultureInfo.InvariantCulture);

    [Fact]
    public async Task Audit014_NormalControl()
    {
        var input = await Input("<TextCode X=\"2\" Y=\"8\">AB</TextCode>");
        Assert.Equal("AB", Assert.Single(Nodes(await Svg(input))).Value);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(input)); Assert.Equal("AB", pdf.GetPage(1).Text);
    }

    [Fact]
    public async Task Audit014_InheritedAxes_BothRenderers()
    {
        var input = await Input("<TextCode X=\"2\" Y=\"8\" DeltaX=\"3\">AB</TextCode><TextCode X=\"20\">C</TextCode><TextCode Y=\"20\">D</TextCode><TextCode>E</TextCode>");
        var nodes = Nodes(await Svg(input)); Assert.Equal(new[] { 2d, 20, 20, 20 }, nodes.Select(n => Axis(n, "x")));
        Assert.Equal(new[] { 8d, 8, 20, 20 }, nodes.Select(n => Axis(n, "y")));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(input)); var letters = pdf.GetPage(1).Letters;
        var a = letters.Single(l => l.Value == "A"); var c = letters.Single(l => l.Value == "C"); var d = letters.Single(l => l.Value == "D");
        Assert.Equal(a.StartBaseLine.Y, c.StartBaseLine.Y, 4); Assert.Equal(c.StartBaseLine.X, d.StartBaseLine.X, 4);
        Assert.Equal(18 * 72 / 25.4, c.StartBaseLine.X - a.StartBaseLine.X, 3);
    }

    [Fact]
    public async Task Audit014_EscapedSpace_DeltaUsesDecodedGlyphIndex()
    {
        var input = await Input("<TextCode X=\"2\" Y=\"8\" DeltaX=\"3 5\">A\\0020B</TextCode>");
        var node = Assert.Single(Nodes(await Svg(input))); Assert.Equal("A B", node.Value);
        Assert.Equal(new[] { 2d, 5, 10 }, node.Elements().Select(n => Axis(n, "x")));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(input)); Assert.Equal("A B", pdf.GetPage(1).Text);
        Assert.Equal(8 * 72 / 25.4, pdf.GetPage(1).Letters.Single(l => l.Value == "B").StartBaseLine.X - pdf.GetPage(1).Letters.Single(l => l.Value == "A").StartBaseLine.X, 3);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit014_DecodedCoverageSucceeds(bool svg)
    {
        var face = new OpenTypeFace(Font); var cmap = new OpenTypeCmap(face);
        face.Tables["cmap"] = OpenTypeCmap.Build(new Dictionary<int, int> { ['A'] = cmap.Glyph('A') });
        var input = await Input("<TextCode X=\"2\" Y=\"8\">\\0041</TextCode>", face.Build());
        if (svg) Assert.Equal("A", Assert.Single(Nodes(await Svg(input))).Value);
        else { using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(input)); Assert.Equal("A", pdf.GetPage(1).Text); }
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit014_DecodedMissingGlyphFails(bool svg)
    {
        var face = new OpenTypeFace(Font); var cmap = new OpenTypeCmap(face);
        face.Tables["cmap"] = OpenTypeCmap.Build("\\0041".Distinct().ToDictionary(c => (int)c, c => cmap.Glyph(c)));
        var input = await Input("<TextCode X=\"2\" Y=\"8\">\\0041</TextCode>", face.Build());
        var error = await Assert.ThrowsAsync<InvalidDataException>(async () => { if (svg) await Svg(input); else await Pdf(input); });
        Assert.Contains("lacks U+0041", error.Message);
    }

    [Fact]
    public async Task Audit014_RawControl()
    {
        var input = await Input("<TextCode X=\"2\" Y=\"8\">A</TextCode><TextCode X=\"20\">\\0020B</TextCode>");
        using var stream = new MemoryStream(input); var package = await new OfdReader().ReadAsync(stream);
        var text = Assert.IsType<OfdTextElement>(package.Pages[0].Elements[0]);
        Assert.Equal("A\\0020B", text.Text); Assert.Equal(4, text.Runs[1].YMillimeters);
        var identity = package.SourceIdentity; var sourceXml = text.SourceXml; var runs = text.Runs.ToArray();
        using var rendered = new MemoryStream(); await new OfdToPdfConverter().ConvertPackageAsync(package, rendered, null, default);
        Assert.Same(identity, package.SourceIdentity); Assert.Equal(sourceXml, text.SourceXml);
        Assert.Equal(runs, text.Runs); Assert.Equal("A\\0020B", text.Text); Assert.Equal(4, text.Runs[1].YMillimeters);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, saved);
        saved.Position = 0; var reread = await new OfdReader().ReadAsync(saved); var after = Assert.IsType<OfdTextElement>(reread.Pages[0].Elements[0]);
        Assert.Equal(text.Text, after.Text); Assert.Equal(4, after.Runs[1].YMillimeters);
    }

    [Fact]
    public async Task Audit014_FreshLiteralSlash_AndExplicitModelEdits()
    {
        var fresh = new OfdTextElement { Text = @"A\0020B" };
        Assert.Same(fresh, OfdTextRenderProjection.Create(fresh, new(new OfdModelLimits()), default));
        var package = new OfdDocumentPackage(); package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit014", Data = Font });
        fresh.FontResourceId = "10"; package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80, Elements = { fresh } });
        using var output = new MemoryStream(); await new OfdToPdfConverter().ConvertPackageAsync(package, output, null, default);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(output.ToArray()); Assert.Equal(@"A\0020B", pdf.GetPage(1).Text);
        using var input = new MemoryStream(await Input("<TextCode X=\"2\" Y=\"8\">A</TextCode>"));
        var read = await new OfdReader().ReadAsync(input); var text = Assert.IsType<OfdTextElement>(read.Pages[0].Elements[0]);
        text.Text = "B";
        Assert.Throws<InvalidDataException>(() => OfdTextRenderProjection.Create(text, new(new OfdModelLimits()), default));
        text.Runs[0].Text = "B"; text.Runs[0].YMillimeters = 12;
        Assert.Same(text, OfdTextRenderProjection.Create(text, new(new OfdModelLimits()), default));
    }

    [Fact]
    public async Task Audit014_EmptySlots_CombiningGlyphs_InheritStart()
    {
        var face = new OpenTypeFace(Font); var cmap = new OpenTypeCmap(face);
        face.Tables["cmap"] = OpenTypeCmap.Build(new Dictionary<int, int> { ['A'] = cmap.Glyph('A'), ['B'] = cmap.Glyph('B'), [0x301] = cmap.Glyph('A') });
        var input = await Input("<TextCode X=\"2\" Y=\"8\"/><TextCode X=\"10\"/><TextCode DeltaX=\"7\">A\\0301B</TextCode>", face.Build());
        var nodes = Nodes(await Svg(input)); Assert.Equal(3, nodes.Length); Assert.Equal("A\u0301B", nodes[2].Value);
        Assert.Equal(new[] { 10d, 17 }, nodes[2].Elements().Select(n => Axis(n, "x")));
        Assert.All(nodes, n => Assert.Equal(8, Axis(n, "y")));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(input)); Assert.Contains("B", pdf.GetPage(1).Text);
        using var stream = new MemoryStream(input); var model = await new OfdReader().ReadAsync(stream);
        var raw = Assert.IsType<OfdTextElement>(model.Pages[0].Elements[0]); Assert.Equal(4, raw.Runs[2].YMillimeters);
        Assert.Equal(@"A\0301B", raw.Runs[2].Text);
    }

    [Theory]
    [InlineData("\\00Q1")] [InlineData("\\002")] [InlineData("\\D800")] [InlineData("\\0001")]
    public async Task Audit014_UnsupportedSourceEncodingFails(string value)
    {
        var input = await Input("<TextCode X=\"2\" Y=\"8\">" + value + "</TextCode>");
        await Assert.ThrowsAsync<InvalidDataException>(() => Svg(input));
        await Assert.ThrowsAsync<InvalidDataException>(() => Pdf(input));
    }

    [Theory]
    [InlineData(false, false)] [InlineData(true, false)] [InlineData(false, true)] [InlineData(true, true)]
    public async Task Audit014_CompositeRepeatedOccurrences_DecodedCoverageAndAxes(bool svg, bool missing)
    {
        var face = new OpenTypeFace(Font); var cmap = new OpenTypeCmap(face);
        var characters = missing ? @"\0041B" : "AB";
        face.Tables["cmap"] = OpenTypeCmap.Build(characters.Distinct().ToDictionary(c => (int)c, c => cmap.Glyph(c)));
        var input = await Input("<TextCode X=\"2\" Y=\"8\">\\0041</TextCode><TextCode X=\"20\">B</TextCode>", face.Build(), composite: true);
        if (missing) { var error = await Assert.ThrowsAsync<InvalidDataException>(async () => { if (svg) await Svg(input); else await Pdf(input); }); Assert.Contains("lacks U+0041", error.Message); return; }
        if (svg)
        {
            var nodes = Nodes(await Svg(input)); Assert.Equal(new[] { "A", "B", "A", "B" }, nodes.Select(n => n.Value));
            Assert.Equal(new[] { 2d, 20, 2, 20 }, nodes.Select(n => Axis(n, "x"))); Assert.All(nodes, n => Assert.Equal(8, Axis(n, "y")));
        }
        else
        {
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(input)); var letters = pdf.GetPage(1).Letters;
            Assert.Equal(2, letters.Count(l => l.Value == "A")); Assert.Equal(2, letters.Count(l => l.Value == "B"));
            Assert.All(letters, l => Assert.Equal(letters[0].StartBaseLine.Y, l.StartBaseLine.Y, 4));
            Assert.Equal(80 * 72 / 25.4, letters.Where(l => l.Value == "A").Max(l => l.StartBaseLine.X) - letters.Where(l => l.Value == "A").Min(l => l.StartBaseLine.X), 3);
        }
        using var source = new MemoryStream(input); var package = await new OfdReader().ReadAsync(source);
        var definition = Assert.IsType<OfdTextElement>(Assert.Single(package.CompositeResources!.Items).Elements[0]);
        Assert.Equal(@"\0041B", definition.Text); Assert.Equal(4, definition.Runs[1].YMillimeters);
        var original = definition.SourceXml; var material = new OfdModelMaterialization(new OfdModelLimits());
        Assert.Equal(2, OfdCompositeProjection.Expand(package.Pages[0].Elements, material, default).Count);
        Assert.Equal(original, definition.SourceXml); Assert.Equal(@"\0041", definition.Runs[0].Text);
    }

    [Theory]
    [InlineData(@"C:\work\file", false)] [InlineData(@"C:\work\file", true)]
    [InlineData(@"A\0020B", false)] [InlineData(@"A\0020B", true)]
    public async Task Audit014_FreshWriterRoundtrip_PreservesLiteralText(string literal, bool run)
    {
        var package = new OfdDocumentPackage(); package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit014", Data = Font });
        var text = new OfdTextElement { Text = literal, FontResourceId = "10", FontSizeMillimeters = 4 };
        if (run) text.Runs.Add(new OfdTextRun { Text = literal, XMillimeters = 0, YMillimeters = 4 });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80, Elements = { text } });
        using var direct = new MemoryStream(); await new OfdToPdfConverter().ConvertPackageAsync(package, direct, null, default);
        using var directPdf = UglyToad.PdfPig.PdfDocument.Open(direct.ToArray()); Assert.Equal(literal, directPdf.GetPage(1).Text);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, saved); var bytes = saved.ToArray();
        Assert.Equal(literal, text.Text); if (run) Assert.Equal(literal, text.Runs[0].Text); Assert.Null(text.SourceXml);
        saved.Position = 0; var read = await new OfdReader().ReadAsync(saved);
        Assert.Equal(literal, new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(read));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(bytes)); Assert.Equal(literal, pdf.GetPage(1).Text);
        Assert.Equal(literal, string.Concat(Nodes(await Svg(bytes)).Select(n => n.Value)));
        using var second = new MemoryStream(); await new OfdPackageWriter().WriteAsync(read, second);
        using var pdf2 = UglyToad.PdfPig.PdfDocument.Open(await Pdf(second.ToArray())); Assert.Equal(literal, pdf2.GetPage(1).Text);
    }

    [Theory]
    [InlineData(@"C:\work\file", false)] [InlineData(@"C:\work\file", true)]
    [InlineData(@"A\0020B", false)] [InlineData(@"A\0020B", true)]
    public async Task Audit014_DocxNativeAndDefault_PreserveLiteralText(string literal, bool explicitNative)
    {
        using var docx = new MemoryStream();
        using (var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Create(docx, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            var part = document.AddMainDocumentPart(); part.Document = new DocumentFormat.OpenXml.Wordprocessing.Document(
                new DocumentFormat.OpenXml.Wordprocessing.Body(new DocumentFormat.OpenXml.Wordprocessing.Paragraph(
                    new DocumentFormat.OpenXml.Wordprocessing.Run(new DocumentFormat.OpenXml.Wordprocessing.Text(literal)))));
        }
        docx.Position = 0; using var ofd = new MemoryStream();
        var options = new Ofdrw.Net.Converter.Docx.DocxConversionOptions();
        if (explicitNative) options.OfdMode = Ofdrw.Net.Converter.Docx.DocxToOfdMode.Native;
        await new Ofdrw.Net.Converter.Docx.Converters.DocxToOfdConverter(options).ConvertAsync(docx, ofd);
        ofd.Position = 0; var model = await new OfdReader().ReadAsync(ofd);
        Assert.Equal(literal, new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(model));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(ofd.ToArray())); Assert.Equal(literal, pdf.GetPage(1).Text);
        Assert.Equal(literal, string.Concat(Nodes(await Svg(ofd.ToArray())).Select(n => n.Value)));
    }

    [Theory]
    [InlineData(@"C:\work\file")] [InlineData(@"A\0020B")]
    public async Task Audit014_FreshCompositeLiteral_IsDecodedExactlyOnce(string literal)
    {
        var package = new OfdDocumentPackage { CompositeResources = new() }; package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit014", Data = Font });
        var definition = new OfdTextElement { Text = literal, FontResourceId = "10", FontSizeMillimeters = 4 };
        definition.Runs.Add(new OfdTextRun { Text = literal, XMillimeters = 0, YMillimeters = 4 });
        var resource = new OfdCompositeResource { WidthMillimeters = 80, HeightMillimeters = 20 }; resource.Elements.Add(definition);
        package.CompositeResources.Items.Add(resource); var page = new OfdPage { WidthMillimeters = 200, HeightMillimeters = 80 }; package.Pages.Add(page);
        foreach (var x in new[] { 10d, 100d }) page.Elements.Add(new OfdCompositeElement { Resource = resource, XMillimeters = x,
            YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 20 });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); var bytes = output.ToArray();
        output.Position = 0; var read = await new OfdReader().ReadAsync(output);
        Assert.Equal(literal + literal, new Ofdrw.Net.Reader.Extraction.OfdTextExtractor().Extract(read));
        Assert.Equal(literal + literal, string.Concat(Nodes(await Svg(bytes)).Select(n => n.Value)));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(await Pdf(bytes)); Assert.Equal(literal + literal, pdf.GetPage(1).Text);
        var raw = Assert.IsType<OfdTextElement>(Assert.Single(read.CompositeResources!.Items).Elements[0]);
        Assert.Equal(literal.Replace("\\", "\\005C"), raw.Text);
    }

    [Fact]
    public async Task Audit014_VisualSample()
    {
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is not { Length: > 0 } directory) return;
        Directory.CreateDirectory(directory);
        var bytes = await Input("<TextCode X=\"2\" Y=\"8\" DeltaX=\"3 4\">A\\0020B</TextCode><TextCode X=\"30\">C</TextCode><TextCode Y=\"25\">D</TextCode><TextCode X=\"45\">E</TextCode>", TrustedFontFixtures.Read("latin"));
        await File.WriteAllBytesAsync(Path.Combine(directory, "source-projection.ofd"), bytes);
        await File.WriteAllBytesAsync(Path.Combine(directory, "source-projection.pdf"), await Pdf(bytes));
        (await Svg(bytes)).Save(Path.Combine(directory, "source-projection.svg"));
    }
}
