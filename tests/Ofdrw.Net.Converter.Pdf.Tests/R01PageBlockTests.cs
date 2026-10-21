using System.Text;
using System.IO.Compression;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class R01PageBlockTests
{
    private const string Ns = "http://www.ofdspec.org";
    private static void Retain(string name, byte[] ofd, byte[] pdf, byte[] svg, byte[] html)
    {
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is not { Length: > 0 } dir) return;
        Directory.CreateDirectory(dir);
        File.WriteAllBytes(Path.Combine(dir, name + ".ofd"), ofd);
        File.WriteAllBytes(Path.Combine(dir, name + ".pdf"), pdf);
        File.WriteAllBytes(Path.Combine(dir, name + ".svg"), svg);
        File.WriteAllBytes(Path.Combine(dir, name + ".html"), html);
    }

    private static async Task<byte[]> Source(string body, bool template = false, string? resources = null, bool font = false)
    {
        var package = new OfdDocumentPackage();
        if (font)
        {
            // The synthetic template is inserted after the first write; retain
            // its font in full because that writer cannot see the future root.
            package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "Ofdrw.Net.sln"))) dir = dir.Parent;
            var path = Path.Combine(dir!.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf");
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = File.ReadAllBytes(path) });
        }
        var page = new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var block = new OfdRawElement { LocalName = "PageBlock", Xml = $"<PageBlock xmlns='{Ns}' ID='500'>{body}</PageBlock>" };
        if (!template) page.Elements.Add(block);
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        if (!template && resources is null) return output.ToArray();
        using var revised = new MemoryStream();
        using (var inputZip = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read))
        using (var outputZip = new ZipArchive(revised, ZipArchiveMode.Create, leaveOpen: true))
        {
            XNamespace ns = Ns;
            foreach (var entry in inputZip.Entries)
            {
                using var entryData = entry.Open(); using var buffer = new MemoryStream(); await entryData.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName == "Doc_0/Document.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    if (template) xml.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage",
                        new XAttribute("ID", "600"), new XAttribute("BaseLoc", "Tpl_0/Content.xml"), new XAttribute("ZOrder", "Background")));
                    if (resources is not null) xml.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "DocumentRes", "DocumentRes.xml"));
                    xml.Root!.Element(ns + "CommonData")!.Element(ns + "MaxUnitID")!.Value = "601";
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                if (template && entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    xml.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "600")));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = outputZip.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
            if (template)
            {
                var templateXml = new XDocument(new XElement(ns + "Page", new XElement(ns + "Content",
                    new XElement(ns + "Layer", new XAttribute("ID", "601"), XElement.Parse(block.Xml)))));
                using var templateStream = outputZip.CreateEntry("Doc_0/Tpl_0/Content.xml").Open();
                await templateStream.WriteAsync(Encoding.UTF8.GetBytes(templateXml.ToString(SaveOptions.DisableFormatting)));
            }
            if (resources is not null)
            {
                var wrapper = XElement.Parse($"<R xmlns='{Ns}'>{resources}</R>");
                var resourceXml = new XDocument(new XElement(ns + "Res", wrapper.Elements()));
                using var resourceStream = outputZip.CreateEntry("Doc_0/DocumentRes.xml").Open();
                await resourceStream.WriteAsync(Encoding.UTF8.GetBytes(resourceXml.ToString(SaveOptions.DisableFormatting)));
            }
        }
        return revised.ToArray();
    }

    [Fact]
    public async Task ReferencedPaletteAndInheritedDrawParamResolveInAllOutputs()
    {
        var resources = "<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette><CV>255 0 0</CV><CV>0 0 255</CV></Palette></ColorSpace>" +
            "<ColorSpace ID='21' Type='CMYK'/></ColorSpaces>" +
            "<DrawParams><DrawParam ID='30'><FillColor ColorSpace='20' Index='1'/></DrawParam>" +
            "<DrawParam ID='31' Relative='30' LineWidth='0.5'><StrokeColor Value='0 255 255'/></DrawParam></DrawParams>";
        var body = "<PathObject ID='501' Boundary='10 10 30 30' DrawParam='31' Fill='true' Stroke='true'>" +
            "<AbbreviatedData>M 0 0 L 30 0 L 30 30 L 0 30 C</AbbreviatedData></PathObject>" +
            "<PathObject ID='502' Boundary='50 10 30 30' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='21' Value='0 255 255 0'/><AbbreviatedData>M 0 0 L 30 0 L 30 30 L 0 30 C</AbbreviatedData></PathObject>";
        var bytes = await Source(body, resources: resources);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        var paths = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants().Where(n => n.Name.LocalName == "path").ToArray();
        Assert.Equal(2, paths.Length);
        Assert.Equal("rgb(0,0,255)", paths[0].Attribute("fill")?.Value);
        Assert.Equal("rgb(0,255,255)", paths[0].Attribute("stroke")?.Value);
        Assert.Equal("rgb(255,0,0)", paths[1].Attribute("fill")?.Value);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        Assert.True(pdf.Length > 0);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("rgb(0,0,255)", Encoding.UTF8.GetString(html.ToArray()));
        var imported = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        using var written = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, written);
        using var again = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(written.ToArray()), again);
        Assert.Equal(XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).ToString(SaveOptions.DisableFormatting),
            XDocument.Parse(Encoding.UTF8.GetString(again.ToArray())).ToString(SaveOptions.DisableFormatting));
        Retain("palette-drawparam", bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ForeignResourceNamespaceDoesNotResolvePaint(bool drawParam)
    {
        var resources = drawParam
            ? "<DrawParams><DrawParam ID='30'><FillColor Value='0 0 255'/></DrawParam></DrawParams>"
            : "<ColorSpaces><ColorSpace ID='20' Type='RGB'/></ColorSpaces>";
        var style = drawParam ? "DrawParam='30'" : "";
        var color = drawParam ? "" : "<FillColor ColorSpace='20' Value='0 0 255'/>";
        var body = $"<PathObject ID='501' Boundary='10 10 20 20' Fill='true' {style}>{color}" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var source = await Source(body, resources: resources);
        using var patched = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(source), ZipArchiveMode.Read))
        using (var output = new ZipArchive(patched, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var data = entry.Open(); using var buffer = new MemoryStream(); await data.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName == "Doc_0/DocumentRes.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    XNamespace extension = "urn:r01:paint-extension";
                    foreach (var node in xml.Descendants())
                    {
                        foreach (var declaration in node.Attributes().Where(a => a.IsNamespaceDeclaration).ToArray())
                            declaration.Remove();
                        node.Name = extension + node.Name.LocalName;
                    }
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var ofd = patched.ToArray();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData("<CV Extra='x'>0 0 255</CV>")]
    [InlineData("<CV><VendorChild>0 0 255</VendorChild></CV>")]
    [InlineData("<CV><?vendor key?>0 0 255</CV>")]
    public async Task ExtendedPaletteEntryFailsAllRenderers(string entry)
    {
        var resources = "<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette><CV>0 0 0</CV>" +
            entry + "</Palette></ColorSpace></ColorSpaces>";
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true'>" +
            "<FillColor ColorSpace='20' Index='1'/>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html));
        Assert.Equal(0, html.Length);
    }

    [Fact]
    public async Task TopLevelGrayAndCmykPathsRemainRenderableAfterWriteAndReread()
    {
        var resources = "<ColorSpaces><ColorSpace ID='20' Type='GRAY'/><ColorSpace ID='21' Type='CMYK'/></ColorSpaces>";
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='20' Value='80'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>" +
            "<PathObject ID='502' Boundary='40 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='21' Value='0 255 255 0'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var nested = await Source(body, resources: resources);
        using var topLevel = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(nested), ZipArchiveMode.Read))
        using (var output = new ZipArchive(topLevel, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var layer = xml.Descendants().Single(node => node.Name.LocalName == "Layer");
                    var block = layer.Elements().Single(node => node.Name.LocalName == "PageBlock");
                    block.ReplaceWith(block.Elements().ToArray());
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var imported = await new OfdReader().ReadAsync(new MemoryStream(topLevel.ToArray()));
        using var rewritten = new MemoryStream(); await new OfdPackageWriter().WriteAsync(imported, rewritten);
        using (var archive = new ZipArchive(new MemoryStream(rewritten.ToArray()), ZipArchiveMode.Read))
        using (var content = archive.Entries.Single(entry => entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal)).Open())
        {
            var xml = XDocument.Load(content);
            var fills = xml.Descendants().Where(node => node.Name.LocalName == "FillColor").ToArray();
            Assert.Equal(new[] { "80 80 80", "255 0 0" }, fills.Select(node => node.Attribute("Value")?.Value));
            Assert.All(fills, node => { Assert.Null(node.Attribute("ColorSpace")); Assert.Null(node.Attribute("Index")); });
        }
        var reread = await new OfdReader().ReadAsync(new MemoryStream(rewritten.ToArray()));
        Assert.Equal(2, reread.Pages[0].Elements.OfType<OfdPathElement>().Count());
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(rewritten.ToArray()), svg);
        var fillsAfter = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants()
            .Where(node => node.Name.LocalName == "path").Select(node => node.Attribute("fill")?.Value).ToArray();
        Assert.Equal(new[] { "rgb(80,80,80)", "rgb(255,0,0)" }, fillsAfter);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(rewritten.ToArray()), pdf);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(rewritten.ToArray()), html);
        Assert.Contains("rgb(80,80,80)", Encoding.UTF8.GetString(html.ToArray()));
        Assert.Contains("rgb(255,0,0)", Encoding.UTF8.GetString(html.ToArray()));
        Retain("top-level-gray-cmyk", rewritten.ToArray(), pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task MissingSealImageFailsInsteadOfDroppingSiblingDrawings()
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>" +
            "<TextObject ID='502' Boundary='10 40 30 10' Font='10' Size='4'><TextCode X='0' Y='4'>sibling</TextCode></TextObject>" +
            "<ImageObject ID='503' Boundary='40 10 20 20' ResourceID='999'/>";
        var seal = await Source(body, font: true);
        var outer = await ImageIoTests.WithSeals([seal], pageCount: 1);
        using var pdf = new MemoryStream();
        var failure = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(outer), pdf));
        Assert.Contains("no resolved payload", failure.Message);
        Assert.Equal(0, pdf.Length);
    }

    [Fact]
    public async Task InvalidPageBlockPaintInSealFailsInsteadOfDroppingAppearance()
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='999' Value='0 0 255'/>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var seal = await Source(body);
        var outer = await ImageIoTests.WithSeals([seal], pageCount: 1);
        using var pdf = new MemoryStream();
        var failure = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(outer), pdf));
        Assert.Contains("ColorSpace", failure.Message);
        Assert.Equal(0, pdf.Length);
    }

    [Theory]
    [InlineData("<VendorDrawing ID='501'/>")]
    [InlineData("<PathObject ID='501' Boundary='10 10 20 20' Fill='true'><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData><VendorChild/></PathObject>")]
    public async Task UnsupportedPageBlockDrawingInSealFailsInsteadOfDroppingAppearance(string body)
    {
        var seal = await Source(body);
        var outer = await ImageIoTests.WithSeals([seal], pageCount: 1);
        using var pdf = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(outer), pdf));
        Assert.Equal(0, pdf.Length);
    }

    [Fact]
    public async Task StrokedScaledTextUsesInheritedPaintAndKeepsLogicalText()
    {
        var resources = "<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette><CV>255 0 0</CV></Palette></ColorSpace></ColorSpaces>" +
            "<DrawParams><DrawParam ID='30' LineWidth='0.25'><StrokeColor ColorSpace='20' Index='0'/></DrawParam>" +
            "<DrawParam ID='31' Relative='30'/></DrawParams>";
        var body = "<TextObject ID='501' Boundary='10 10 60 20' Font='10' Size='6' HScale='0.5' Stroke='true' Fill='false' DrawParam='31'>" +
            "<TextCode X='2' Y='8'>R01 text</TextCode></TextObject>";
        var bytes = await Source(body, resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        Assert.Equal("R01 text", new OfdTextExtractor().Extract(read));
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        var rendered = Assert.Single(XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants().Where(n => n.Name.LocalName == "text"));
        Assert.Equal("none", rendered.Attribute("fill")?.Value);
        Assert.Equal("rgb(255,0,0)", rendered.Attribute("stroke")?.Value);
        Assert.Equal("0.25", rendered.Attribute("stroke-width")?.Value);
        Assert.Contains("0.5", rendered.Attribute("transform")?.Value);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("R01 text", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Contains("R01 text", parsed.GetPage(1).Text);
        Retain("stroke-hscale", bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task OutlineOnlyTextWithoutStrokeColorUsesOpaqueBlack()
    {
        var body = "<TextObject ID='501' Boundary='10 10 60 20' Font='10' Size='6' Fill='false' Stroke='true'>" +
            "<TextCode X='2' Y='8'>Outline</TextCode></TextObject>";
        var bytes = await Source(body, font: true);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        var text = Assert.Single(XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants(),
            node => node.Name.LocalName == "text");
        Assert.Equal("none", text.Attribute("fill")?.Value);
        Assert.Equal("rgb(0,0,0)", text.Attribute("stroke")?.Value);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("rgb(0,0,0)", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = reader.GetPageReader(0);
        var pixels = page.GetImage();
        Assert.Contains(Enumerable.Range(0, pixels.Length / 4), index =>
            pixels[index * 4] < 180 && pixels[index * 4 + 1] < 180 && pixels[index * 4 + 2] < 180);
        Retain("implicit-stroke", bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NestedPageBlockKeepsOrderTextAndOriginalRoot(bool template)
    {
        var inner = $"<PageBlock ID='501'><PathObject ID='502' Boundary='10 10 30 30' Fill='true' Stroke='false'><FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 30 0 L 30 30 L 0 30 C</AbbreviatedData></PathObject>" +
            "<TextObject ID='503' Boundary='10 40 50 10' Font='10' Size='4'><TextCode X='0' Y='4'>nested text</TextCode></TextObject></PageBlock>";
        var bytes = await Source(inner, template, font: true);
        using var source = new MemoryStream(bytes);
        var read = await new OfdReader().ReadAsync(source);
        var root = Assert.IsType<OfdRawElement>(Assert.Single(template ? read.Pages[0].Templates[0].Elements : read.Pages[0].Elements));
        Assert.Equal("PageBlock", root.LocalName);
        Assert.Contains("ID=\"503\"", root.Xml);
        Assert.Equal("nested text", new OfdTextExtractor().Extract(read, includeTemplates: true));
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        var shapes = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants()
            .Where(n => n.Name.LocalName is "path" or "text").Select(n => n.Name.LocalName).ToArray();
        Assert.Equal(new[] { "path", "text" }, shapes);
        using var html = new MemoryStream();
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("nested text", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("nested text", parsed.GetPage(1).Text);
        Retain(template ? "template-block" : "ordinary-block", bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
        using var roundTrip = new MemoryStream(); await new OfdPackageWriter().WriteAsync(read, roundTrip);
        roundTrip.Position = 0; var again = await new OfdReader().ReadAsync(roundTrip);
        var preserved = Assert.IsType<OfdRawElement>(Assert.Single(template ? again.Pages[0].Templates[0].Elements : again.Pages[0].Elements));
        Assert.Contains("ID=\"502\"", preserved.Xml);
        Assert.Contains("ID=\"503\"", preserved.Xml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SplitClonesPageBlockProjectionWithoutSharingChildren(bool template)
    {
        var body = "<PageBlock ID='501'><TextObject ID='502' Boundary='10 10 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>split nested</TextCode></TextObject></PageBlock>";
        var source = await new OfdReader().ReadAsync(new MemoryStream(await Source(body, template, font: true)));
        var original = Assert.IsType<OfdRawElement>(Assert.Single(template ? source.Pages[0].Templates[0].Elements : source.Pages[0].Elements));
        var split = OfdDocumentSplitter.Split(source, [0]);
        var cloned = Assert.IsType<OfdRawElement>(Assert.Single(template ? split.Pages[0].Templates[0].Elements : split.Pages[0].Elements));
        Assert.NotSame(original, cloned);
        Assert.NotSame(original.PageBlockChildren, cloned.PageBlockChildren);
        Assert.Equal("split nested", new OfdTextExtractor().Extract(split, includeTemplates: true));
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(split, output);
        var reread = await new OfdReader().ReadAsync(new MemoryStream(output.ToArray()));
        Assert.Equal("split nested", new OfdTextExtractor().Extract(reread, includeTemplates: true));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(output.ToArray()), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(output.ToArray()), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(output.ToArray()), html);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("split nested", parsed.GetPage(1).Text);
        Assert.Contains("split nested", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("split nested", Encoding.UTF8.GetString(html.ToArray()));
        Retain(template ? "split-template-block" : "split-ordinary-block", output.ToArray(), pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("")]
    [InlineData(" Visible='true'")]
    public async Task UnknownNestedDrawingFailsAllRenderers(string visibility)
    {
        var bytes = await Source($"<VendorDrawing ID='501'{visibility}/>");
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task UnsupportedTypedPageBlockChildFailsAllRenderers(bool text, bool template)
    {
        var body = text
            ? "<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4'><TextCode X='0' Y='4'>text</TextCode><VendorChild/></TextObject>"
            : "<PathObject ID='501' Boundary='10 10 20 20' Fill='true'><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData><VendorChild/></PathObject>";
        var bytes = await Source(body, template, font: text);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageBlockTextCodePreservesXmlSpace(bool template)
    {
        var bytes = await Source("<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4' xml:space='preserve'>two  spaces</TextCode></TextObject>", template, font: true);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        Assert.True(pdf.Length > 0);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        Assert.Contains("two  spaces", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream();
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("two  spaces", Encoding.UTF8.GetString(html.ToArray()));
    }

    [Theory]
    [InlineData("Cap='Round'", false)]
    [InlineData("Join='Bevel'", false)]
    [InlineData("MiterLimit='3'", false)]
    [InlineData("DashPattern='1 2'", false)]
    [InlineData("DashOffset='1'", false)]
    [InlineData("DashPattern='1 2'", true)]
    public async Task UnsupportedTextStrokeParametersFailAllRenderers(string parameter, bool inherited)
    {
        var resources = inherited ? $"<DrawParams><DrawParam ID='30' {parameter}/></DrawParams>" : null;
        var attribute = inherited ? "DrawParam='30'" : parameter;
        var body = $"<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4' Stroke='true' Fill='false' {attribute}>" +
            "<TextCode X='0' Y='4'>outline</TextCode></TextObject>";
        var bytes = await Source(body, resources: resources, font: true);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData("TextObject", false)]
    [InlineData("TextObject", true)]
    [InlineData("PathObject", false)]
    [InlineData("PathObject", true)]
    [InlineData("ImageObject", false)]
    [InlineData("ImageObject", true)]
    public async Task ForeignNamespaceObjectsRemainOpaque(string localName, bool template)
    {
        var children = localName switch
        {
            "TextObject" => "<ext:TextCode X='0' Y='4'>foreign</ext:TextCode>",
            "PathObject" => "<ext:AbbreviatedData>M 0 0 L 20 0 L 20 20 C</ext:AbbreviatedData>",
            _ => string.Empty
        };
        var body = $"<ext:{localName} xmlns:ext='urn:vendor' ID='501' Boundary='10 10 20 20' Font='10' Size='4' ResourceID='10'>" +
            children + $"</ext:{localName}>";
        var bytes = await Source(body, template, font: true);
        var source = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(template ? source.Pages[0].Templates[0].Elements : source.Pages[0].Elements));
        var child = Assert.IsType<OfdRawElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Equal(localName, child.LocalName);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task HiddenUnknownNestedDrawingDoesNotBlockVisibleContent(bool template)
    {
        var body = "<VendorDrawing ID='501' Visible='false'/>" +
            "<TextObject ID='502' Boundary='10 10 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>visible sibling</TextCode></TextObject>";
        var bytes = await Source(body, template, font: true);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("visible sibling", parsed.GetPage(1).Text);
        Assert.Contains("visible sibling", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("visible sibling", Encoding.UTF8.GetString(html.ToArray()));
        Retain(template ? "hidden-opaque-template" : "hidden-opaque-ordinary", bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task MissingNestedImagePayloadFailsAllRenderers()
    {
        var bytes = await Source("<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='999'/>");
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Equal(0, html.Length);
    }

    [Fact]
    public void PaintNormalizationPreservesMaterializationLimit()
    {
        var paint = new OfdPaintResources();
        var material = new OfdModelMaterialization(new OfdModelLimits { MaxMaterializedBytes = 1024 });
        var source = XElement.Parse($"<PathObject xmlns='{Ns}' ID='501'/>");
        var error = Assert.Throws<InvalidDataException>(() => paint.Normalize(source, 100, material, default));
        Assert.True(OfdTraversalBudget.IsLimit(error));
    }

    [Theory]
    [InlineData("<FillColor ColorSpace='999' Value='0 0 255'/>", "")]
    [InlineData("<FillColor><AxialShd/></FillColor>", "")]
    [InlineData("<FillColor Index='1' ColorSpace='20'/>", "<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette><CV>0 0 0</CV></Palette></ColorSpace></ColorSpaces>")]
    [InlineData("<FillColor Value='0 0 255' Alpha='bogus'/>", "")]
    public async Task UnsupportedPaintFailsBeforeHtmlOutput(string color, string resources)
    {
        var body = $"<PathObject ID='501' Boundary='10 10 20 20' Fill='true'>{color}<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var bytes = await Source(body, resources: resources.Length == 0 ? null : resources);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NestedImageUsesOriginalMediaResourceAndPaintOrder(bool substitution)
    {
        using var pixels = new Image<Rgba32>(2, 2, Color.Red);
        using var png = new MemoryStream(); pixels.SaveAsPng(png);
        var package = new OfdDocumentPackage();
        var page = new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        page.Elements.Add(new OfdRawElement { LocalName = "PageBlock", Xml =
            $"<PageBlock xmlns='{Ns}' ID='500'><PathObject ID='501' Boundary='10 10 10 10' Fill='true' Stroke='false'>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></PathObject></PageBlock>" });
        page.Elements.Add(new OfdImageElement { XMillimeters = 30, YMillimeters = 10, WidthMillimeters = 20, HeightMillimeters = 20,
            Data = png.ToArray() });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output);
        using var patched = new MemoryStream();
        using (var original = new ZipArchive(new MemoryStream(output.ToArray()), ZipArchiveMode.Read))
        using (var destination = new ZipArchive(patched, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in original.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var layer = xml.Descendants().First(n => n.Name.LocalName == "Layer");
                    var block = layer.Elements().First(n => n.Name.LocalName == "PageBlock");
                    var image = layer.Elements().First(n => n.Name.LocalName == "ImageObject");
                    image.Remove();
                    if (substitution) image.SetAttributeValue("Substitution", "999");
                    block.Add(image);
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = destination.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var input = patched.ToArray();
        if (substitution)
        {
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), new MemoryStream()));
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), new MemoryStream()));
            using var rejectedHtml = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(input), rejectedHtml));
            Assert.Equal(0, rejectedHtml.Length);
            return;
        }
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(input), svg);
        var names = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants()
            .Where(n => n.Name.LocalName is "path" or "image").Select(n => n.Name.LocalName).ToArray();
        Assert.Equal(new[] { "path", "image" }, names);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(input), pdf);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(input), html);
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(html.ToArray()));
        var read = await new OfdReader().ReadAsync(new MemoryStream(input));
        using var roundTrip = new MemoryStream(); await new OfdPackageWriter().WriteAsync(read, roundTrip);
        using var svgAgain = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(roundTrip.ToArray()), svgAgain);
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(svgAgain.ToArray()));
        Retain("image-block", input, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }
}
