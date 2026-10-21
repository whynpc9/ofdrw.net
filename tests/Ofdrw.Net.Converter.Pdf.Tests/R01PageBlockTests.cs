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

    private static async Task<byte[]> Source(string body, bool template = false, string? resources = null, bool font = false, bool image = false,
        string? resourceNamespace = null)
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
                var resNs = resourceNamespace is null ? ns : XNamespace.Get(resourceNamespace);
                var resourceXml = new XDocument(new XElement(resNs + "Res", wrapper.Elements()));
                using var resourceStream = outputZip.CreateEntry("Doc_0/DocumentRes.xml").Open();
                await resourceStream.WriteAsync(Encoding.UTF8.GetBytes(resourceXml.ToString(SaveOptions.DisableFormatting)));
            }
            if (image)
            {
                using var pixel = new Image<Rgba32>(1, 1, Color.Red);
                using var imageData = new MemoryStream(); await pixel.SaveAsPngAsync(imageData);
                using var imageStream = outputZip.CreateEntry("Doc_0/image.png").Open();
                await imageStream.WriteAsync(imageData.ToArray());
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
    [InlineData("between", "<CV>255 0 0</CV><!--between--><CV>0 0 255</CV>")]
    [InlineData("inside", "<CV>255 0 0</CV><CV>0 <!--inside--> 0 255</CV>")]
    public async Task IndexedPaletteCommentsDoNotChangePaint(string location, string entries)
    {
        var resources = $"<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette>{entries}</Palette></ColorSpace></ColorSpaces>";
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='20' Index='1'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("rgb(0,0,255)", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("rgb(0,0,255)", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 0);
        Retain("palette-comment-" + location, ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PaletteValueOverridesConflictingIndexInAllOutputs(bool drawParam)
    {
        // GB/T 33190-2016, 8.3.2 table 26: Index is used only if Value is absent.
        const string palette = "<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette>" +
            "<CV>255 0 0</CV><CV>0 0 255</CV></Palette></ColorSpace></ColorSpaces>";
        var resources = palette + (drawParam
            ? "<DrawParams><DrawParam ID='30'><FillColor ColorSpace='20' Index='1' Value='255 0 0'/></DrawParam></DrawParams>"
            : "");
        var style = drawParam ? "DrawParam='30'" : "";
        var color = drawParam ? "" : "<FillColor ColorSpace='20' Index='1' Value='255 0 0'/>";
        var body = $"<PathObject ID='501' Boundary='10 10 30 30' {style} Fill='true' Stroke='false'>" +
            color + "<AbbreviatedData>M 0 0 L 30 0 L 30 30 L 0 30 C</AbbreviatedData></PathObject>";
        var bytes = await Source(body, resources: resources);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        var path = Assert.Single(XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants(),
            node => node.Name.LocalName == "path");
        Assert.Equal("rgb(255,0,0)", path.Attribute("fill")?.Value);

        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("rgb(255,0,0)", Encoding.UTF8.GetString(html.ToArray()));

        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        using var document = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = document.GetPageReader(0);
        var pixels = page.GetImage();
        var offset = ((int)(20 * 72 / 25.4 * 2) * page.GetPageWidth() + (int)(20 * 72 / 25.4 * 2)) * 4;
        Assert.True(pixels[offset] < 30 && pixels[offset + 1] < 30 && pixels[offset + 2] > 230,
            "PDF path interior should be red from Value when Index is also present.");
        Retain(drawParam ? "palette-value-priority-drawparam" : "palette-value-priority-direct",
            bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ForeignResourceNamespaceDoesNotResolvePaint(bool drawParam, bool foreignRoot)
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
                    if (foreignRoot)
                    {
                        foreach (var node in xml.Descendants())
                        {
                            foreach (var declaration in node.Attributes().Where(a => a.IsNamespaceDeclaration).ToArray())
                                declaration.Remove();
                            node.Name = extension + node.Name.LocalName;
                        }
                    }
                    else xml.Root!.Elements().Single().Name = extension + (drawParam ? "DrawParams" : "ColorSpaces");
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
    public async Task DuplicatePalettesInReferencedColorSpaceFailAllRenderers()
    {
        var resources = "<ColorSpaces><ColorSpace ID='20' Type='RGB'>" +
            "<Palette><CV>255 0 0</CV></Palette><Palette><CV>0 0 255</CV></Palette>" +
            "</ColorSpace></ColorSpaces>";
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true'>" +
            "<FillColor ColorSpace='20' Index='0'/>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData("TextObject", "FillColor")]
    [InlineData("TextObject", "StrokeColor")]
    [InlineData("PathObject", "FillColor")]
    [InlineData("PathObject", "StrokeColor")]
    public async Task DuplicateDrawingColorsFailAllRenderers(string objectName, string colorName)
    {
        var colors = $"<{colorName} Value='255 0 0'/><{colorName} Value='0 0 255'/>";
        var body = objectName == "TextObject"
            ? $"<TextObject ID='501' Boundary='10 10 40 10' Font='10' Size='4'>{colors}<TextCode X='0' Y='4'>color</TextCode></TextObject>"
            : $"<PathObject ID='501' Boundary='10 10 20 20' Fill='true'>{colors}<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, font: objectName == "TextObject");
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html));
        Assert.Equal(0, html.Length);
    }

    [Fact]
    public async Task DuplicateDrawParamColorsFailAllRenderers()
    {
        const string resources = "<DrawParams><DrawParam ID='20'><FillColor Value='255 0 0'/>" +
            "<FillColor Value='0 0 255'/></DrawParam></DrawParams>";
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' DrawParam='20' Fill='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html));
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData("<?vendor key?>")]
    [InlineData("vendor extension")]
    public async Task ReferencedDrawParamRejectsIgnoredContent(string content)
    {
        var resources = $"<DrawParams><DrawParam ID='20'>{content}<FillColor Value='255 0 0'/></DrawParam></DrawParams>";
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' DrawParam='20' Fill='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("DrawParam has unsupported content", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData("<FillColor Value='0 0 0' xmlns:v='urn:vendor' v:Mode='x'/>")]
    [InlineData("<FillColor Value='0 0 0'><?vendor key?></FillColor>")]
    [InlineData("<FillColor ColorSpace='999' Value='0 0 0'/>")]
    public async Task OverriddenAncestorDrawParamColorStillNeedsValidPaint(string ancestorColor)
    {
        var resources = $"<DrawParams><DrawParam ID='20'>{ancestorColor}</DrawParam>" +
            "<DrawParam ID='21' Relative='20'><FillColor Value='255 0 0'/></DrawParam></DrawParams>";
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' DrawParam='21' Fill='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task ReferencedDrawParamAllowsCommentsAndWhitespace()
    {
        const string resources = "<DrawParams><DrawParam ID='20'> \n<!-- note -->" +
            "<FillColor Value='255 0 0'/></DrawParam></DrawParams>";
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' DrawParam='20' Fill='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        var path = Assert.Single(XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants(),
            node => node.Name.LocalName == "path");
        Assert.Equal("rgb(255,0,0)", path.Attribute("fill")?.Value);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task ReferencedPaintNodesRejectIgnoredExtensions(bool colorSpace)
    {
        var resources = colorSpace
            ? "<ColorSpaces><ColorSpace ID='20' Type='RGB'><?vendor key?></ColorSpace></ColorSpaces>"
            : "<DrawParams><DrawParam ID='20'><FillColor Value='255 0 0'>vendor extension</FillColor></DrawParam></DrawParams>";
        var style = colorSpace ? "" : "DrawParam='20'";
        var color = colorSpace ? "<FillColor ColorSpace='20' Value='255 0 0'/>" : "";
        var body = $"<PathObject ID='501' Boundary='10 10 20 20' {style} Fill='true'>" +
            color + "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Equal(0, output.Length);
        }
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
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnresolvedPageBlockTextPaintPreservesExtractionButFailsAllRenderers(bool template)
    {
        var body = "<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4'>" +
            "<FillColor ColorSpace='999' Value='0 0 255'/>" +
            "<TextCode X='0' Y='4'>recoverable text</TextCode></TextObject>";
        var bytes = await Source(body, template: template, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        Assert.Equal("recoverable text", new OfdTextExtractor().Extract(read, includeTemplates: true));

        using var pdf = new MemoryStream();
        var pdfFailure = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf));
        Assert.Contains("ColorSpace", pdfFailure.Message);
        Assert.Equal(0, pdf.Length);

        using var svg = new MemoryStream();
        var svgFailure = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg));
        Assert.Contains("ColorSpace", svgFailure.Message);
        Assert.Equal(0, svg.Length);

        using var html = new MemoryStream();
        var htmlFailure = await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html));
        Assert.Contains("ColorSpace", htmlFailure.Message);
        Assert.Equal(0, html.Length);
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData(false, "999")]
    [InlineData(true, null)]
    [InlineData(true, "999")]
    public async Task UnresolvedPageBlockFontPreservesTextButFailsAllRenderers(bool template, string? fontId)
    {
        var fontAttribute = fontId is null ? "" : $" Font='{fontId}'";
        var bytes = await Source($"<TextObject ID='501' Boundary='10 10 50 10'{fontAttribute} Size='4'>" +
            "<TextCode X='0' Y='4'>recoverable text</TextCode></TextObject>", template: template, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        Assert.Equal("recoverable text", new OfdTextExtractor().Extract(read, includeTemplates: true));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(bytes), output));
            Assert.Contains("Font", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData("Bold")]
    [InlineData("Italic")]
    public async Task MalformedFontFlagPreservesPageBlockTextButFailsAllRenderers(string flag)
    {
        var resources = $"<Fonts><Font ID='20' FontName='Noto Sans' {flag}='bogus'/></Fonts>";
        var ofd = await Source("<TextObject ID='501' Boundary='10 10 50 10' Font='20' Size='4'>" +
            "<TextCode X='0' Y='4'>recoverable text</TextCode></TextObject>", resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains(read.Fonts, font => font.Id == "20");
        Assert.Equal("recoverable text", new OfdTextExtractor().Extract(read));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("Font", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData("root")]
    [InlineData("container")]
    [InlineData("font")]
    public async Task ForeignNamespaceFontCannotResolvePageBlockText(string foreignPart)
    {
        const string text = "<TextObject ID='501' Boundary='10 10 50 10' Font='20' Size='4'>" +
            "<TextCode X='0' Y='4'>recoverable text</TextCode></TextObject>";
        const string vendor = "urn:vendor";
        var fonts = foreignPart switch
        {
            "container" => "<Fonts xmlns='urn:vendor'><Font ID='20' FontName='Vendor'/></Fonts>",
            "font" => "<Fonts><Font xmlns='urn:vendor' ID='20' FontName='Vendor'/></Fonts>",
            _ => "<Fonts><Font ID='20' FontName='Vendor'/></Fonts>"
        };
        var bytes = await Source(text, resources: fonts, resourceNamespace: foreignPart == "root" ? vendor : null);
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        Assert.Equal("recoverable text", new OfdTextExtractor().Extract(read));
        Assert.DoesNotContain(read.Fonts, f => f.Id == "20");
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(bytes), output));
            Assert.Contains("Font", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(Ns)]
    [InlineData("http://www.ofdspec.org/2016")]
    public async Task AdmittedResourceFontResolvesPageBlockText(string resourceNamespace)
    {
        var fonts = $"<Fonts xmlns='{resourceNamespace}'><Font ID='20' FontName='Noto Sans'/></Fonts>";
        var bytes = await Source("<TextObject ID='501' Boundary='10 10 50 10' Font='20' Size='4'>" +
            "<TextCode X='0' Y='4'>known font</TextCode></TextObject>", resources: fonts,
            resourceNamespace: resourceNamespace);
        var read = await new OfdReader().ReadAsync(new MemoryStream(bytes));
        Assert.Contains(read.Fonts, f => f.Id == "20" && f.FontName == "Noto Sans");
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        Assert.True(pdf.Length > 0);
        using var svg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        using var html = new MemoryStream();
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.Contains("known font", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("known font", Encoding.UTF8.GetString(html.ToArray()));
        Retain(resourceNamespace.Length == 0 ? "font-resource-unqualified" :
            resourceNamespace == Ns ? "font-resource-2009" : "font-resource-2016",
            bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("root", false)]
    [InlineData("container", false)]
    [InlineData("media", false)]
    [InlineData("root", true)]
    [InlineData("container", true)]
    [InlineData("media", true)]
    public async Task ForeignNamespaceMediaCannotResolvePageBlockImage(string foreignPart, bool template)
    {
        var media = foreignPart switch
        {
            "container" => "<MultiMedias xmlns='urn:vendor'><MultiMedia ID='20' Type='Image' Format='PNG'><MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>",
            "media" => "<MultiMedias><MultiMedia xmlns='urn:vendor' ID='20' Type='Image' Format='PNG'><MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>",
            _ => "<MultiMedias><MultiMedia ID='20' Type='Image' Format='PNG'><MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>"
        };
        var bytes = await Source("<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='20'/>",
            template: template, resources: media, image: true, resourceNamespace: foreignPart == "root" ? "urn:vendor" : null);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(bytes), output));
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(Ns)]
    [InlineData("http://www.ofdspec.org/2016")]
    public async Task AdmittedResourceMediaResolvesPageBlockImage(string resourceNamespace)
    {
        var media = $"<MultiMedias xmlns='{resourceNamespace}'><MultiMedia ID='20' Type='Image' Format='PNG'><MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>";
        var bytes = await Source("<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='20'/>",
            resources: media, image: true, resourceNamespace: resourceNamespace);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(bytes), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(html.ToArray()));
        Retain(resourceNamespace.Length == 0 ? "media-resource-unqualified" :
            resourceNamespace == Ns ? "media-resource-2009" : "media-resource-2016",
            bytes, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("ImageObject", null)]
    [InlineData("ImageObject", "10 10 0 20")]
    [InlineData("ImageObject", "10 10 -1 20")]
    [InlineData("PathObject", "10 10 20 -1")]
    [InlineData("TextObject", null)]
    public async Task PageBlockDrawingRequiresPositiveBoundary(string kind, string? boundary)
    {
        var boundaryAttribute = boundary is null ? "" : $" Boundary='{boundary}'";
        var content = kind switch
        {
            "PathObject" => "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData>",
            "TextObject" => "<TextCode X='0' Y='4'>recoverable text</TextCode>",
            _ => ""
        };
        var attribute = kind switch
        {
            "ImageObject" => " ResourceID='99'",
            "TextObject" => " Font='10' Size='4'",
            _ => " Fill='true'"
        };
        var body = $"<{kind} ID='501'{boundaryAttribute}{attribute}>{content}</{kind}>";
        var resources = kind == "ImageObject"
            ? "<MultiMedias><MultiMedia ID='99' Type='Image' Format='PNG'><MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>"
            : null;
        var bytes = await Source(body, resources: resources, font: kind == "TextObject", image: kind == "ImageObject");
        if (kind == "TextObject")
            Assert.Equal("recoverable text", new OfdTextExtractor().Extract(await new OfdReader().ReadAsync(new MemoryStream(bytes))));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(bytes), output));
            Assert.Contains("Malformed PageBlock drawing boundary", error.Message);
            Assert.Equal(0, output.Length);
        }
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

    [Theory]
    [InlineData("", false)]
    [InlineData("", true)]
    [InlineData("http://www.ofdspec.org/2016", false)]
    [InlineData("http://www.ofdspec.org/2016", true)]
    public async Task DrawParamColorsUseDrawingNamespaceAndKeepLocalOverrides(string resourceNamespace, bool layer)
    {
        var resources = "<DrawParams><DrawParam ID='30'><FillColor Value='0 255 0'/>" +
            "<StrokeColor Value='255 0 0'/></DrawParam><DrawParam ID='31' Relative='30'/></DrawParams>";
        var body = "<PathObject ID='501' Boundary='5 5 20 20' DrawParam='31' Fill='true' Stroke='false'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>" +
            "<TextObject ID='502' Boundary='5 30 80 15' Font='10' Size='6' DrawParam='31' Fill='false' Stroke='true'>" +
            "<TextCode X='2' Y='8'>inherited red</TextCode></TextObject>" +
            "<TextObject ID='503' Boundary='5 50 80 15' Font='10' Size='6' DrawParam='31' Fill='false' Stroke='true'>" +
            "<StrokeColor Value='0 0 255'/><TextCode X='2' Y='8'>local blue</TextCode></TextObject>";
        var original = await Source(body, resources: resources, font: true);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var stream = entry.Open(); using var buffer = new MemoryStream(); await stream.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName == "Doc_0/DocumentRes.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    XNamespace ns = resourceNamespace;
                    foreach (var node in xml.Descendants())
                    {
                        foreach (var declaration in node.Attributes().Where(a => a.IsNamespaceDeclaration).ToArray())
                            declaration.Remove();
                        node.Name = ns + node.Name.LocalName;
                    }
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                if (layer && entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var block = xml.Descendants().Single(node => node.Name.LocalName == "PageBlock");
                    block.ReplaceWith(block.Elements().ToArray());
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var ofd = rewritten.ToArray();
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        var drawing = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var path = Assert.Single(drawing.Descendants().Where(node => node.Name.LocalName == "path"));
        Assert.Equal("rgb(0,255,0)", path.Attribute("fill")?.Value);
        var texts = drawing.Descendants().Where(node => node.Name.LocalName == "text").ToArray();
        Assert.Equal(2, texts.Length);
        Assert.Equal("rgb(255,0,0)", texts[0].Attribute("stroke")?.Value);
        Assert.Equal("rgb(0,0,255)", texts[1].Attribute("stroke")?.Value);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("rgb(255,0,0)", Encoding.UTF8.GetString(html.ToArray()));
        Assert.Contains("rgb(0,0,255)", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Contains("inherited red", parsed.GetPage(1).Text);
        Assert.Contains("local blue", parsed.GetPage(1).Text);
        Retain($"drawparam-ns-{(resourceNamespace.Length == 0 ? "unqualified" : "2016")}-{(layer ? "layer" : "block")}",
            ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(Ns, "http://www.ofdspec.org/2016", false)]
    [InlineData(Ns, "http://www.ofdspec.org/2016", true)]
    [InlineData("http://www.ofdspec.org/2016", Ns, false)]
    [InlineData("http://www.ofdspec.org/2016", Ns, true)]
    [InlineData("", Ns, false)]
    [InlineData("", Ns, true)]
    [InlineData(Ns, "", false)]
    [InlineData(Ns, "", true)]
    public async Task RelativeDrawParamChildColorsOverrideAcrossResourceNamespaces(string documentNamespace,
        string pageNamespace, bool layer)
    {
        const string resources = "<DrawParams><DrawParam ID='30'><FillColor Value='0 255 0'/>" +
            "<StrokeColor Value='255 0 0'/></DrawParam></DrawParams>";
        var body = "<PathObject ID='501' Boundary='5 5 20 20' DrawParam='32' Fill='true' Stroke='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>" +
            "<TextObject ID='502' Boundary='5 30 80 15' Font='10' Size='6' DrawParam='31' Fill='false' Stroke='true'>" +
            "<TextCode X='2' Y='8'>child blue</TextCode></TextObject>" +
            "<TextObject ID='503' Boundary='5 50 80 15' Font='10' Size='6' DrawParam='31' Fill='false' Stroke='true'>" +
            "<StrokeColor Value='255 0 255'/><TextCode X='2' Y='8'>local magenta</TextCode></TextObject>";
        var original = await Source(body, resources: resources, font: true);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(original), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var stream = entry.Open(); using var buffer = new MemoryStream(); await stream.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName == "Doc_0/DocumentRes.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    XNamespace ns = documentNamespace;
                    foreach (var node in xml.Descendants())
                    {
                        foreach (var declaration in node.Attributes().Where(a => a.IsNamespaceDeclaration).ToArray())
                            declaration.Remove();
                        node.Name = ns + node.Name.LocalName;
                    }
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                if (layer && entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var block = xml.Descendants().Single(node => node.Name.LocalName == "PageBlock");
                    block.ReplaceWith(block.Elements().ToArray());
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
            XNamespace pageNs = pageNamespace;
            var pageRes = new XElement(pageNs + "Res", new XElement(pageNs + "DrawParams",
                new XElement(pageNs + "DrawParam", new XAttribute("ID", "31"), new XAttribute("Relative", "30"),
                    new XElement(pageNs + "StrokeColor", new XAttribute("Value", "0 0 255"))),
                new XElement(pageNs + "DrawParam", new XAttribute("ID", "32"), new XAttribute("Relative", "31"),
                    new XElement(pageNs + "FillColor", new XAttribute("Value", "255 255 0")))));
            using var pageStream = output.CreateEntry("Doc_0/Pages/Page_0/PageRes.xml").Open();
            await pageStream.WriteAsync(Encoding.UTF8.GetBytes(pageRes.ToString(SaveOptions.DisableFormatting)));
        }
        var ofd = rewritten.ToArray();
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        var drawing = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()));
        var path = Assert.Single(drawing.Descendants().Where(node => node.Name.LocalName == "path"));
        Assert.Equal("rgb(255,255,0)", path.Attribute("fill")?.Value);
        Assert.Equal("rgb(0,0,255)", path.Attribute("stroke")?.Value);
        var texts = drawing.Descendants().Where(node => node.Name.LocalName == "text").ToArray();
        Assert.Equal(2, texts.Length);
        Assert.Equal("rgb(0,0,255)", texts[0].Attribute("stroke")?.Value);
        Assert.Equal("rgb(255,0,255)", texts[1].Attribute("stroke")?.Value);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("rgb(255,255,0)", Encoding.UTF8.GetString(html.ToArray()));
        Assert.Contains("rgb(255,0,255)", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Contains("child blue", parsed.GetPage(1).Text);
        Assert.Contains("local magenta", parsed.GetPage(1).Text);
        var pair = (documentNamespace.Length == 0 ? "none" : documentNamespace == Ns ? "2009" : "2016") + "-" +
            (pageNamespace.Length == 0 ? "none" : pageNamespace == Ns ? "2009" : "2016");
        Retain($"drawparam-relative-{pair}-{(layer ? "layer" : "block")}", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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
    [InlineData("PathObject", "Fill='bogus'", false)]
    [InlineData("PathObject", "CTM='bogus'", false)]
    [InlineData("PathObject", "LineWidth='bogus'", false)]
    [InlineData("PathObject", "Stroke='2'", false)]
    [InlineData("PathObject", "Visible='bogus'", false)]
    [InlineData("PathObject", "Boundary='10 10 20 20 5'", false)]
    [InlineData("ImageObject", "Alpha='bogus'", false)]
    [InlineData("ImageObject", "Alpha='300'", false)]
    [InlineData("TextObject", "Size='bogus'", true)]
    [InlineData("TextObject", "Italic='bogus'", true)]
    [InlineData("TextObject", "Weight='bogus'", true)]
    [InlineData("TextObject", "Weight='0'", true)]
    [InlineData("TextObject", "Weight='1000'", true)]
    [InlineData("TextObject", "Weight='150'", true)]
    public async Task MalformedKnownPageBlockAttributesFailBeforeReaderDefaults(
        string kind, string malformedAttribute, bool font)
    {
        var content = kind switch
        {
            "PathObject" => "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData>",
            "TextObject" => "<TextCode X='0' Y='4'>original text</TextCode>",
            _ => ""
        };
        var extra = kind == "TextObject" ? "Font='10'" : kind == "ImageObject" ? "ResourceID='99'" : "Fill='true'";
        // A malformed Fill replaces the valid default instead of duplicating it.
        if (malformedAttribute.StartsWith("Fill=", StringComparison.Ordinal)) extra = "";
        var boundary = malformedAttribute.StartsWith("Boundary=", StringComparison.Ordinal)
            ? malformedAttribute : "Boundary='10 10 20 20'";
        var attribute = malformedAttribute.StartsWith("Boundary=", StringComparison.Ordinal) ? "" : malformedAttribute;
        var body = $"<{kind} ID='501' {boundary} {extra} {attribute}>{content}</{kind}>";
        var resources = kind == "ImageObject"
            ? "<MultiMedias><MultiMedia ID='99' Type='Image' Format='PNG'><MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>"
            : null;
        var bytes = await Source(body, resources: resources, font: font, image: kind == "ImageObject");
        if (kind == "TextObject")
        {
            var imported = await new OfdReader().ReadAsync(new MemoryStream(bytes));
            Assert.Equal("original text", new OfdTextExtractor().Extract(imported));
        }
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(bytes), output));
            Assert.Contains("Malformed PageBlock", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData("Alpha", "0255")]
    [InlineData("ReadDirection", "00")]
    [InlineData("CharDirection", " 0 ")]
    public async Task EquivalentTextIntegerLexemesRenderInAllOutputs(string name, string value)
    {
        var body = $"<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4' {name}='{value}'>" +
            "<TextCode X='0' Y='4'>numeric text</TextCode></TextObject>";
        var ofd = await Source(body, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Equal("numeric text", new OfdTextExtractor().Extract(read));
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("numeric text", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("numeric text", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("numeric text", parsed.GetPage(1).Text);
        Retain("text-integer-" + name.ToLowerInvariant(), ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("DeltaX", "1 bogus")]
    [InlineData("DeltaY", "1 bogus")]
    [InlineData("DeltaX", "1 g 2")]
    public async Task PageBlockTextAdvanceRejectsMalformedTrailingTokens(string axis, string expression)
    {
        var body = $"<TextObject ID='501' Boundary='10 10 20 10' Font='10' Size='4'>" +
            $"<TextCode X='0' Y='4' {axis}='{expression}'>AB</TextCode></TextObject>";
        var ofd = await Source(body, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Equal("AB", new OfdTextExtractor().Extract(read));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("Malformed PageBlock", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task PageBlockTextAdvanceAcceptsCompleteRepeatExpression()
    {
        var ofd = await Source("<TextObject ID='501' Boundary='10 10 20 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4' DeltaX='g 1 3'>AB</TextCode></TextObject>", font: true);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("AB", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("AB", Encoding.UTF8.GetString(html.ToArray()));
        Retain("advance-repeat", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(100)]
    [InlineData(900)]
    public async Task PageBlockTextAcceptsWeightEndpoints(int weight)
    {
        var ofd = await Source($"<TextObject ID='501' Boundary='10 10 20 10' Font='10' Size='4' Weight='{weight}'>" +
            "<TextCode X='0' Y='4'>AB</TextCode></TextObject>", font: true);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("AB", Encoding.UTF8.GetString(svg.ToArray()));
        Retain($"weight-{weight}", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("0 0", true)]
    [InlineData("1 2", false)]
    public async Task PageBlockClipStartMustPreserveRenderedGeometry(string start, bool supported)
    {
        var body = $"<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            $"<Clips><Clip><Area Start='{start}'><Path><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></Path></Area></Clip></Clips>" +
            "<FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var bytes = await Source(body);
        using var pdf = new MemoryStream();
        if (supported)
        {
            await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
            Assert.True(pdf.Length > 0);
        }
        else
        {
            var error = await Assert.ThrowsAsync<NotSupportedException>(() =>
                new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf));
            Assert.Contains("Malformed PageBlock", error.Message);
            Assert.Equal(0, pdf.Length);
        }
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
    public async Task PageDrawingCannotImpersonateAnnotationMetadata(bool pageBlock)
    {
        const string xml = "<ext:UnmodeledAnnotationMetadata xmlns:ext='urn:vendor' ID='501'/>";
        byte[] ofd;
        if (pageBlock) ofd = await Source(xml);
        else
        {
            var package = new OfdDocumentPackage();
            var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
            page.Elements.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = xml });
            package.Pages.Add(page);
            using var written = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, written);
            ofd = written.ToArray();
        }
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var root = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.Equal(pageBlock ? "PageBlock" : "UnmodeledAnnotationMetadata", root.LocalName);
        if (pageBlock) Assert.Equal("UnmodeledAnnotationMetadata", Assert.IsType<OfdRawElement>(Assert.Single(root.PageBlockChildren!)).LocalName);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), new MemoryStream()));
        using var html = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html));
        Assert.Equal(0, html.Length);
    }

    [Fact]
    public void UnmodeledAnnotationIndexEnvelopeRequiresReaderProvenance()
    {
        const string xml = "<UnmodeledAnnotations><VendorDrawing/></UnmodeledAnnotations>";
        var forged = new OfdRawElement { LocalName = "Annotations", Xml = xml };
        var page = new OfdPage();
        page.Elements.Add(forged);
        Assert.Throws<NotSupportedException>(() => OfdRenderPreflight.Validate(page.Elements));
        page.AnnotationAppearances.Add(forged);
        Assert.Throws<NotSupportedException>(() => OfdRenderPreflight.Validate(page.AnnotationAppearances));

        var readerOwned = new OfdRawElement { LocalName = "Annotations", Xml = xml, ReaderAnnotationMetadataXml = xml };
        OfdRenderPreflight.Validate(new[] { readerOwned });
        readerOwned.Xml = "<UnmodeledAnnotations><DifferentDrawing/></UnmodeledAnnotations>";
        Assert.Throws<NotSupportedException>(() => OfdRenderPreflight.Validate(new[] { readerOwned }));
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
