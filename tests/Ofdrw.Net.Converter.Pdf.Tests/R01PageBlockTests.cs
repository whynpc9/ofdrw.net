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

    private static async Task<byte[]> ReNamespacePageBlock(byte[] seed, string layerNamespace, string blockNamespace,
        string? compositeNamespace = null)
    {
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var layer = Assert.Single(xml.Descendants(), node => node.Name.LocalName == "Layer");
                    var block = Assert.Single(layer.Elements(), node => node.Name.LocalName == "PageBlock");
                    var children = block.Nodes().Select(node => node is XElement child &&
                        child.Name.LocalName == "CompositeObject" && compositeNamespace is not null
                        ? new XElement(XNamespace.Get(compositeNamespace) + "CompositeObject",
                            child.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), child.Nodes())
                        : node).ToArray();
                    var changedBlock = new XElement(XNamespace.Get(blockNamespace) + "PageBlock",
                        block.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), children);
                    layer.ReplaceWith(new XElement(XNamespace.Get(layerNamespace) + "Layer",
                        layer.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration), changedBlock));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        return rewritten.ToArray();
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

    [Fact]
    public async Task CompositeGraphicUsesDocumentPaintResources()
    {
        const string resources = "<ColorSpaces><ColorSpace ID='20' Type='RGB'><Palette>" +
            "<CV>255 0 0</CV><CV>0 0 255</CV></Palette></ColorSpace></ColorSpaces>" +
            "<DrawParams><DrawParam ID='30'><FillColor ColorSpace='20' Index='0'/></DrawParam></DrawParams>" +
            "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
            "<PathObject ID='552' Boundary='0 0 20 20' DrawParam='30' Fill='true' Stroke='false'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>" +
            "<PathObject ID='553' Boundary='30 0 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='20' Index='1'/>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        var ofd = await Source("<CompositeObject ID='554' Boundary='10 10 50 20' ResourceID='551'/>", resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var paths = Assert.IsType<OfdCompositeResource>(Assert.Single(read.CompositeResources!.Items))
            .Elements.OfType<OfdPathElement>().ToArray();
        Assert.Equal(2, paths.Length);
        Assert.All(paths, path => Assert.Null(path.RenderPaintError));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        foreach (var markup in new[] { Encoding.UTF8.GetString(svg.ToArray()), Encoding.UTF8.GetString(html.ToArray()) })
        {
            Assert.Contains("rgb(255,0,0)", markup);
            Assert.Contains("rgb(0,0,255)", markup);
        }
        Retain("composite-paint-resources", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task NumericPaintResourceIdsResolveEquivalentSpellings()
    {
        const string resources = "<ColorSpaces><ColorSpace ID='020' Type='RGB'><Palette><CV>255 0 0</CV></Palette></ColorSpace></ColorSpaces>" +
            "<DrawParams><DrawParam ID='031'><FillColor ColorSpace='20' Index='0'/></DrawParam>" +
            "<DrawParam ID='030' Relative='31'/></DrawParams>";
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' DrawParam='30' Fill='true' Stroke='false'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("rgb(255,0,0)", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("rgb(255,0,0)", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 0);
        Retain("numeric-paint-ids", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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
    [InlineData(false, "valid")]
    [InlineData(true, "valid")]
    [InlineData(false, "attribute")]
    [InlineData(true, "attribute")]
    [InlineData(false, "instruction")]
    [InlineData(true, "instruction")]
    [InlineData(false, "text")]
    [InlineData(true, "text")]
    [InlineData(false, "child")]
    [InlineData(true, "child")]
    public async Task ReferencedPaintContainerExtensionsFailClosed(bool drawParam, string extension)
    {
        var containerName = drawParam ? "DrawParams" : "ColorSpaces";
        var definition = drawParam
            ? "<DrawParam ID='30'><FillColor Value='255 0 0'/></DrawParam>"
            : "<ColorSpace ID='20' Type='RGB'/>";
        var attr = extension == "attribute" ? " xmlns:ext='urn:vendor' ext:keep='1'" : "";
        var content = extension switch
        {
            "instruction" => "<?vendor keep?>",
            "text" => "vendor content",
            "child" => "<ext:State xmlns:ext='urn:vendor'/>",
            _ => " \n<!-- harmless comment -->"
        };
        var resources = $"<{containerName}{attr}>{content}{definition}</{containerName}>";
        var paint = drawParam ? "DrawParam='30'" : "";
        var color = drawParam ? "" : "<FillColor ColorSpace='20' Value='255 0 0'/>";
        var drawing = $"<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false' {paint}>" +
            color + "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(drawing, resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains(containerName, Encoding.UTF8.GetString(read.PreservedEntries["Doc_0/DocumentRes.xml"]));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            if (extension == "valid")
            {
                await convert(new MemoryStream(ofd), output);
                Assert.True(output.Length > 0);
            }
            else
            {
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
                Assert.Contains("paint", error.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(0, output.Length);
            }
        }
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

    [Theory]
    [InlineData("<Palette Vendor='x'><CV>0 0 255</CV></Palette>")]
    [InlineData("<Palette><?vendor key?><CV>0 0 255</CV></Palette>")]
    [InlineData("<Palette><CV>0 0 255</CV><CV>bad</CV></Palette>")]
    public async Task DirectValueDoesNotBypassPaletteAdmission(string palette)
    {
        var resources = $"<ColorSpaces><ColorSpace ID='20' Type='RGB'>{palette}</ColorSpace></ColorSpaces>";
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='20' Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
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
    [InlineData(false, false, false)]
    [InlineData(false, true, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, true, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, true)]
    public async Task PaintDefinitionsRequireOneContainerPerKind(bool drawParam, bool second, bool duplicateContainer)
    {
        var container = drawParam ? "DrawParams" : "ColorSpaces";
        var first = drawParam ? "<DrawParam ID='30'><FillColor Value='255 0 0'/></DrawParam>"
            : "<ColorSpace ID='20' Type='RGB'/>";
        var next = drawParam ? "<DrawParam ID='31'><FillColor Value='0 0 255'/></DrawParam>"
            : "<ColorSpace ID='21' Type='RGB'/>";
        var resources = duplicateContainer
            ? $"<{container}>{first}</{container}><{container}>{next}</{container}>"
            : $"<{container}>{first}{next}</{container}>";
        var reference = drawParam
            ? $" DrawParam='{(second ? "31" : "30")}'"
            : $"<FillColor ColorSpace='{(second ? "21" : "20")}' Value='0 0 255'/>";
        var body = drawParam
            ? $"<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'{reference}>" +
              "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>"
            : "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" + reference +
              "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body, resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var path = Assert.IsType<OfdPathElement>(Assert.Single(block.PageBlockChildren!));
        if (duplicateContainer) Assert.Contains(drawParam ? "DrawParam" : "ColorSpace", path.RenderPaintError);
        else Assert.Null(path.RenderPaintError);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            if (duplicateContainer)
            {
                await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
                Assert.Equal(0, sink.Length);
            }
            else
            {
                await convert(new MemoryStream(ofd), sink);
                Assert.True(sink.Length > 0);
            }
        }
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
    [InlineData("Bold", "020")]
    [InlineData("Italic", "20")]
    public async Task MalformedFontFlagPreservesPageBlockTextButFailsAllRenderers(string flag, string fontId)
    {
        var resources = $"<Fonts><Font ID='{fontId}' FontName='Noto Sans' {flag}='bogus'/></Fonts>";
        var ofd = await Source("<TextObject ID='501' Boundary='10 10 50 10' Font='20' Size='4'>" +
            "<TextCode X='0' Y='4'>recoverable text</TextCode></TextObject>", resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains(read.Fonts, font => font.Id == fontId);
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
    [InlineData("public", false)]
    [InlineData("document", false)]
    [InlineData("page", false)]
    [InlineData("template", false)]
    [InlineData("public", true)]
    [InlineData("document", true)]
    [InlineData("page", true)]
    [InlineData("template", true)]
    public async Task ExplicitMissingFontPayloadIsNotNameOnlyFallback(string scope, bool nameOnly)
    {
        const string drawing = "<TextObject ID='501' Boundary='10 10 50 10' Font='30' Size='4'>" +
            "<TextCode X='0' Y='4'>font declaration</TextCode></TextObject>";
        var seed = await Source(drawing, template: scope == "template");
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            XNamespace ns = Ns;
            var resource = new XElement(ns + "Res", new XAttribute("BaseLoc", "Res"),
                new XElement(ns + "Fonts", new XElement(ns + "Font", new XAttribute("ID", "30"),
                    new XAttribute("FontName", "Noto Sans"), nameOnly ? null : new XElement(ns + "FontFile", "missing.ttf"))));
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (scope == "public" && entry.FullName == "Doc_0/PublicRes.xml")
                    bytes = Encoding.UTF8.GetBytes(resource.ToString(SaveOptions.DisableFormatting));
                if (scope == "document" && entry.FullName == "Doc_0/Document.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    xml.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "DocumentRes", "DocumentRes.xml"));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
            if (scope != "public")
            {
                var path = scope switch
                {
                    "document" => "Doc_0/DocumentRes.xml",
                    "page" => "Doc_0/Pages/Page_0/PageRes.xml",
                    _ => "Doc_0/Tpl_0/PageRes.xml"
                };
                using var target = output.CreateEntry(path).Open();
                await target.WriteAsync(Encoding.UTF8.GetBytes(resource.ToString(SaveOptions.DisableFormatting)));
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var declaredFont = Assert.Single(read.Fonts);
        Assert.Equal(nameOnly ? null : "missing.ttf", declaredFont.FileName);
        var text = Assert.IsType<OfdTextElement>(Assert.Single(Assert.IsType<OfdRawElement>(
            Assert.Single(scope == "template" ? read.Pages[0].Templates[0].Elements : read.Pages[0].Elements)).PageBlockChildren!));
        Assert.Contains("font declaration", new OfdTextExtractor().Extract(read, includeTemplates: true));
        if (nameOnly) Assert.Null(text.RenderPaintError);
        else Assert.Contains("unresolved Font", text.RenderPaintError);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            if (nameOnly) { await convert(new MemoryStream(ofd), sink); Assert.True(sink.Length > 0); }
            else
            {
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
                Assert.Contains("Font", error.Message);
                Assert.Equal(0, sink.Length);
            }
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageBlockTextWithoutTextCodeFailsAllRenderers(bool template)
    {
        var ofd = await Source("<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4'/>",
            template: template, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(template ? read.Pages[0].Templates[0].Elements : read.Pages[0].Elements));
        var text = Assert.IsType<OfdTextElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Empty(text.Runs);
        Assert.Contains("TextCode", text.RenderPaintError);
        Assert.Contains("TextObject", text.SourceXml);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Contains("TextCode", error.Message);
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData("duplicate-file")]
    [InlineData("font-attribute")]
    [InlineData("font-child")]
    [InlineData("container-attribute")]
    [InlineData("container-child")]
    [InlineData("file-attribute")]
    [InlineData("file-child")]
    public async Task AmbiguousFontDeclarationCannotResolvePageBlockText(string change)
    {
        var seed = await Source("<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>recoverable text</TextCode></TextObject>", font: true);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName == "Doc_0/PublicRes.xml")
                {
                    var xml = XDocument.Load(new MemoryStream(bytes));
                    var fonts = xml.Descendants().Single(node => node.Name.LocalName == "Fonts");
                    var font = fonts.Elements().Single(node => node.Name.LocalName == "Font");
                    var file = font.Elements().Single(node => node.Name.LocalName == "FontFile");
                    switch (change)
                    {
                        case "duplicate-file": font.Add(new XElement(file)); break;
                        case "font-attribute": font.SetAttributeValue("Vendor", "opaque"); break;
                        case "font-child": font.Add(new XElement(font.Name.Namespace + "Vendor")); break;
                        case "container-attribute": fonts.SetAttributeValue("Vendor", "opaque"); break;
                        case "container-child": fonts.Add(new XElement(fonts.Name.Namespace + "Vendor")); break;
                        case "file-attribute": file.SetAttributeValue("Vendor", "opaque"); break;
                        case "file-child": file.Add(new XElement(file.Name.Namespace + "Vendor")); break;
                    }
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains("recoverable text", new OfdTextExtractor().Extract(read));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var text = Assert.IsType<OfdTextElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Contains("unresolved Font", text.RenderPaintError);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Contains("Font", error.Message);
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData("Bold")]
    [InlineData("Italic")]
    public async Task MalformedCompositeFontFlagPreservesTextButFailsAllRenderers(string flag)
    {
        var resources = $"<Fonts><Font ID='550' FontName='Noto Sans' {flag}='bogus'/></Fonts>" +
            "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='550' Size='4'>" +
            "<TextCode X='0' Y='4'>composite text</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        var body = "<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>" +
            "<TextObject ID='554' Boundary='10 40 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>sibling text</TextCode></TextObject>";
        var ofd = await Source(body, resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var extracted = new OfdTextExtractor().Extract(read);
        Assert.Contains("composite text", extracted);
        Assert.Contains("sibling text", extracted);
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
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageLocalFontCannotResolveLaterPageBlockOrTemplate(bool template)
    {
        static OfdRawElement Block(string id, string child) => new()
        {
            LocalName = "PageBlock",
            Xml = $"<PageBlock xmlns='{Ns}' ID='{id}'>{child}</PageBlock>"
        };
        const string path = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        const string text = "<TextObject ID='711' Boundary='10 10 50 10' Font='30' Size='4'>" +
            "<TextCode X='0' Y='4'>later page text</TextCode></TextObject>";
        var package = new OfdDocumentPackage();
        var first = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        first.Elements.Add(Block("500", path)); package.Pages.Add(first);
        var second = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        second.Elements.Add(Block("700", template ? path.Replace("501", "701", StringComparison.Ordinal) : text));
        package.Pages.Add(second);
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed.ToArray()), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            XNamespace ns = Ns;
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (template && entry.FullName == "Doc_0/Document.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    xml.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage",
                        new XAttribute("ID", "750"), new XAttribute("BaseLoc", "Tpl_0/Content.xml")));
                    xml.Root.Element(ns + "CommonData")!.Element(ns + "MaxUnitID")!.Value = "800";
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                if (template && entry.FullName.EndsWith("/Pages/Page_1/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    xml.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "750")));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
            var pageRes = new XElement(ns + "Res", new XElement(ns + "Fonts",
                new XElement(ns + "Font", new XAttribute("ID", "30"), new XAttribute("FontName", "Noto Sans"))));
            using (var pageResStream = output.CreateEntry("Doc_0/Pages/Page_0/PageRes.xml").Open())
                await pageResStream.WriteAsync(Encoding.UTF8.GetBytes(pageRes.ToString(SaveOptions.DisableFormatting)));
            if (template)
            {
                var templatePage = new XElement(ns + "Page", new XElement(ns + "Content",
                    new XElement(ns + "Layer", new XAttribute("ID", "760"),
                        XElement.Parse($"<PageBlock xmlns='{Ns}' ID='761'>{text.Replace("711", "762", StringComparison.Ordinal)}</PageBlock>"))));
                using var templateStream = output.CreateEntry("Doc_0/Tpl_0/Content.xml").Open();
                await templateStream.WriteAsync(Encoding.UTF8.GetBytes(templatePage.ToString(SaveOptions.DisableFormatting)));
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var laterBlock = Assert.IsType<OfdRawElement>(Assert.Single(template
            ? read.Pages[1].Templates[0].Elements : read.Pages[1].Elements));
        var laterText = Assert.IsType<OfdTextElement>(Assert.Single(laterBlock.PageBlockChildren!));
        Assert.Contains("unresolved Font", laterText.RenderPaintError);
        Assert.Contains("later page text", new OfdTextExtractor().Extract(read, includeTemplates: template));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output, pageIndex: 1),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output, pageIndexes: new[] { 1 })
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("Font", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SharedTemplateUsesOnlyDocumentAndTemplateFonts(bool templateOwnFont)
    {
        var templateFontId = templateOwnFont ? "40" : "30";
        var package = new OfdDocumentPackage();
        for (var index = 0; index < 2; index++)
        {
            var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
            page.Elements.Add(new OfdRawElement { LocalName = "PageBlock",
                Xml = $"<PageBlock xmlns='{Ns}' ID='{500 + index * 100}'>" +
                    $"<PathObject ID='{501 + index * 100}' Boundary='60 60 20 20' Fill='true' Stroke='false'>" +
                    "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject></PageBlock>" });
            package.Pages.Add(page);
        }
        using var seed = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, seed);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed.ToArray()), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            XNamespace ns = Ns;
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName == "Doc_0/Document.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    xml.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage",
                        new XAttribute("ID", "750"), new XAttribute("BaseLoc", "Tpl_0/Content.xml")));
                    xml.Root.Element(ns + "CommonData")!.Element(ns + "MaxUnitID")!.Value = "800";
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                if (entry.FullName.EndsWith("/Content.xml", StringComparison.Ordinal) &&
                    entry.FullName.Contains("/Pages/Page_", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    xml.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "750")));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
            for (var index = 0; index < 2; index++)
            {
                var pageRes = new XElement(ns + "Res", new XElement(ns + "Fonts",
                    new XElement(ns + "Font", new XAttribute("ID", index == 0 ? "30" : "31"),
                        new XAttribute("FontName", "Noto Sans"))));
                using var stream = output.CreateEntry($"Doc_0/Pages/Page_{index}/PageRes.xml").Open();
                await stream.WriteAsync(Encoding.UTF8.GetBytes(pageRes.ToString(SaveOptions.DisableFormatting)));
            }
            var templateText = XElement.Parse($"<PageBlock xmlns='{Ns}' ID='761'>" +
                $"<TextObject ID='762' Boundary='10 10 50 10' Font='{templateFontId}' Size='4'>" +
                "<TextCode X='0' Y='4'>shared template text</TextCode></TextObject></PageBlock>");
            var templatePage = new XElement(ns + "Page", new XElement(ns + "Content",
                new XElement(ns + "Layer", new XAttribute("ID", "760"), templateText)));
            using (var stream = output.CreateEntry("Doc_0/Tpl_0/Content.xml").Open())
                await stream.WriteAsync(Encoding.UTF8.GetBytes(templatePage.ToString(SaveOptions.DisableFormatting)));
            if (templateOwnFont)
            {
                var templateRes = new XElement(ns + "Res", new XElement(ns + "Fonts",
                    new XElement(ns + "Font", new XAttribute("ID", "40"), new XAttribute("FontName", "Noto Sans"),
                        new XElement(ns + "FontFile", "NotoSans.ttf"))));
                using (var stream = output.CreateEntry("Doc_0/Tpl_0/PageRes.xml").Open())
                    await stream.WriteAsync(Encoding.UTF8.GetBytes(templateRes.ToString(SaveOptions.DisableFormatting)));
                var directory = new DirectoryInfo(AppContext.BaseDirectory);
                while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln")))
                    directory = directory.Parent;
                var fontBytes = File.ReadAllBytes(Path.Combine(directory!.FullName,
                    "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
                using var templateFontStream = output.CreateEntry("Doc_0/Tpl_0/NotoSans.ttf").Open();
                await templateFontStream.WriteAsync(fontBytes);
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Equal(2, read.Pages.Count);
        foreach (var page in read.Pages)
        {
            var block = Assert.IsType<OfdRawElement>(Assert.Single(Assert.Single(page.Templates).Elements));
            var text = Assert.IsType<OfdTextElement>(Assert.Single(block.PageBlockChildren!));
            if (templateOwnFont) Assert.Null(text.RenderPaintError);
            else Assert.Contains("unresolved Font", text.RenderPaintError);
        }
        Assert.Contains("shared template text", new OfdTextExtractor().Extract(read, includeTemplates: true));
        using var pdf = new MemoryStream();
        using var svg = new MemoryStream();
        using var html = new MemoryStream();
        if (templateOwnFont)
        {
            await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
            await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg, pageIndex: 0);
            await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html, pageIndexes: new[] { 0, 1 });
            Assert.True(pdf.Length > 0);
            Assert.Contains("shared template text", Encoding.UTF8.GetString(svg.ToArray()));
            Assert.Contains("shared template text", Encoding.UTF8.GetString(html.ToArray()));
            Retain("template-own-font", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
        }
        else
        {
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf));
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg, pageIndex: 0));
            await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html, pageIndexes: new[] { 0, 1 }));
            Assert.Equal(0, pdf.Length); Assert.Equal(0, svg.Length); Assert.Equal(0, html.Length);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DuplicateFontIdsAreRejectedBeforeAnyRenderer(bool acrossResourceFiles)
    {
        var id = acrossResourceFiles ? "10" : "20";
        var resources = $"<Fonts><Font ID='{id}' FontName='First Face'/>" +
            (acrossResourceFiles ? "" : $"<Font ID='{id}' FontName='Second Face'/>") +
            "</Fonts><VendorExtension Key='keep'/>" +
            "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
            $"<TextObject ID='552' Boundary='0 0 50 10' Font='{id}' Size='4'>" +
            "<TextCode X='0' Y='4'>composite text</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        var body = "<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>" +
            $"<TextObject ID='554' Boundary='10 40 50 10' Font='{id}' Size='4'>" +
            "<TextCode X='0' Y='4'>sibling text</TextCode></TextObject>";
        var ofd = await Source(body, resources: resources, font: acrossResourceFiles);
        // SourceIdentity rejects repeated active numeric IDs across and within
        // resource files before a Reader package can reach any font consumer.
        var readError = await Assert.ThrowsAsync<InvalidDataException>(() =>
            new OfdReader().ReadAsync(new MemoryStream(ofd)));
        Assert.Contains("Duplicate active document-unit ID", readError.Message);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<InvalidDataException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("Duplicate active document-unit ID", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task UniqueFontIdsAcrossResourcesStillRender()
    {
        const string resources = "<Fonts><Font ID='20' FontName='Noto Sans'/></Fonts><VendorExtension Key='keep'/>" +
            "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='20' Size='4'>" +
            "<TextCode X='0' Y='4'>composite text</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        const string body = "<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>" +
            "<TextObject ID='554' Boundary='10 40 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>sibling text</TextCode></TextObject>";
        var ofd = await Source(body, resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var composite = Assert.IsType<OfdCompositeElement>(block.PageBlockChildren![0]);
        var text = Assert.IsType<OfdTextElement>(Assert.Single(composite.Resource.Elements));
        Assert.Null(text.RenderPaintError);
        var extracted = new OfdTextExtractor().Extract(read);
        Assert.Contains("composite text", extracted);
        Assert.Contains("sibling text", extracted);
        Assert.Contains("VendorExtension", Encoding.UTF8.GetString(read.PreservedEntries["Doc_0/DocumentRes.xml"]));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 0);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("composite text", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("composite text", Encoding.UTF8.GetString(html.ToArray()));
        Retain("unique-font-resources", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("foreign-attribute")]
    [InlineData("processing-instruction")]
    [InlineData("unexpected-sibling")]
    [InlineData("nonwhitespace-text")]
    [InlineData("second-container")]
    public async Task ExtendedCompositeContainerDoesNotRenderFromPageBlock(string change)
    {
        XNamespace ns = Ns;
        var container = new XElement(ns + "CompositeGraphicUnits", new XElement(ns + "CompositeGraphicUnit",
            new XAttribute("ID", "551"), new XAttribute("Width", "50"), new XAttribute("Height", "20"),
            new XElement(ns + "TextObject", new XAttribute("ID", "552"), new XAttribute("Boundary", "0 0 50 10"),
                new XAttribute("Font", "10"), new XAttribute("Size", "4"),
                new XElement(ns + "TextCode", new XAttribute("X", "0"), new XAttribute("Y", "4"), "中文资源文字"))));
        switch (change)
        {
            case "foreign-attribute": container.SetAttributeValue(XName.Get("opaque", "urn:vendor"), "value"); break;
            case "processing-instruction": container.Add(new XProcessingInstruction("vendor", "opaque")); break;
            case "unexpected-sibling": container.Add(new XElement(ns + "Vendor")); break;
            case "nonwhitespace-text": container.Add(new XText("opaque")); break;
        }
        var resources = container.ToString(SaveOptions.DisableFormatting) +
            (change == "second-container" ? new XElement(ns + "CompositeGraphicUnits").ToString(SaveOptions.DisableFormatting) : "");
        var ofd = await Source("<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>",
            resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains("中文资源文字", new OfdTextExtractor().Extract(read));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var composite = Assert.IsType<OfdCompositeElement>(Assert.Single(block.PageBlockChildren!));
        Assert.False(composite.Resource.SourceContainerSupported);
        Assert.Contains("CompositeGraphicUnits", Encoding.UTF8.GetString(read.PreservedEntries["Doc_0/DocumentRes.xml"]));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Contains("container", error.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public async Task ExplicitGlyphMappingInPageBlockFailsClosedWithoutLosingText(bool composite, bool mapped)
    {
        var glyphs = mapped
            ? "<CGTransform CodePosition='0' CodeCount='1' GlyphCount='1'><Glyphs>2</Glyphs></CGTransform>"
            : "";
        var text = "<TextObject ID='552' Boundary='0 0 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>logical text</TextCode>" + glyphs + "</TextObject>";
        var resources = composite
            ? "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
              text + "</CompositeGraphicUnit></CompositeGraphicUnits>"
            : null;
        var body = composite
            ? "<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>"
            : text;
        var ofd = await Source(body, resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains("logical text", new OfdTextExtractor().Extract(read));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            if (mapped)
            {
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
                Assert.Contains("glyph", error.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(0, output.Length);
            }
            else
            {
                await convert(new MemoryStream(ofd), output);
                Assert.True(output.Length > 0);
            }
        }
    }

    [Theory]
    [InlineData("50", "20", true)]
    [InlineData(null, "20", false)]
    [InlineData("0", "20", false)]
    [InlineData("bogus", "20", false)]
    [InlineData("NaN", "20", false)]
    [InlineData("50", null, false)]
    public async Task CompositeResourceDimensionsMustBeExplicitAndPositive(string? width, string? height, bool valid)
    {
        var widthAttribute = width is null ? "" : $" Width='{width}'";
        var heightAttribute = height is null ? "" : $" Height='{height}'";
        var resources = $"<CompositeGraphicUnits><CompositeGraphicUnit ID='551'{widthAttribute}{heightAttribute}>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>resource text</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        var ofd = await Source("<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>",
            resources: resources, font: true);
        if (valid)
        {
            var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
            Assert.Contains("resource text", new OfdTextExtractor().Extract(read));
        }
        else
        {
            var readError = await Assert.ThrowsAsync<InvalidDataException>(() =>
                new OfdReader().ReadAsync(new MemoryStream(ofd)));
            Assert.Contains("Composite dimensions", readError.Message);
        }
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            if (valid)
            {
                await convert(new MemoryStream(ofd), output);
                Assert.True(output.Length > 0);
            }
            else
            {
                var error = await Assert.ThrowsAsync<InvalidDataException>(() => convert(new MemoryStream(ofd), output));
                Assert.Contains("Composite dimensions", error.Message);
                Assert.Equal(0, output.Length);
            }
        }
    }

    [Fact]
    public async Task NumericCompositeIdsResolveEquivalentSpellingsInPageBlock()
    {
        const string resources = "<CompositeGraphicUnits><CompositeGraphicUnit ID='0551' Width='50' Height='20'>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>numeric composite</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        const string body = "<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>";
        var ofd = await Source(body, resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains(read.CompositeResources!.Items, resource => resource.Id == "0551");
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var composite = Assert.IsType<OfdCompositeElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Equal("0551", composite.Resource.Id);
        Assert.Contains("numeric composite", new OfdTextExtractor().Extract(read));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("numeric composite", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("numeric composite", Encoding.UTF8.GetString(html.ToArray()));
        Retain("numeric-composite-ids", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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

    [Fact]
    public async Task NumericFontIdsResolveEquivalentSpellingsInPageBlockAndComposite()
    {
        const string resources = "<Fonts><Font ID='030' FontName='Noto Sans'/></Fonts>" +
            "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='10'>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='30' Size='4'>" +
            "<TextCode X='0' Y='4'>composite font</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        const string body = "<TextObject ID='501' Boundary='10 10 50 10' Font='30' Size='4'>" +
            "<TextCode X='0' Y='4'>pageblock font</TextCode></TextObject>" +
            "<CompositeObject ID='553' Boundary='10 30 50 10' ResourceID='551'/>";
        var ofd = await Source(body, resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains(read.Fonts, font => font.Id == "030" && font.FontName == "Noto Sans");
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var direct = Assert.IsType<OfdTextElement>(block.PageBlockChildren![0]);
        var composite = Assert.IsType<OfdCompositeElement>(block.PageBlockChildren[1]);
        var nested = Assert.IsType<OfdTextElement>(Assert.Single(composite.Resource.Elements));
        Assert.Equal("Noto Sans", direct.FontName);
        Assert.Equal("Noto Sans", nested.FontName);
        Assert.Null(direct.RenderPaintError);
        Assert.Null(nested.RenderPaintError);
        var extracted = new OfdTextExtractor().Extract(read);
        Assert.Contains("pageblock font", extracted);
        Assert.Contains("composite font", extracted);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 0);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("pageblock font", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("composite font", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("pageblock font", Encoding.UTF8.GetString(html.ToArray()));
        Assert.Contains("composite font", Encoding.UTF8.GetString(html.ToArray()));
        Retain("numeric-font-ids", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task NumericFontReferenceSelectsItsFaceBeforeSameNamedFallback()
    {
        const string resources = "<Fonts><Font ID='20' FontName='Noto Sans' Bold='false'/>" +
            "<Font ID='030' FontName='Noto Sans' Bold='true'/></Fonts>" +
            "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='10'>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='30' Size='4'>" +
            "<TextCode X='0' Y='4'>nested face</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        const string body = "<TextObject ID='501' Boundary='10 10 50 10' Font='30' Size='4'>" +
            "<TextCode X='0' Y='4'>direct face</TextCode></TextObject>" +
            "<CompositeObject ID='553' Boundary='10 30 50 10' ResourceID='551'/>";
        var ofd = await Source(body, resources: resources);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var direct = Assert.IsType<OfdTextElement>(block.PageBlockChildren![0]);
        var composite = Assert.IsType<OfdCompositeElement>(block.PageBlockChildren[1]);
        var nested = Assert.IsType<OfdTextElement>(Assert.Single(composite.Resource.Elements));
        foreach (var text in new[] { direct, nested })
        {
            Assert.Equal("30", text.FontResourceId);
            var selected = Assert.IsType<OfdFontResource>(OfdFontSelection.Resolve(read.Fonts, text));
            Assert.Equal("030", selected.Id);
            Assert.True(selected.Bold);
        }
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("direct face", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("nested face", Encoding.UTF8.GetString(html.ToArray()));
        Retain("numeric-font-face-ids", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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
    [InlineData("duplicate-file", false)]
    [InlineData("duplicate-file", true)]
    [InlineData("both-selectors", false)]
    [InlineData("media-attribute", false)]
    [InlineData("media-child", false)]
    [InlineData("media-pi", false)]
    [InlineData("container-attribute", false)]
    [InlineData("container-pi", false)]
    [InlineData("container-text", false)]
    [InlineData("file-attribute", false)]
    [InlineData("file-pi", false)]
    public async Task AmbiguousMediaDeclarationCannotResolvePageBlockImage(string change, bool template)
    {
        XNamespace ns = Ns;
        var container = new XElement(ns + "MultiMedias", new XElement(ns + "MultiMedia",
            new XAttribute("ID", "20"), new XAttribute("Type", "Image"), new XAttribute("Format", "PNG"),
            new XElement(ns + "MediaFile", "image.png")));
        var media = container.Element(ns + "MultiMedia")!;
        var file = media.Element(ns + "MediaFile")!;
        switch (change)
        {
            case "duplicate-file": media.Add(new XElement(file)); break;
            case "both-selectors": media.SetAttributeValue("MediaFile", "image.png"); break;
            case "media-attribute": media.SetAttributeValue("Vendor", "opaque"); break;
            case "media-child": media.Add(new XElement(ns + "Vendor")); break;
            case "media-pi": media.Add(new XProcessingInstruction("vendor", "opaque")); break;
            case "container-attribute": container.SetAttributeValue("Vendor", "opaque"); break;
            case "container-pi": container.Add(new XProcessingInstruction("vendor", "opaque")); break;
            case "container-text": container.Add(new XText("opaque")); break;
            case "file-attribute": file.SetAttributeValue("Vendor", "opaque"); break;
            case "file-pi": file.Add(new XProcessingInstruction("vendor", "opaque")); break;
        }
        var body = "<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='20'/>" +
            "<TextObject ID='502' Boundary='10 40 50 10' Size='4'><TextCode X='0' Y='4'>中文资源文字</TextCode></TextObject>";
        var ofd = await Source(body, template: template, resources: container.ToString(SaveOptions.DisableFormatting), image: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(template ? read.Pages[0].Templates[0].Elements : read.Pages[0].Elements));
        var image = Assert.IsType<OfdImageElement>(block.PageBlockChildren![0]);
        Assert.Empty(image.Data);
        Assert.Contains("中文资源文字", new OfdTextExtractor().Extract(read, includeTemplates: true));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task PageBlockCompositeImageChecksMediaDeclaration(bool ambiguous)
    {
        var media = $"<MultiMedias{(ambiguous ? " Vendor='opaque'" : "")}><MultiMedia ID='20' Type='Image' Format='PNG'>" +
            "<MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>";
        const string composite = "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
            "<ImageObject ID='552' Boundary='0 0 10 10' ResourceID='20'/>" +
            "<TextObject ID='553' Boundary='12 0 35 10' Font='10' Size='4'><TextCode X='0' Y='4'>nested media</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        var ofd = await Source("<CompositeObject ID='554' Boundary='10 10 50 20' ResourceID='551'/>",
            resources: media + composite, font: true, image: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var placement = Assert.IsType<OfdCompositeElement>(Assert.Single(block.PageBlockChildren!));
        var image = Assert.IsType<OfdImageElement>(placement.Resource.Elements[0]);
        Assert.NotEmpty(image.Data);
        Assert.Equal(ambiguous, image.SourceMediaAmbiguous);
        Assert.Contains("nested media", new OfdTextExtractor().Extract(read));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            if (ambiguous)
            {
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
                Assert.Contains("media", error.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Equal(0, sink.Length);
            }
            else
            {
                await convert(new MemoryStream(ofd), sink);
                Assert.True(sink.Length > 0);
            }
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
    [InlineData(false)]
    [InlineData(true)]
    public async Task NumericMediaIdsResolveEquivalentSpellingsInPageBlock(bool template)
    {
        const string media = "<MultiMedias><MultiMedia ID='030' Type='Image' Format='PNG'>" +
            "<MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>";
        var ofd = await Source("<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='30'/>",
            template: template, resources: media, image: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(template ? read.Pages[0].Templates[0].Elements : read.Pages[0].Elements));
        var image = Assert.IsType<OfdImageElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Equal("30", image.ResourceId);
        Assert.NotEmpty(image.Data);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(html.ToArray()));
        Retain(template ? "numeric-media-ids-template" : "numeric-media-ids", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("Audio", false)]
    [InlineData("Video", true)]
    [InlineData("Other", false)]
    public async Task ExplicitNonImageMediaCannotResolvePageBlockImage(string type, bool template)
    {
        var media = $"<MultiMedias><MultiMedia ID='30' Type='{type}' Format='PNG'>" +
            "<MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>";
        var ofd = await Source("<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='30'/>",
            template: template, resources: media, image: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(template ? read.Pages[0].Templates[0].Elements : read.Pages[0].Elements));
        var image = Assert.IsType<OfdImageElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Empty(image.Data);
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
    public async Task OmittedMediaTypeRetainsLegacyPageBlockImageLookup()
    {
        const string media = "<MultiMedias><MultiMedia ID='30' Format='PNG'>" +
            "<MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>";
        var ofd = await Source("<ImageObject ID='501' Boundary='10 10 20 20' ResourceID='30'/>",
            resources: media, image: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.NotEmpty(Assert.IsType<OfdImageElement>(Assert.Single(block.PageBlockChildren!)).Data);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("data:image/png;base64,", Encoding.UTF8.GetString(html.ToArray()));
        Retain("media-type-omitted", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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
    [InlineData("")]
    [InlineData("<AbbreviatedData/>")]
    [InlineData("<AbbreviatedData>M 0 0 L 20</AbbreviatedData>")]
    [InlineData("<AbbreviatedData>M 1<!--a--><!--b-->0 2 3</AbbreviatedData>")]
    [InlineData("<AbbreviatedData>M 0 0 L 20 0 C</AbbreviatedData><AbbreviatedData>M 0 0 L 20 20</AbbreviatedData>")]
    [InlineData("<AbbreviatedData>M 0 0 L 10 0 L 10 10 C,</AbbreviatedData>")]
    [InlineData("<AbbreviatedData>M 0 0 L 10 0 L 10 10, </AbbreviatedData>")]
    public async Task PageBlockPathDataFailsBeforeAnyRenderer(string pathData)
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor Value='0 0 255'/>" + pathData + "</PathObject>";
        var ofd = await Source(body);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("path data", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task CommaSeparatedPageBlockPathStillRenders()
    {
        var ofd = await Source("<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0,0 L 10,0 L 10,10 C</AbbreviatedData></PathObject>");
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
            (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
        })
        {
            using var sink = new MemoryStream();
            await convert(new MemoryStream(ofd), sink);
            Assert.True(sink.Length > 0);
        }
    }

    [Fact]
    public async Task PageBlockPathKeepsWhitespaceSeparatedByComments()
    {
        const string path = "M 0 0 L 20 0 L 20 20 C";
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0<!--a--> <!--b-->0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.Equal(path, Assert.IsType<OfdPathElement>(Assert.Single(block.PageBlockChildren!)).AbbreviatedData);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("M 0 0 L 20 0 L 20 20 Z", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("M 0 0 L 20 0 L 20 20 Z", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 0);
        Retain("pageblock-comment-whitespace", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("0 10 0 0 1")]
    [InlineData("-10 10 0 0 1")]
    [InlineData("10 10 0 2 2")]
    [InlineData("10 10 0 -1 1")]
    [InlineData("10 10 0 1.0 0")]
    [InlineData("10 10 0 +1 0")]
    public async Task PageBlockArcRejectsInvalidRadiiAndFlags(string operands)
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='false' Stroke='true'>" +
            $"<StrokeColor Value='0 0 255'/><AbbreviatedData>M 0 10 A {operands} 20 10</AbbreviatedData></PathObject>";
        var ofd = await Source(body);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var output = new MemoryStream();
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
            Assert.Contains("path data", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task PageBlockValidArcRendersInAllOutputs()
    {
        const string body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='false' Stroke='true'>" +
            "<StrokeColor Value='0 0 255'/><AbbreviatedData>M 0 10 A 10 10 0 0 1 20 10</AbbreviatedData></PathObject>";
        var ofd = await Source(body);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("A 10 10 0 0 1 20 10", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("A 10 10 0 0 1 20 10", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        Assert.True(pdf.Length > 0);
        Retain("pageblock-arc", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("0 1", true)]
    [InlineData("2 2", false)]
    [InlineData("1.0 0", false)]
    public async Task PageBlockClipArcUsesTheSameFlagAdmission(string flags, bool supported)
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<Clips><Clip><Area><Path Stroke='false' Fill='true'>" +
            $"<AbbreviatedData>M 0 10 A 10 10 0 {flags} 20 10 L 20 20 L 0 20 C</AbbreviatedData>" +
            "</Path></Area></Clip></Clips>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body);
        var converters = new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        };
        if (!supported)
        {
            foreach (var convert in converters)
            {
                using var output = new MemoryStream();
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
                Assert.Contains("clip path data", error.Message);
                Assert.Equal(0, output.Length);
            }
            return;
        }
        using var pdf = new MemoryStream(); await converters[0](new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await converters[1](new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await converters[2](new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("A 10 10 0 0 1 20 10", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("A 10 10 0 0 1 20 10", Encoding.UTF8.GetString(html.ToArray()));
        Retain("pageblock-clip-arc", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(" false ", " true ", true)]
    [InlineData(" 0 ", " 1 ", true)]
    [InlineData(" true ", " true ", false)]
    [InlineData(" false ", " false ", false)]
    [InlineData("invalid", "true", false)]
    public async Task RetainedClipBooleanWhitespaceMatchesRenderedGeometry(string stroke, string fill, bool supported)
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            $"<Clips><Clip><Area><Path Stroke='{stroke}' Fill='{fill}'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 0 20 C</AbbreviatedData>" +
            "</Path></Area></Clip></Clips>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var path = Assert.IsType<OfdPathElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Null(path.Clips);
        Assert.Contains(stroke, path.ClippingXml);
        Assert.Contains(fill, path.ClippingXml);
        var converters = new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        };
        if (!supported)
        {
            foreach (var convert in converters)
            {
                using var output = new MemoryStream();
                await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
                Assert.Equal(0, output.Length);
            }
            return;
        }
        using var pdf = new MemoryStream(); await converters[0](new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await converters[1](new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await converters[2](new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("clipPath", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("clipPath", Encoding.UTF8.GetString(html.ToArray()));
        if (stroke == " false ") Retain("clip-whitespace-flags", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("0 1", true)]
    [InlineData("2 2", false)]
    [InlineData("1.0 0", false)]
    public async Task DefaultPageBlockClipArcUsesTheSameFlagAdmission(string flags, bool supported)
    {
        var body = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<Clips><Clip><Area><Path>" +
            $"<AbbreviatedData>M 0 10 A 10 10 0 {flags} 20 10 L 20 20 L 0 20 C</AbbreviatedData>" +
            "</Path></Area></Clip></Clips>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var ofd = await Source(body);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var path = Assert.IsType<OfdPathElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Null(path.Clips);
        Assert.NotNull(path.ClippingXml);
        var converters = new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        };
        if (!supported)
        {
            foreach (var convert in converters)
            {
                using var output = new MemoryStream();
                var error = await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), output));
                Assert.Contains("clip path data", error.Message);
                Assert.Equal(0, output.Length);
            }
            return;
        }
        using var pdf = new MemoryStream(); await converters[0](new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await converters[1](new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await converters[2](new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("A 10 10 0 0 1 20 10", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("A 10 10 0 0 1 20 10", Encoding.UTF8.GetString(html.ToArray()));
        Retain("pageblock-clip-arc-default", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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

    [Fact]
    public async Task WhitespaceTextPaintBooleansRenderInAllOutputs()
    {
        var body = "<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4' Fill=' true ' Stroke=' false '>" +
            "<TextCode X='0' Y='4'>paint text</TextCode></TextObject>";
        var ofd = await Source(body, font: true);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("paint text", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("paint text", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("paint text", parsed.GetPage(1).Text);
        Retain("text-paint-boolean-whitespace", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task WhitespaceIntegerTextBooleansRenderInAllOutputs()
    {
        var body = "<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4' Fill=' 1 ' Stroke=' 0 ' Visible=' 1 '>" +
            "<TextCode X='0' Y='4'>numeric boolean text</TextCode></TextObject>";
        var ofd = await Source(body, font: true);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        Assert.Contains("numeric boolean text", Encoding.UTF8.GetString(svg.ToArray()));
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.Contains("numeric boolean text", Encoding.UTF8.GetString(html.ToArray()));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("numeric boolean text", parsed.GetPage(1).Text);
        Retain("text-paint-integer-whitespace", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WhitespaceIntegerPathBooleansKeepFillAndVisibility(bool composite)
    {
        var visible = "<PathObject ID='552' Boundary='10 10 20 20' Fill=' 1 ' Stroke=' 0 ' Visible=' 1 '>" +
            "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var hidden = "<PathObject ID='553' Boundary='40 10 20 20' Fill=' 1 ' Stroke=' 0 ' Visible=' 0 '>" +
            "<FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var resources = composite ? "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='60' Height='30'>" +
            visible + hidden + "</CompositeGraphicUnit></CompositeGraphicUnits>" : null;
        var body = (composite ? "<CompositeObject ID='501' Boundary='10 10 60 30' ResourceID='551'/>" : visible + hidden) +
            "<TextObject ID='554' Boundary='10 50 80 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>visible sibling</TextCode></TextObject>";
        var ofd = await Source(body, resources: resources, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var parsedPath = composite
            ? Assert.IsType<OfdPathElement>(read.CompositeResources!.Items.Single().Elements.First())
            : Assert.IsType<OfdPathElement>(Assert.IsType<OfdRawElement>(read.Pages[0].Elements.Single()).PageBlockChildren!.First());
        Assert.True(parsedPath.Fill);
        Assert.False(parsedPath.Stroke);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        var markup = Encoding.UTF8.GetString(svg.ToArray());
        Assert.Contains("rgb(0,0,255)", markup);
        Assert.DoesNotContain("rgb(255,0,0)", markup);
        var blue = Assert.Single(XDocument.Parse(markup).Descendants(), node =>
            node.Name.LocalName == "path" && node.Attribute("fill")?.Value == "rgb(0,0,255)");
        Assert.Equal("none", blue.Attribute("stroke")?.Value);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        markup = Encoding.UTF8.GetString(html.ToArray());
        Assert.Contains("rgb(0,0,255)", markup);
        Assert.DoesNotContain("rgb(255,0,0)", markup);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("visible sibling", parsed.GetPage(1).Text);
        Retain(composite ? "path-integer-whitespace-composite" : "path-integer-whitespace-pageblock",
            ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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
    [InlineData("urn:vendor")]
    [InlineData("http://www.ofdspec.org")]
    public async Task ForeignNamespacePageBlockContainerRemainsOpaque(string blockNamespace)
    {
        const string drawing = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var seed = await Source(drawing);
        var ofd = await ReNamespacePageBlock(seed, "urn:vendor", blockNamespace);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var raw = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.Equal("Layer", raw.LocalName);
        Assert.Null(raw.PageBlockChildren);
        Assert.Contains("urn:vendor", raw.Xml);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var sink = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData("http://www.ofdspec.org", "", "layer09-block-unqualified")]
    [InlineData("http://www.ofdspec.org/2016", "", "layer16-block-unqualified")]
    [InlineData("http://www.ofdspec.org", "http://www.ofdspec.org/2016", "layer09-block16")]
    [InlineData("http://www.ofdspec.org/2016", "http://www.ofdspec.org", "layer16-block09")]
    public async Task SupportedMixedPageBlockNamespacesStillRender(string layerNamespace, string blockNamespace, string label)
    {
        const string drawing = "<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>mixed block text</TextCode></TextObject>";
        var ofd = await ReNamespacePageBlock(await Source(drawing, font: true), layerNamespace, blockNamespace);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.NotNull(block.PageBlockChildren);
        Assert.Contains("mixed block text", new OfdTextExtractor().Extract(read));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("mixed block text", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("mixed block text", Encoding.UTF8.GetString(html.ToArray()));
        Retain(label, ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Theory]
    [InlineData("http://www.ofdspec.org", "block16-composite09")]
    [InlineData("", "block16-composite-unqualified")]
    public async Task SupportedMixedCompositeNamespaceStillResolves(string compositeNamespace, string label)
    {
        const string resources = "<CompositeGraphicUnits><CompositeGraphicUnit ID='551' Width='50' Height='20'>" +
            "<TextObject ID='552' Boundary='0 0 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>mixed composite text</TextCode></TextObject>" +
            "</CompositeGraphicUnit></CompositeGraphicUnits>";
        var seed = await Source("<CompositeObject ID='553' Boundary='10 10 50 20' ResourceID='551'/>",
            resources: resources, font: true);
        var ofd = await ReNamespacePageBlock(seed, Ns, "http://www.ofdspec.org/2016", compositeNamespace);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.IsType<OfdCompositeElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Contains("mixed composite text", new OfdTextExtractor().Extract(read));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("mixed composite text", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("mixed composite text", Encoding.UTF8.GetString(html.ToArray()));
        Retain(label, ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task LaterPageBlockTextCodesInheritPreviousStartCoordinates()
    {
        const string drawing = "<TextObject ID='501' Boundary='10 10 50 20' Font='10' Size='4'>" +
            "<TextCode X='3' Y='5'>first</TextCode><TextCode>second</TextCode>" +
            "<TextCode X='9'>third</TextCode><TextCode Y='12'>fourth</TextCode></TextObject>";
        var ofd = await Source(drawing, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var block = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        var text = Assert.IsType<OfdTextElement>(Assert.Single(block.PageBlockChildren!));
        Assert.Equal(new[] { (3d, 5d), (3d, 5d), (9d, 5d), (9d, 12d) },
            text.Runs.Select(run => (run.XMillimeters, run.YMillimeters)));
        Assert.Contains("firstsecondthirdfourth", new OfdTextExtractor().Extract(read));
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Retain("textcode-inherited-start", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
    }

    [Fact]
    public async Task InheritedPageBlockAxesStillDecodeSourceEscapes()
    {
        const string drawing = "<TextObject ID='501' Boundary='10 10 70 20' Font='10' Size='4'>" +
            "<TextCode X='0' Y='5'>lead</TextCode><TextCode X='20'>escaped\\0020text</TextCode>" +
            "<TextCode Y='12'>next\\0020line</TextCode></TextObject>";
        var ofd = await Source(drawing, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var text = Assert.IsType<OfdTextElement>(Assert.Single(Assert.IsType<OfdRawElement>(
            Assert.Single(read.Pages[0].Elements)).PageBlockChildren!));
        Assert.Equal(new[] { (0d, 5d), (20d, 5d), (20d, 12d) },
            text.Runs.Select(run => (run.XMillimeters, run.YMillimeters)));
        Assert.Contains("\\0020", text.Text);
        var logical = new OfdTextExtractor().Extract(read);
        Assert.Contains("leadescaped textnext line", logical);
        Assert.DoesNotContain("\\0020", logical);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
            Assert.Contains("escaped text", parsed.GetPage(1).Text);
        Assert.Contains("escaped text", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("escaped text", Encoding.UTF8.GetString(html.ToArray()));
        Retain("textcode-inherited-escape", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
        text.Runs[1].XMillimeters = 0;
        Assert.Same(text, OfdTextRenderProjection.Create(text, new OfdModelMaterialization(new OfdModelLimits()), default));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnqualifiedPageBlockClipUsesSameNamespaceContract(bool vendorClip)
    {
        const string drawing = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<Clips><Clip><Area><Path Stroke='false' Fill='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 0 20 C</AbbreviatedData>" +
            "</Path></Area></Clip></Clips><FillColor Value='0 0 255'/>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 L 0 20 C</AbbreviatedData></PathObject>";
        var seed = await Source(drawing);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var block = Assert.Single(xml.Descendants(), node => node.Name.LocalName == "PageBlock");
                    var unqualified = Strip(block);
                    if (vendorClip)
                    {
                        var clips = Assert.Single(unqualified.Descendants(), node => node.Name.LocalName == "Clips");
                        clips.Name = XNamespace.Get("urn:vendor") + "Clips";
                    }
                    block.ReplaceWith(unqualified);
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var path = Assert.IsType<OfdPathElement>(Assert.Single(Assert.IsType<OfdRawElement>(
            Assert.Single(read.Pages[0].Elements)).PageBlockChildren!));
        if (vendorClip)
        {
            Assert.Null(path.Clips);
            foreach (var convert in new Func<Stream, Stream, Task>[]
            {
                (source, sink) => new OfdToPdfConverter().ConvertAsync(source, sink),
                (source, sink) => new OfdToSvgConverter().ConvertAsync(source, sink),
                (source, sink) => new OfdToHtmlConverter().ConvertAsync(source, sink)
            })
            {
                using var sink = new MemoryStream();
                await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
                Assert.Equal(0, sink.Length);
            }
            return;
        }
        Assert.NotNull(path.Clips);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 0);
        Assert.Contains("clipPath", Encoding.UTF8.GetString(svg.ToArray()));
        Assert.Contains("clipPath", Encoding.UTF8.GetString(html.ToArray()));
        Retain("unqualified-pageblock-clip", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());

        static XElement Strip(XElement node) => new(node.Name.LocalName,
            node.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration),
            node.Nodes().Select(child => child is XElement element ? Strip(element) : child));
    }

    [Theory]
    [InlineData("<TextCode Y='5'>missing X</TextCode>")]
    [InlineData("<TextCode X='3'>missing Y</TextCode>")]
    public async Task FirstPageBlockTextCodeRequiresBothAxes(string code)
    {
        var ofd = await Source($"<TextObject ID='501' Boundary='10 10 50 20' Font='10' Size='4'>{code}</TextObject>", font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Contains("missing", new OfdTextExtractor().Extract(read));
        var text = Assert.IsType<OfdTextElement>(Assert.Single(Assert.IsType<OfdRawElement>(
            Assert.Single(read.Pages[0].Elements)).PageBlockChildren!));
        Assert.Contains("requires X and Y", text.RenderPaintError);
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var sink = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Equal(0, sink.Length);
        }
    }

    [Fact]
    public async Task ForeignContentCannotExposeSupportedLayerOrPageBlock()
    {
        const string drawing = "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var seed = await Source(drawing);
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var content = Assert.Single(xml.Root!.Elements(), node => node.Name.LocalName == "Content");
                    content.ReplaceWith(new XElement(XNamespace.Get("urn:vendor") + "Content", content.Attributes(), content.Nodes()));
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var raw = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.Equal("Content", raw.LocalName);
        Assert.Null(raw.PageBlockChildren);
        Assert.Contains("urn:vendor", raw.Xml);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(read, new MemoryStream()));
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var sink = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public async Task ForeignLayerIsOpaqueAndWriterRejectsUnmodeledStructure(int children)
    {
        var seed = await Source("<PathObject ID='501' Boundary='10 10 20 20' Fill='true'>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>");
        using var rewritten = new MemoryStream();
        using (var input = new ZipArchive(new MemoryStream(seed), ZipArchiveMode.Read))
        using (var output = new ZipArchive(rewritten, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in input.Entries)
            {
                using var source = entry.Open(); using var buffer = new MemoryStream(); await source.CopyToAsync(buffer);
                var bytes = buffer.ToArray();
                if (entry.FullName.EndsWith("/Pages/Page_0/Content.xml", StringComparison.Ordinal))
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    var layer = Assert.Single(xml.Descendants(), node => node.Name.LocalName == "Layer");
                    var opaque = new XElement(XNamespace.Get("urn:vendor") + "Layer",
                        layer.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration),
                        Enumerable.Range(0, children).Select(index =>
                            new XElement(XNamespace.Get("urn:vendor") + "Node", new XAttribute("ID", 700 + index))));
                    layer.ReplaceWith(opaque);
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
        }
        var ofd = rewritten.ToArray();
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var raw = Assert.IsType<OfdRawElement>(Assert.Single(read.Pages[0].Elements));
        Assert.Equal("Layer", raw.LocalName);
        Assert.Equal(children, XElement.Parse(raw.Xml).Elements().Count());
        Assert.Contains("urn:vendor", raw.Xml);
        using var written = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteAsync(read, written));
        Assert.Equal(0, written.Length);
    }

    [Theory]
    [InlineData("M 0 0")]
    [InlineData("M 0 0 C")]
    [InlineData("M 0 0 M 5 5")]
    public async Task MoveOnlyPageBlockPathFailsEveryRenderer(string data)
    {
        var ofd = await Source($"<PathObject ID='501' Boundary='10 10 20 20' Stroke='true'>" +
            $"<AbbreviatedData>{data}</AbbreviatedData></PathObject>");
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var sink = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Equal(0, sink.Length);
        }
    }

    [Theory]
    [InlineData("M 0 0")]
    [InlineData("M 0 0 C")]
    public async Task MoveOnlyPageBlockClipFailsEveryRenderer(string clipData)
    {
        var ofd = await Source("<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false'>" +
            "<Clips><Clip><Area><Path Stroke='false' Fill='true'>" +
            $"<AbbreviatedData>{clipData}</AbbreviatedData></Path></Area></Clip></Clips>" +
            "<AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>");
        foreach (var convert in new Func<Stream, Stream, Task>[]
        {
            (input, output) => new OfdToPdfConverter().ConvertAsync(input, output),
            (input, output) => new OfdToSvgConverter().ConvertAsync(input, output),
            (input, output) => new OfdToHtmlConverter().ConvertAsync(input, output)
        })
        {
            using var sink = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() => convert(new MemoryStream(ofd), sink));
            Assert.Equal(0, sink.Length);
        }
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task WhitespaceFalseVisibilityHidesPageBlockChild(bool text)
    {
        var hidden = text
            ? "<TextObject ID='501' Boundary='10 10 50 10' Font='10' Size='4' Visible=' false '>" +
              "<TextCode X='0' Y='4'>hidden payload</TextCode></TextObject>"
            : "<PathObject ID='501' Boundary='10 10 20 20' Fill='true' Stroke='false' Visible=' false '>" +
              "<FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var body = hidden + "<TextObject ID='502' Boundary='10 30 50 10' Font='10' Size='4'>" +
            "<TextCode X='0' Y='4'>visible sibling</TextCode></TextObject>";
        var ofd = await Source(body, font: true);
        var read = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        Assert.Equal("visible sibling", new OfdTextExtractor().Extract(read));
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        using var html = new MemoryStream(); await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        foreach (var markup in new[] { Encoding.UTF8.GetString(svg.ToArray()), Encoding.UTF8.GetString(html.ToArray()) })
        {
            Assert.Contains("visible sibling", markup);
            Assert.DoesNotContain("hidden payload", markup);
            Assert.DoesNotContain("rgb(0,0,255)", markup);
        }
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray()))
        {
            Assert.Contains("visible sibling", parsed.GetPage(1).Text);
            Assert.DoesNotContain("hidden payload", parsed.GetPage(1).Text);
        }
        Retain(text ? "hidden-whitespace-text" : "hidden-whitespace-path", ofd, pdf.ToArray(), svg.ToArray(), html.ToArray());
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
