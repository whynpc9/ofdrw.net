using System.Globalization;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class SvgGeometryTests
{
    [Fact]
    public async Task PositionedText_ShouldUseOriginalOriginsAndPreserveSpaces()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        var text = new OfdTextElement { Text = "Wi 中", XMillimeters = 10, YMillimeters = 10 };
        text.Runs.Add(new OfdTextRun { Text = "Wi 中", XMillimeters = 0, YMillimeters = 4, DeltaX = "1 3 2", DeltaY = "0 0 1" });
        page.Elements.Add(text);
        package.Pages.Add(page);
        var svg = await ConvertAsync(package);
        var node = svg.Descendants().Single(element => element.Name.LocalName == "text");
        var glyphs = node.Elements().ToList();
        Assert.Equal("Wi 中", node.Value);
        Assert.Equal(new[] { 0d, 1d, 4d, 6d }, glyphs.Select(glyph => double.Parse(glyph.Attribute("x")!.Value, CultureInfo.InvariantCulture)));
        Assert.Equal(new[] { 4d, 4d, 4d, 5d }, glyphs.Select(glyph => double.Parse(glyph.Attribute("y")!.Value, CultureInfo.InvariantCulture)));
        Assert.Null(node.Attribute("dx"));
        Assert.Equal("preserve", node.Attribute(XNamespace.Xml + "space")?.Value);
    }

    [Fact]
    public async Task EmbeddedFontsAndImageStyles_ShouldBeSelfContained()
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource
        {
            Id = "10", FontName = "Ofdrw Test Face", Bold = true, Italic = true,
            Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "narrow.ttf"))
        });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdTextElement { FontName = "Ofdrw Test Face", FontResourceId = "10", Text = "ABCD" });
        using var pixels = new Image<Rgba32>(2, 2, Color.Red);
        using var png = new MemoryStream();
        pixels.SaveAsPng(png);
        page.Elements.Add(new OfdImageElement
        {
            WidthMillimeters = 20, HeightMillimeters = 20, Alpha = 128,
            Transform = [0, 20, -20, 0, 20, 0], Data = png.ToArray(),
            ClipsXml = "<Clips><Clip><Area><Path><AbbreviatedData>M 0 0 L 10 0 L 10 10 C</AbbreviatedData></Path></Area></Clip></Clips>"
        });
        package.Pages.Add(page);
        var svg = await ConvertAsync(package);
        var text = svg.Descendants().Single(element => element.Name.LocalName == "text");
        var css = svg.Descendants().Single(element => element.Name.LocalName == "style").Value;
        Assert.Contains("data:font/ttf;base64,", css);
        Assert.Contains("font-weight:normal", css); // real file is regular; requested bold is synthesized
        Assert.Equal("bold", text.Attribute("font-weight")?.Value);
        Assert.Equal("italic", text.Attribute("font-style")?.Value);
        var image = svg.Descendants().Single(element => element.Name.LocalName == "image");
        Assert.Contains("matrix(0 20 -20 0 20 0)", image.Attribute("transform")?.Value);
        Assert.Contains(svg.Descendants(), element => element.Name.LocalName == "clipPath");
        Assert.Contains(image.Ancestors(), element => element.Attribute("opacity")?.Value == "0.502");
    }

    [Fact]
    public async Task GraphicsEvenOddPath_ShouldExportSvgFillRule()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        var path = new OfdGraphicsPath(OfdFillRule.EvenOdd)
            .AddRectangle(5, 5, 50, 50).AddRectangle(15, 15, 20, 20);
        graphics.FillPath(new OfdBrush(OfdColor.Black), path);
        var svg = await ConvertAsync(package);
        var node = Assert.Single(svg.Descendants(), element => element.Name.LocalName == "path");
        Assert.Equal("evenodd", node.Attribute("fill-rule")?.Value);
    }

    [Fact]
    public async Task GraphicsAcuteMiter_ShouldExportMiterLimitTen()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        var path = new OfdGraphicsPath().MoveTo(30, 70).LineTo(40, 10).LineTo(50, 70);
        graphics.DrawPath(path, new OfdPen(OfdColor.Black, 1));
        var svg = await ConvertAsync(package);
        var node = Assert.Single(svg.Descendants(), element => element.Name.LocalName == "path");
        Assert.Equal("miter", node.Attribute("stroke-linejoin")?.Value);
        Assert.Equal("10", node.Attribute("stroke-miterlimit")?.Value);
    }

    [Fact]
    public async Task FractionalGraphicsScale_ShouldSurviveBothClipAndObjectSvgTransforms()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 1001, HeightMillimeters = 20 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        graphics.Scale(0.9996, 1);
        graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(0, 0, 1000, 10));
        graphics.DrawRectangle(new OfdPen(OfdColor.Black, 0.5), 0, 0, 1000, 10);
        var svg = await ConvertAsync(package);
        var paths = svg.Descendants().Where(node => node.Name.LocalName == "path").ToArray();
        Assert.Equal(2, paths.Length);
        Assert.All(paths, path => Assert.Contains("matrix(0.9996 0 0 1 ", path.Attribute("transform")?.Value));
        Assert.Equal(paths[0].Attribute("transform")?.Value, paths[1].Attribute("transform")?.Value);
    }

    [Fact]
    public async Task TinyAdvance_UnderLargeScale_ShouldRetainSecondGlyphSvgOrigin()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        graphics.Scale(1000, 1);
        graphics.DrawString("AB", new OfdFont("SimSun", 3.333), new OfdBrush(OfdColor.Black), 0.01, 10, new[] { 0.0004 });
        var svg = await ConvertAsync(package);
        var text = Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text");
        var glyphs = text.Elements().Where(node => node.Name.LocalName == "tspan").ToArray();
        Assert.Equal(2, glyphs.Length);
        Assert.Equal("0", glyphs[0].Attribute("x")?.Value);
        Assert.Equal("0.0004", glyphs[1].Attribute("x")?.Value);
        Assert.Contains("matrix(1000 0 0 1", text.Attribute("transform")?.Value);
    }

    private static async Task<XDocument> ConvertAsync(OfdDocumentPackage package)
    {
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var output = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(ofd, output);
        output.Position = 0;
        return XDocument.Load(output);
    }
}
