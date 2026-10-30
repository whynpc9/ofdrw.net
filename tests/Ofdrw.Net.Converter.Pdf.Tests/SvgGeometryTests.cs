using OfdFillRule = Ofdrw.Net.Layout.Graphics.OfdFillRule;
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
    [Theory]
    [InlineData("cjk", "中", 4, .04, "normal")]
    [InlineData("latin", "A", 4, .1, "normal")]
    [InlineData("rounded-cjk", "中", 3.704, .037, "normal")]
    [InlineData("thin-cjk", "中", 4, .02, "bold")]
    [InlineData("thin-latin", "A", 4, .04, "bold")]
    [InlineData("hairline", "中", 4, 0, "bold")]
    [InlineData("no-stroke", "中", 4, .04, "bold")]
    [InlineData("different-color", "中", 4, .04, "bold")]
    [InlineData("fill-alpha", "中", 4, .04, "bold")]
    [InlineData("stroke-alpha", "中", 4, .04, "bold")]
    [InlineData("outline-only", "中", 4, .04, "bold")]
    [InlineData("name-only", "中", 4, .04, "bold")]
    [InlineData("physical-bold", "中", 4, .04, "bold")]
    public async Task DeclaredOpaqueStrokeOnlyReplacesAdditionalSyntheticBoldForAKnownUnboldFace(
        string kind, string value, double size, double width, string expectedWeight)
    {
        var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "narrow.ttf"));
        if (kind == "physical-bold")
        {
            var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(bytes);
            Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(face.Table("OS/2"), 62, 0x20);
            Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(face.Table("OS/2"), 4, 700);
            Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(face.Table("head"), 44, 1);
            bytes = face.Build(); // physical style metadata, not a visual700 fixture
        }
        var fill = kind != "outline-only"; var stroke = kind != "no-stroke";
        var fillAlpha = kind == "fill-alpha" ? 128 : 255;
        var strokeAlpha = kind == "stroke-alpha" ? 128 : 255;
        var strokeColor = kind == "different-color" ? "190 20 30" : "0 0 0";
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource
        {
            Id = "10", FontName = "fixture", Data = kind == "name-only" ? [] : bytes
        });
        package.Pages.Add(new OfdPage
        {
            WidthMillimeters = 100, HeightMillimeters = 100,
            Elements = { new OfdTextElement
            {
                FontName = "fixture", FontResourceId = "10", Text = value, Weight = 700,
                FontSizeMillimeters = size, XMillimeters = 10, YMillimeters = 10,
                WidthMillimeters = 40, HeightMillimeters = 15,
                FillColor = new OfdColor(0, 0, 0, (byte)fillAlpha),
                SourceXml = $"<TextObject xmlns='http://www.ofdspec.org' Boundary='10 10 40 15' Font='10' " +
                    $"Size='{size.ToString(CultureInfo.InvariantCulture)}' Weight='700' Fill='{fill.ToString().ToLowerInvariant()}' " +
                    $"Stroke='{stroke.ToString().ToLowerInvariant()}' LineWidth='{width.ToString(CultureInfo.InvariantCulture)}'>" +
                    $"<FillColor Value='0 0 0' Alpha='{fillAlpha}'/><StrokeColor Value='{strokeColor}' Alpha='{strokeAlpha}'/>" +
                    $"<TextCode X='0' Y='{size.ToString(CultureInfo.InvariantCulture)}'>{value}</TextCode></TextObject>"
            } }
        });
        var svg = await ConvertAsync(package);
        var text = Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text");
        Assert.Equal(expectedWeight, text.Attribute("font-weight")?.Value);
        Assert.Equal(value, text.Value);
        // Fresh name-only bold retains the writer's4.5%-em compatibility
        // stroke. SVG must not reinterpret that unknown physical face.
        if (stroke) Assert.Equal(kind == "name-only" ? size * .045 : width,
            double.Parse(text.Attribute("stroke-width")!.Value, CultureInfo.InvariantCulture));
        else Assert.Null(text.Attribute("stroke-width"));
    }

    [Theory]
    [InlineData("Fixture Family", "Fixture Family")]
    [InlineData(null, "Fixture Resource|bold|italic")]
    public async Task NameOnlyText_ShouldUseDeclaredFamilyWithoutLeakingStyleResourceKey(string? family, string expected)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource
        {
            Id = "10", FontName = "Fixture Resource|bold|italic", FamilyName = family,
            Bold = true, Italic = true
        });
        package.Pages.Add(new OfdPage
        {
            WidthMillimeters = 100, HeightMillimeters = 100,
            Elements = { new OfdTextElement
            {
                FontResourceId = "10", FontName = "Fixture Resource|bold|italic", Text = "Ab",
                XMillimeters = 10, YMillimeters = 10
            } }
        });

        var svg = await ConvertAsync(package);
        var text = Assert.Single(svg.Descendants(), node => node.Name.LocalName == "text");
        Assert.Equal(expected, text.Attribute("font-family")?.Value);
        Assert.Equal("bold", text.Attribute("font-weight")?.Value);
        Assert.Equal("italic", text.Attribute("font-style")?.Value);
        Assert.Equal("Ab", text.Value);
    }

    [Theory]
    [InlineData(1, false)]
    [InlineData(128, false)]
    [InlineData(254, false)]
    [InlineData(1, true)]
    [InlineData(128, true)]
    [InlineData(254, true)]
    public async Task BlankText_ShouldStayNonpaintingWithPartialFillAlpha(byte alpha, bool hasRun)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "fixture",
            Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "style-metrics.ttf")) });
        var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
        package.Pages.Add(page);
        foreach (var value in new[] { "\u202F", "A" })
        {
            var text = new OfdTextElement { Text = value, FontName = "fixture", FontResourceId = "10",
                FontSizeMillimeters = 8, XMillimeters = value == "A" ? 40 : 15,
                YMillimeters = 20, WidthMillimeters = 20, HeightMillimeters = 15,
                FillColor = new OfdColor(0, 0, 0, alpha) };
            if (hasRun) text.Runs.Add(new OfdTextRun { Text = value, YMillimeters = 8 });
            page.Elements.Add(text);
        }
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        // A valid external OFD may explicitly preserve a whitespace-only TextCode.
        // Exercise that input through the public archive reader and SVG exporter.
        using (var archive = new System.IO.Compression.ZipArchive(ofd, System.IO.Compression.ZipArchiveMode.Update, true))
        {
            var entry = archive.GetEntry("Doc_0/Pages/Page_0/Content.xml")!;
            XDocument content;
            using (var input = entry.Open()) content = XDocument.Load(input, LoadOptions.PreserveWhitespace);
            foreach (var code in content.Descendants().Where(node => node.Name.LocalName == "TextCode"))
                code.SetAttributeValue(XNamespace.Xml + "space", "preserve");
            entry.Delete();
            using var replacement = archive.CreateEntry("Doc_0/Pages/Page_0/Content.xml").Open();
            content.Save(replacement);
        }
        ofd.Position = 0;
        using var output = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(ofd, output);
        output.Position = 0;
        var svg = XDocument.Load(output);
        if (alpha == 128 && hasRun && Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            svg.Save(Path.Combine(directory, "blank-run-alpha.svg"));
            await File.WriteAllBytesAsync(Path.Combine(directory, "blank-run-alpha.ofd"), ofd.ToArray());
            ofd.Position = 0;
            using var pdf = new MemoryStream();
            await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(ofd, pdf);
            await File.WriteAllBytesAsync(Path.Combine(directory, "blank-run-alpha.pdf"), pdf.ToArray());
        }
        var nodes = svg.Descendants().Where(node => node.Name.LocalName == "text").ToArray();
        Assert.Equal("0", Assert.Single(nodes, node => node.Value == "\u202F").Attribute("fill-opacity")?.Value);
        Assert.Equal((alpha / 255d).ToString("0.###", CultureInfo.InvariantCulture),
            Assert.Single(nodes, node => node.Value == "A").Attribute("fill-opacity")?.Value);
    }

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
