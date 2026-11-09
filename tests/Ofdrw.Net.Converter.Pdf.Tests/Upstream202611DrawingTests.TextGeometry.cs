using System.Globalization;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using PdfDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed partial class Upstream202611DrawingTests
{
    [Fact]
    public async Task TextCodeOriginsTransformPointsAndDeltasIndependently()
    {
        // For Boundary B, origin p and delta d, page origins are
        // B + A*p + t and B + A*(p+d) + t. The delta is A*d.
        const string body = """
            <TextObject ID='501' Boundary='10 10 42 35' Font='10' Size='6' CTM='1 0.25 0.5 1 100 200'>
              <TextCode X='2' Y='8' DeltaX='7' DeltaY='3'>AB</TextCode>
              <TextCode X='12' Y='21' DeltaX='8' DeltaY='2'>CD</TextCode>
            </TextObject>
            <TextObject ID='502' Boundary='60 10 42 35' Font='10' Size='6' CTM='0 1 -1 0 100 200'>
              <TextCode X='2' Y='8' DeltaX='7' DeltaY='3'>EF</TextCode>
            </TextObject>
            <TextObject ID='503' Boundary='90 10 42 35' Font='10' Size='6' CTM='1 0.25 0.5 1 100 200'>
              <CGTransform CodePosition='0' CodeCount='2' GlyphCount='2'><Glyphs>36 36</Glyphs></CGTransform>
              <TextCode X='2' Y='8' DeltaX='7' DeltaY='3'>AA</TextCode>
            </TextObject>
            """;
        var input = await TopLevelText(body);
        var render = await Render("text-delta", input);
        using var pdf = PdfDocument.Open(render.Pdf);
        var letters = pdf.GetPage(1).Letters;
        Assert.Equal("ABCDEFAA", string.Concat(letters.Select(letter => letter.Value)));
        (double X, double Y)[] origins =
        [
            (116, 218.5), (124.5, 223.25), (132.5, 234), (141.5, 238),
            (152, 212), (149, 219), (196, 218.5), (204.5, 223.25)
        ];
        const double mmPerPoint = 25.4 / 72;
        for (var i = 0; i < origins.Length; i++)
        {
            Assert.InRange(letters[i].StartBaseLine.X * mmPerPoint, origins[i].X - 0.03, origins[i].X + 0.03);
            Assert.InRange(250 - letters[i].StartBaseLine.Y * mmPerPoint, origins[i].Y - 0.03, origins[i].Y + 0.03);
        }

        foreach (var svg in new[] { render.Svg, render.Html })
        {
            var texts = Nodes(svg, "text").ToArray();
            Assert.Equal(new[] { "AB", "CD", "EF" }, texts.Select(node => node.Value));
            Assert.Equal(new[]
            {
                "translate(10 10) matrix(1 0.25 0.5 1 100 200)",
                "translate(10 10) matrix(1 0.25 0.5 1 100 200)",
                "translate(60 10) matrix(0 1 -1 0 100 200)"
            }, texts.Select(node => node.Attribute("transform")!.Value));
            var spans = Nodes(svg, "tspan");
            Assert.Equal(new (double, double)[] { (2, 8), (9, 11), (12, 21), (20, 23), (2, 8), (9, 11) },
                spans.Select(node => (Number(node, "x"), Number(node, "y"))));

            var group = Assert.Single(Nodes(svg, "g"), node => node.Attribute("data-ofd-text")?.Value == "AA");
            Assert.Equal("translate(90 10) matrix(1 0.25 0.5 1 100 200)", group.Attribute("transform")!.Value);
            var paths = Nodes(group, "path").ToArray();
            Assert.Equal(new[] { "36", "36" }, paths.Select(node => node.Attribute("data-original-gid")!.Value));
            // Identical A outlines. The difference of their first Move points
            // must be the local delta (7,3); the group maps it to (8.5,4.75).
            var first = FirstMove(paths[0]);
            var second = FirstMove(paths[1]);
            Assert.Equal(7, second.X - first.X, 5);
            Assert.Equal(3, second.Y - first.Y, 5);
        }

        using var raster = Image.Load<Rgb24>(render.Png);
        for (var i = 0; i < origins.Length; i++)
        {
            var (x, y) = origins[i];
            var rotated = i is 4 or 5;
            Assert.True(HasInk(raster, x - 2, y - (rotated ? 1 : 5), x + 6, y + (rotated ? 5 : 2)),
                $"No ink for glyph {i} near ({x},{y}) mm");
        }
        Assert.False(HasInk(raster, 8, 10, 50, 45)); // original untransformed Boundary
        Assert.False(HasInk(raster, 175, 205, 187, 240)); // gap between rotation and explicit glyphs
    }

    [Fact]
    public async Task VerticalDirectionIsPreservedInXmlButRejectedByRenderers()
    {
        foreach (var direction in new[] { "ReadDirection", "CharDirection" })
        {
            var input = await TopLevelText($"<TextObject ID='501' Boundary='10 10 42 35' Font='10' Size='6' {direction}='90'>" +
                "<TextCode X='2' Y='8'>vertical</TextCode></TextObject>");
            RetainInput("vertical-" + direction, input);
            var package = await new OfdReader().ReadAsync(new MemoryStream(input));
            Assert.Equal("vertical", new OfdTextExtractor().Extract(package));
            using var output = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, output);
            output.Position = 0;
            var reparsed = await new OfdReader().ReadAsync(output);
            var element = Assert.IsType<OfdTextElement>(Assert.Single(reparsed.Pages.Single().Elements));
            Assert.Equal("90", XElement.Parse(element.SourceXml!).Attribute(direction)!.Value);
            await RejectInAllOutputs(input, "text direction");
        }
    }

    private static async Task<byte[]> TopLevelText(string body)
    {
        var seed = await R01PageBlockTests.Source(body, font: true, width: 230, height: 250);
        // Explicit glyph maps are only admitted on top-level text.
        return await R01PageBlockTests.RewriteXml(seed, PageXml, xml =>
        {
            var block = Assert.Single(Nodes(xml, "PageBlock"));
            block.ReplaceWith(block.Elements());
        });
    }

    private static (double X, double Y) FirstMove(XElement path)
    {
        var tokens = path.Attribute("d")!.Value.Split(' ');
        Assert.Equal("M", tokens[0]);
        return (double.Parse(tokens[1], CultureInfo.InvariantCulture),
            double.Parse(tokens[2], CultureInfo.InvariantCulture));
    }

    private static bool HasInk(Image<Rgb24> raster, double left, double top, double right, double bottom)
    {
        for (var y = (int)(top * PixelsPerMillimeter); y < bottom * PixelsPerMillimeter; y++)
        for (var x = (int)(left * PixelsPerMillimeter); x < right * PixelsPerMillimeter; x++)
        {
            var pixel = raster[x, y];
            if (pixel.R < 100 && pixel.G < 100 && pixel.B < 100) return true;
        }
        return false;
    }
}
