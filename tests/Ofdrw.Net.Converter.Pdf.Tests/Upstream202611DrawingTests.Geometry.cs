using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed partial class Upstream202611DrawingTests
{
    private static string Rect(int left, int top, int right, int bottom) =>
        $"M {left} {top} L {right} {top} L {right} {bottom} L {left} {bottom} C";

    private static string Clip(int left, int top, int right, int bottom) =>
        "<Clip><Area><Path Stroke='false' Fill='true'><AbbreviatedData>" +
        Rect(left, top, right, bottom) + "</AbbreviatedData></Path></Area></Clip>";

    private static string Box(int id, string boundary, string data, string color, string attributes = "",
        string clips = "", int colorAlpha = 255) =>
        $"<PathObject ID='{id}' Boundary='{boundary}' Fill='true' Stroke='false' {attributes}>{clips}" +
        $"<FillColor Value='{color}' Alpha='{colorAlpha}'/><AbbreviatedData>{data}</AbbreviatedData></PathObject>";

    [Fact]
    public async Task NestedClipIntersectionAndAlphaDoNotLeakToSiblings()
    {
        var first = Box(550, "0 0 24 30", Rect(0, 0, 24, 30), "255 0 0", "Alpha='255'",
            $"<Clips>{Clip(4, 6, 18, 25)}{Clip(8, 2, 22, 16)}</Clips>");
        var second = Box(551, "30 0 30 30", Rect(0, 0, 30, 30), "0 0 255", "Alpha='128'", colorAlpha: 128);
        var transparent = Box(552, "0 40 20 16", Rect(0, 0, 20, 16), "0 0 0", "Alpha='0'");
        var sibling = Box(553, "70 10 18 24", Rect(0, 0, 18, 24), "0 160 0");
        var resources = "<CompositeGraphicUnits><CompositeGraphicUnit ID='30' Width='60' Height='30'>" +
            "<CompositeObject ID='540' Boundary='0 0 60 30' ResourceID='31'/></CompositeGraphicUnit>" +
            $"<CompositeGraphicUnit ID='31' Width='60' Height='30'>{first}{second}</CompositeGraphicUnit></CompositeGraphicUnits>";
        var placement = "<CompositeObject ID='501' Boundary='10 10 60 30' ResourceID='30'>" +
            $"<Clips>{Clip(0, 0, 54, 24)}</Clips></CompositeObject>";
        var input = await R01PageBlockTests.Source(placement + sibling + transparent, resources: resources, height: 65);
        var render = await Render("clip-scope-alpha", input);

        foreach (var svg in new[] { render.Svg, render.Html })
        {
            var paint = Nodes(svg, "path").Where(node => node.Attribute("fill") is not null).ToArray();
            Assert.Equal(new[] { "rgb(255,0,0)", "rgb(0,0,255)", "rgb(0,160,0)", "rgb(0,0,0)" },
                paint.Select(node => node.Attribute("fill")!.Value));
            Assert.True(ClipAncestors(paint[0]) >= ClipAncestors(paint[1]) + 2);
            Assert.True(ClipAncestors(paint[1]) > 0);
            Assert.Equal(0, ClipAncestors(paint[2]));
            Assert.Equal(Math.Pow(128 / 255d, 2), Number(paint[1], "fill-opacity"), 6);
            Assert.Equal(0, Number(paint[3], "fill-opacity"));
        }

        using var raster = Image.Load<Rgb24>(render.Png);
        var white = new Rgb24(255, 255, 255);
        // Red intersection is page x=[18,28], y=[16,26]. These two
        // white probes are in only one of the two component clips.
        Pixel(raster, 23, 20, new Rgb24(255, 0, 0));
        Pixel(raster, 16, 20, white);
        Pixel(raster, 23, 30, white);
        // Blue survives the first object's clips, but not the parent's.
        // Over white, blue with (128/255)^2 opacity yields (191,191,255).
        Pixel(raster, 52, 20, new Rgb24(191, 191, 255));
        Pixel(raster, 67, 20, white);
        Pixel(raster, 52, 37, white);
        Pixel(raster, 78, 20, new Rgb24(0, 160, 0));
        Pixel(raster, 10, 46, white);

        static int ClipAncestors(XElement node) => node.Ancestors().Count(parent => parent.Attribute("clip-path") is not null);
    }

    [Fact]
    public async Task ImageQuadrantsAndPathUseAffineTransformOnce()
    {
        const string resources = "<MultiMedias><MultiMedia ID='30' Type='Image' Format='PNG'>" +
            "<MediaFile>image.png</MediaFile></MultiMedia></MultiMedias>";
        // Unit-square corners TL,TR,BL,BR after Boundary + CTM:
        // 1: (8,10),(32,10),(8,26),(32,26) despite whole-page Boundary.
        // 2: (61,10),(61,34),(45,10),(45,34) after a clockwise turn.
        // 3: (70,10),(94,14),(78,26),(102,30) after a shear.
        // 4: (34,48),(10,48),(34,64),(10,64) after reflection.
        const string images = """
            <ImageObject ID='501' ResourceID='30' Boundary='0 0 110 80' CTM='24 0 0 16 8 10'/>
            <ImageObject ID='502' ResourceID='30' Boundary='40 5 26 34' CTM='0 24 -16 0 21 5'/>
            <ImageObject ID='503' ResourceID='30' Boundary='66 5 40 30' CTM='24 4 8 16 4 5'/>
            <ImageObject ID='504' ResourceID='30' Boundary='5 45 32 24' CTM='-24 0 0 16 29 3'/>
            """;
        // B=(50,48), t=(-50,-48), A=I. The absolute coordinates in
        // the path are intentionally left at (56,52)–(69,66).
        var path = Box(505, "50 48 30 24", Rect(56, 52, 69, 66), "120 0 180", "CTM='1 0 0 1 -50 -48'");
        var seed = await R01PageBlockTests.Source(images + path, resources: resources, width: 110, height: 80);
        var input = await R01PageBlockTests.WithEntry(seed, "Doc_0/image.png", Quadrants());
        var render = await Render("image-path-ctm", input);
        foreach (var svg in new[] { render.Svg, render.Html })
        {
            var nodes = Nodes(svg, "image").ToArray();
            Assert.Equal(4, nodes.Length);
            Assert.Equal(new[] { "24 0 0 16 8 10", "0 24 -16 0 21 5", "24 4 8 16 4 5", "-24 0 0 16 29 3" }
                .Select(ctm => $"translate(0 0) matrix({ctm})"), nodes.Select(node => node.Attribute("transform")!.Value));
            Assert.Equal(new[] { "translate(0 0)", "translate(40 5)", "translate(66 5)", "translate(5 45)" },
                nodes.Select(node => node.Ancestors().Last(parent => parent.Name.LocalName == "g").Attribute("transform")!.Value));
            Assert.All(nodes, node =>
            {
                Assert.Equal("1", node.Attribute("width")!.Value);
                Assert.Equal("1", node.Attribute("height")!.Value);
            });
            var shape = Assert.Single(Nodes(svg, "path"));
            Assert.Equal("translate(50 48) matrix(1 0 0 1 -50 -48)", shape.Attribute("transform")!.Value);
            Assert.Equal("M 56 52 L 69 52 L 69 66 L 56 66 Z", shape.Attribute("d")!.Value);
        }

        using var raster = Image.Load<Rgb24>(render.Png);
        // Images sampled at (u,v) = (1/4,1/4), (3/4,1/4), (1/4,3/4), (3/4,3/4).
        (double X, double Y)[][] samples =
        [
            [(14, 14), (26, 14), (14, 22), (26, 22)],
            [(57, 16), (57, 28), (49, 16), (49, 28)],
            [(78, 15), (90, 17), (82, 23), (94, 25)],
            [(28, 52), (16, 52), (28, 60), (16, 60)]
        ];
        Rgb24[] colors = [new(255, 0, 0), new(0, 255, 0), new(0, 0, 255), new(255, 255, 0)];
        foreach (var quadrants in samples)
        for (var i = 0; i < colors.Length; i++)
            Pixel(raster, quadrants[i].X, quadrants[i].Y, colors[i]);

        var white = new Rgb24(255, 255, 255);
        Pixel(raster, 34, 20, white);
        Pixel(raster, 42, 20, white);
        Pixel(raster, 71, 23, white);
        Pixel(raster, 35.5, 56, white);
        Pixel(raster, 62, 59, new Rgb24(120, 0, 180));
        Pixel(raster, 53, 59, white);
        Pixel(raster, 72, 59, white);
    }

    private static byte[] Quadrants()
    {
        using var image = new Image<Rgb24>(32, 16);
        for (var y = 0; y < image.Height; y++)
        for (var x = 0; x < image.Width; x++)
            image[x, y] = (x < 16, y < 8) switch
            {
                (true, true) => new Rgb24(255, 0, 0),
                (false, true) => new Rgb24(0, 255, 0),
                (true, false) => new Rgb24(0, 0, 255),
                _ => new Rgb24(255, 255, 0)
            };
        using var png = new MemoryStream();
        image.SaveAsPng(png);
        return png.ToArray();
    }
}
