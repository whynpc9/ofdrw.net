using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using UglyToad.PdfPig.Graphics.Operations;
using UglyToad.PdfPig.Graphics.Operations.PathPainting;
using PdfDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed partial class Upstream202611DrawingTests
{
    private const string PageXml = "Doc_0/Pages/Page_0/Content.xml";

    [Fact]
    public async Task DrawParamPresenceAndRemovalPreserveLocalOverridesAndSiblingPaint()
    {
        const string resources = "<DrawParams><DrawParam ID='30' Cap='Square' Join='Bevel' LineWidth='5'/>" +
            "<DrawParam ID='31' Relative='30' Cap='Round' Join='Round' LineWidth='4'/><DrawParam ID='32' Relative='31'/></DrawParams>";
        var rows = new[] { "", "Cap='Butt' Join='Miter' LineWidth='0.353' Alpha='255'", "Cap='Square' Join='Bevel' LineWidth='2' Alpha='128'", "" };
        var body = string.Concat(rows.Select((attributes, i) =>
            $"<PathObject ID='{501 + i}' Boundary='10 {10 + i * 15} 40 10' DrawParam='32' {attributes}>" +
            "<AbbreviatedData>M 10 5 L 20 5 L 20 10</AbbreviatedData></PathObject>"));
        var input = await R01PageBlockTests.Source(body, resources: resources);

        RetainInput("drawparam-local", input);
        await Check(input, ["round", "butt", "square", "round"], ["round", "miter", "bevel", "round"], [4, 0.353, 2, 4], [255, 255, 128, 255]);
        var removed = await R01PageBlockTests.RewriteXml(input, PageXml, xml =>
        {
            var edited = Nodes(xml, "PathObject").ElementAt(1);
            foreach (var name in new[] { "Cap", "Join", "LineWidth", "Alpha" })
                edited.Attribute(name)!.Remove();
        });
        RetainInput("drawparam-removed", removed);
        await Check(removed, ["round", "round", "square", "round"], ["round", "round", "bevel", "round"], [4, 4, 2, 4], [255, 255, 128, 255]);

        static async Task Check(byte[] ofd, string[] caps, string[] joins, double[] widths, int[] alphas)
        {
            var paths = (await ReadBlock(ofd)).Cast<OfdPathElement>().ToArray();
            Assert.Equal(4, paths.Length);
            for (var i = 0; i < paths.Length; i++)
            {
                var style = OfdPathStyle.Read(paths[i]);
                Assert.Equal(caps[i], style.Cap);
                Assert.Equal(joins[i], style.Join);
                Assert.Equal(widths[i], paths[i].LineWidthMillimeters, 8);
                Assert.Equal(alphas[i] / 255d, style.Alpha, 8);
            }
            var local = XElement.Parse(paths[1].SourceXml!);
            Assert.Equal(caps[1] == "butt" ? "Butt" : null, local.Attribute("Cap")?.Value);
            Assert.Equal(caps[1] == "butt" ? "Miter" : null, local.Attribute("Join")?.Value);
            Assert.Equal(caps[1] == "butt" ? "0.353" : null, local.Attribute("LineWidth")?.Value);
            Assert.Equal(caps[1] == "butt" ? "255" : null, local.Attribute("Alpha")?.Value);
        }
    }

    [Fact]
    public async Task DrawParamAccepts32LinksAndRejects33OrCycle()
    {
        var chain = "<DrawParams><DrawParam ID='100' Cap='Round'/>";
        for (var id = 101; id <= 132; id++)
            chain += $"<DrawParam ID='{id}' Relative='{id - 1}'/>";
        chain += "<DrawParam ID='133' Relative='134'/><DrawParam ID='134' Relative='133'/></DrawParams>";
        const string path = "<PathObject ID='501' Boundary='5 5 40 10' DrawParam='{0}'>" +
            "<AbbreviatedData>M 2 5 L 30 5</AbbreviatedData></PathObject>";
        var valid = await R01PageBlockTests.Source(string.Format(path, 131), resources: chain);
        RetainInput("drawparam-32", valid);
        Assert.Equal("round", OfdPathStyle.Read(Assert.IsType<OfdPathElement>(Assert.Single(await ReadBlock(valid)))).Cap);
        foreach (var id in new[] { 132, 133 })
        {
            var invalid = await R01PageBlockTests.Source(string.Format(path, id), resources: chain);
            RetainInput(id == 132 ? "drawparam-33" : "drawparam-cycle", invalid);
            await RejectInAllOutputs(invalid, "inheritance cycle or depth limit");
        }
    }

    [Fact]
    public async Task BitsPerComponentDefaultsToEightAndRejectsOtherDepths()
    {
        const string body = "<PathObject ID='501' Boundary='5 5 20 20' Fill='true' Stroke='false'>" +
            "<FillColor ColorSpace='20' Value='17 91 201'/><AbbreviatedData>M 0 0 L 20 0 L 20 20 C</AbbreviatedData></PathObject>";
        var input = await R01PageBlockTests.Source(body,
            resources: "<ColorSpaces><ColorSpace ID='20' Type='RGB' BitsPerComponent='8'/></ColorSpaces>");
        RetainInput("bits-explicit-eight", input);
        await Check(input);
        var unsupported = await R01PageBlockTests.RewriteXml(input, "Doc_0/DocumentRes.xml",
            xml => Nodes(xml, "ColorSpace").Single().SetAttributeValue("BitsPerComponent", "4"));
        RetainInput("bits-unsupported-four", unsupported);
        await RejectInAllOutputs(unsupported, "ColorSpace");
        var removed = await R01PageBlockTests.RewriteXml(unsupported, "Doc_0/DocumentRes.xml",
            xml => Nodes(xml, "ColorSpace").Single().Attribute("BitsPerComponent")!.Remove());
        RetainInput("bits-absent", removed);
        await Check(removed);

        static async Task Check(byte[] ofd)
        {
            var path = Assert.IsType<OfdPathElement>(Assert.Single(await ReadBlock(ofd)));
            Assert.Null(path.RenderPaintError);
            Assert.Equal((17, 91, 201, 255), Channels(path.FillColor!));
        }
    }

    [Fact]
    public async Task ImportedStrokeFlagControlsPaintAndPreservesText()
    {
        const string red = "<StrokeColor Value='255 0 0'/>";
        const string blue = "<StrokeColor Value='0 0 255'/>";
        const string resources = "<DrawParams><DrawParam ID='30' LineWidth='0.26'>" + red + "</DrawParam></DrawParams>";
        var rows = (from flag in new[] { "", "false", "true" }
                    from color in new[] { "none", "local", "inherited" }
                    select (Flag: flag, Color: color, Label: $"{(flag == "" ? "absent" : flag)}-{color}")).ToArray();
        var body = string.Concat(rows.Select((row, i) =>
        {
            var stroke = row.Flag.Length == 0 ? "" : $"Stroke='{row.Flag}'";
            var draw = row.Color == "inherited" ? "DrawParam='30'" : "LineWidth='0.26'";
            return $"<TextObject ID='{501 + i}' Boundary='10 {5 + i * 10} 75 9' Font='10' Size='5' {stroke} {draw}>" +
                (row.Color == "local" ? blue : "") + $"<TextCode X='0' Y='6'>{row.Label}</TextCode></TextObject>";
        }));
        var producer = new OfdTextElement
        {
            ObjectId = "520", FontResourceId = "10", FontName = "Noto Sans", Text = "producer-bold",
            XMillimeters = 10, YMillimeters = 95, WidthMillimeters = 75, HeightMillimeters = 9, FontSizeMillimeters = 5,
            ProducerStrokeColor = OfdColor.Black, ProducerLineWidthMillimeters = 0.8,
            Runs = { new OfdTextRun { Text = "producer-bold", XMillimeters = 0, YMillimeters = 6 } }
        };
        Assert.True(OfdTextPaint.Read(producer).Stroke);
        var input = await R01PageBlockTests.Source(body, resources: resources, font: true, height: 110,
            producerText: producer);
        var labels = rows.Select(row => row.Label).Append(producer.Text).ToArray();
        var package = await new OfdReader().ReadAsync(new MemoryStream(input));
        var imported = (await ReadBlock(input)).Cast<OfdTextElement>().ToArray();
        Assert.Equal(rows.Select(row => row.Label), imported.Select(text => text.Text));
        for (var i = 0; i < rows.Length; i++)
        {
            var paint = OfdTextPaint.Read(imported[i]);
            Assert.Equal(rows[i].Flag == "true", paint.Stroke);
            Assert.Equal(rows[i].Color switch
            {
                "local" => (0, 0, 255, 255),
                "inherited" => (255, 0, 0, 255),
                _ => (0, 0, 0, 255)
            }, Channels(paint.StrokeColor));
        }
        Assert.Equal(string.Concat(labels), string.Concat(new OfdTextExtractor().Extract(package)
            .Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)));
        var render = await Render("text-stroke", input);
        foreach (var svg in new[] { render.Svg, render.Html })
        {
            var text = Nodes(svg, "text").ToArray();
            Assert.Equal(labels, text.Select(node => node.Value));
            Assert.Equal(Enumerable.Repeat("none", 6).Concat(new[] { "rgb(0,0,0)", "rgb(0,0,255)", "rgb(255,0,0)", "rgb(0,0,0)" }),
                text.Select(node => node.Attribute("stroke")!.Value));
        }

        using var pdf = PdfDocument.Open(render.Pdf);
        var page = pdf.GetPage(1);
        Assert.Equal(string.Concat(labels), page.Text);
        // The PDF emits outline paths alongside one searchable text run per
        // object. Only the two colored strokes may set a non-default RGB pen.
        Assert.Contains(page.Operations, operation => operation is StrokePath);
        Assert.Equal(new (double, double, double)[] { (0, 0, 1), (1, 0, 0) },
            page.Operations.OfType<SetStrokeColorDeviceRgb>().Select(color => (color.R, color.G, color.B)));
        using var raster = Image.Load<Rgb24>(render.Png);
        for (var i = 0; i < labels.Length; i++)
        {
            var (dark, redInk, blueInk) = TextInk(raster, 5 + i * 10);
            Assert.True(dark > 25, $"{labels[i]} has no dark fill");
            Assert.Equal(i == 8, redInk > 0);
            Assert.Equal(i == 7, blueInk > 0);
        }
    }

    private static (int Red, int Green, int Blue, int Alpha) Channels(OfdColor color) =>
        (color.Red, color.Green, color.Blue, color.Alpha);

    private static (int Dark, int Red, int Blue) TextInk(Image<Rgb24> raster, int top)
    {
        var dark = 0;
        var red = 0;
        var blue = 0;
        for (var y = top * PixelsPerMillimeter; y < (top + 9) * PixelsPerMillimeter; y++)
        for (var x = 9 * PixelsPerMillimeter; x < 85 * PixelsPerMillimeter; x++)
        {
            var pixel = raster[x, y];
            if (pixel.R < 100 && pixel.G < 100 && pixel.B < 100) dark++;
            if (pixel.R > 150 && pixel.G < 80 && pixel.B < 80) red++;
            if (pixel.B > 150 && pixel.G < 80 && pixel.R < 80) blue++;
        }
        return (dark, red, blue);
    }
}
