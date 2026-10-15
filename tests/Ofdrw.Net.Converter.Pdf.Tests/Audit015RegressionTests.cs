using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.Advanced;
using PdfSharpCore.Pdf.IO;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class Audit015RegressionTests
{
    [Theory]
    [InlineData(0, 255)] [InlineData(128, 128)] [InlineData(255, 128)] [InlineData(255, 255)] [InlineData(255, 0)] [InlineData(128, 0)]
    public async Task Audit015_ObjectAlpha(int alpha, int colorAlpha)
    {
        var bytes = await Fixture($"Alpha='{alpha}'", colorAlpha: colorAlpha);
        var (svg, pdf) = await Export(bytes);
        var node = PaintedPath(svg);
        var expected = alpha / 255d * colorAlpha / 255d;
        Assert.Equal(expected, Opacity(node, "stroke-opacity"), 4);
        Assert.Equal(expected, Opacity(node, "fill-opacity"), 4);
        using var parsed = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
        var states = parsed.Pages[0].Resources.Elements.GetDictionary("/ExtGState");
        var dictionaries = states?.Elements.Values.Select(v => v is PdfReference reference ? reference.Value : v).OfType<PdfDictionary>().ToArray() ?? [];
        if (expected == 0)
        {
            using var pig = UglyToad.PdfPig.PdfDocument.Open(pdf);
            Assert.DoesNotContain(pig.GetPage(1).Operations, op => op.Operator is "B" or "B*" or "S" or "f" or "f*");
        }
        else if (expected != 1)
        {
            Assert.Contains(dictionaries, d => Math.Abs(d.Elements.GetReal("/CA") - expected) < 0.001);
            Assert.Contains(dictionaries, d => Math.Abs(d.Elements.GetReal("/ca") - expected) < 0.001);
        }
    }

    [Theory]
    [InlineData("6 2", "6 2")] [InlineData("6 2 1", "6 2 1 6 2 1")]
    public async Task Audit015_DashOffset(string pattern, string effective)
    {
        var (svg, pdf) = await Export(await Fixture($"DashPattern='{pattern}' DashOffset='1.5'"));
        var node = PaintedPath(svg);
        Assert.Equal(effective, node.Attribute("stroke-dasharray")?.Value);
        Assert.Equal("1.5", node.Attribute("stroke-dashoffset")?.Value);
        var content = Content(pdf);
        var match = System.Text.RegularExpressions.Regex.Match(content, @"\[([^\]]+)\]([\d.]+) d");
        Assert.True(match.Success, content);
        var actual = match.Groups[1].Value.Split(' ', StringSplitOptions.RemoveEmptyEntries).Select(Number).ToArray();
        Assert.Equal(effective.Split(' ').Length, actual.Length);
        Assert.All(actual.Zip(effective.Split(' ').Select(Number)), pair => Assert.InRange(Math.Abs(pair.Second * 72 / 25.4 - pair.First), 0, 0.0006));
        Assert.Equal(1.5 * 72 / 25.4, Number(match.Groups[2].Value), 2);
    }

    [Theory]
    [InlineData("Butt", "Miter", 0, 0)] [InlineData("Round", "Round", 1, 1)] [InlineData("Square", "Bevel", 2, 2)]
    public async Task Audit015_CapsAndJoins(string cap, string join, int pdfCap, int pdfJoin)
    {
        var (svg, pdf) = await Export(await Fixture($"Cap='{cap}' Join='{join}'"));
        var node = PaintedPath(svg);
        Assert.Equal(cap.ToLowerInvariant(), node.Attribute("stroke-linecap")?.Value);
        Assert.Equal(join.ToLowerInvariant(), node.Attribute("stroke-linejoin")?.Value);
        Assert.Contains($"{pdfCap} J", Content(pdf)); Assert.Contains($"{pdfJoin} j", Content(pdf));
    }

    [Fact]
    public async Task Audit015_MiterLimit()
    {
        var (svg, pdf) = await Export(await Fixture("Cap='Butt' Join='Miter' MiterLimit='7'"));
        Assert.Equal("7", PaintedPath(svg).Attribute("stroke-miterlimit")?.Value);
        Assert.Contains("7 M", Content(pdf));
    }

    [Theory]
    [InlineData("Alpha='NaN'")] [InlineData("Alpha='256'")] [InlineData("Alpha='-1'")]
    [InlineData("DashPattern='1 -1'")] [InlineData("DashPattern='NaN 2'")] [InlineData("DashPattern='0 0'")]
    [InlineData("DashPattern=''")] [InlineData("DashOffset='Infinity'")] [InlineData("DashOffset='-1'")]
    [InlineData("MiterLimit='NaN'")] [InlineData("MiterLimit='0.5'")] [InlineData("LineWidth='NaN'")]
    [InlineData("LineWidth='-1'")] [InlineData("DashPattern='1e308 1e308'")]
    [InlineData("Cap='Triangle'")] [InlineData("Join='Triangles'")]
    public async Task Audit015_InvalidNumericStyle(string attributes)
    {
        var bytes = await Fixture(attributes);
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
    }

    [Fact]
    public async Task Audit015_FreshDefaults()
    {
        var (svg, pdf) = await Export(await Fixture(null));
        var node = PaintedPath(svg);
        Assert.Null(node.Attribute("stroke-dasharray")); Assert.Equal(1, Opacity(node, "stroke-opacity"));
        Assert.Contains("0 J", Content(pdf)); Assert.Contains("0 j", Content(pdf));
        Assert.Contains("[]0 d", Content(pdf));
    }

    [Fact]
    public async Task Audit015_SourceOwnershipSurvivesExport()
    {
        var bytes = await Fixture("Alpha='128' DashPattern='6 2' DashOffset='1.5' Cap='Round'");
        var package = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(new MemoryStream(bytes));
        var source = Assert.IsType<OfdPathElement>(package.Pages[0].Elements[0]).SourceXml;
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, saved);
        await Export(saved.ToArray());
        Assert.Equal(source, Assert.IsType<OfdPathElement>(package.Pages[0].Elements[0]).SourceXml);
        using var zip = new ZipArchive(new MemoryStream(saved.ToArray()));
        using var stream = zip.GetEntry("Doc_0/Pages/Page_0/Content.xml")!.Open();
        Assert.Equal("6 2", XDocument.Load(stream).Descendants().Single(n => n.Name.LocalName == "PathObject").Attribute("DashPattern")?.Value);
    }

    [Theory]
    [InlineData("DashPattern='0 2'", "stroke-dasharray", "0 2")]
    [InlineData("MiterLimit='4.5'", "stroke-miterlimit", "4.5")]
    [InlineData("Cap='Round' MiterLimit='7'", "stroke-miterlimit", "7")]
    [InlineData("Cap='Square' MiterLimit='7'", "stroke-miterlimit", "7")]
    public async Task Audit015_PdfBackendLimitationsAreExplicit(string attributes, string svgAttribute, string value)
    {
        var bytes = await Fixture(attributes);
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        Assert.Equal(value, PaintedPath(XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray()))).Attribute(svgAttribute)?.Value);
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), new MemoryStream()));
    }

    [Fact]
    public void Audit015_InstalledPdfBackendMiterAndOddDashLimitations()
    {
        using var document = new PdfDocument(); var page = document.AddPage();
        using (var graphics = PdfSharpCore.Drawing.XGraphics.FromPdfPage(page))
        {
            var pen = new PdfSharpCore.Drawing.XPen(PdfSharpCore.Drawing.XColors.Blue, 2)
            { MiterLimit = 4.5, DashPattern = [3], LineCap = PdfSharpCore.Drawing.XLineCap.Flat };
            graphics.DrawLine(pen, 0, 0, 100, 100);
            pen.LineCap = PdfSharpCore.Drawing.XLineCap.Round; pen.MiterLimit = 7;
            graphics.DrawLine(pen, 0, 10, 100, 110);
            Assert.Throws<ArgumentException>(() => pen.DashPattern = [0, 1]);
        }
        using var output = new MemoryStream(); document.Save(output, false);
        var content = Content(output.ToArray());
        Assert.Contains("4 M", content); Assert.DoesNotContain("4.5 M", content);
        Assert.DoesNotContain("7 M", content); Assert.Contains("[6 0.4]0 d", content);
    }

    [Fact]
    public async Task Audit015_NestedRepeatedOccurrenceKeepsStyleAndSourceOwnership()
    {
        var package = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(new MemoryStream(await Fixture("Alpha='128' DashPattern='6 2' DashOffset='1.5' Cap='Round' Join='Round'")));
        var original = Assert.IsType<OfdPathElement>(package.Pages[0].Elements[0]); var source = original.SourceXml;
        var leaf = new OfdCompositeResource { WidthMillimeters = 100, HeightMillimeters = 70 }; leaf.Elements.Add(original);
        var parent = new OfdCompositeResource { WidthMillimeters = 100, HeightMillimeters = 70 };
        parent.Elements.Add(new OfdCompositeElement { Resource = leaf, WidthMillimeters = 100, HeightMillimeters = 70 });
        package.CompositeResources = new(); package.CompositeResources.Items.Add(leaf); package.CompositeResources.Items.Add(parent);
        package.Pages[0].WidthMillimeters = 220; package.Pages[0].HeightMillimeters = 90; package.Pages[0].Elements.Clear();
        for (var i = 0; i < 2; i++) package.Pages[0].Elements.Add(new OfdCompositeElement { Resource = parent, XMillimeters = i * 105, WidthMillimeters = 100, HeightMillimeters = 70 });
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd);
        var (svg, pdf) = await Export(ofd.ToArray());
        var paths = svg.Descendants().Where(n => n.Name.LocalName == "path" && n.Attribute("stroke-dasharray") is not null).ToArray();
        Assert.Equal(2, paths.Length); Assert.All(paths, p => { Assert.Equal("6 2", p.Attribute("stroke-dasharray")?.Value); Assert.Equal(128 / 255d, Opacity(p, "stroke-opacity"), 6); Assert.Equal("round", p.Attribute("stroke-linecap")?.Value); });
        using var pig = UglyToad.PdfPig.PdfDocument.Open(pdf); Assert.Equal(2, pig.GetPage(1).Operations.Count(op => op.Operator == "B"));
        Assert.Equal(source, original.SourceXml);
        await SaveArtifacts("nested-repeated", ofd.ToArray(), svg, pdf);
    }

    [Fact]
    public async Task Audit015_VisualControls()
    {
        foreach (var (name, attributes) in new[] {
            ("fresh", (string?)null), ("alpha-zero", "Alpha='0'"), ("alpha-half", "Alpha='128'"),
            ("dash-offset", "DashPattern='6 2' DashOffset='1.5'"), ("odd-dash", "DashPattern='6 2 1' DashOffset='1.5'"),
            ("round", "Cap='Round' Join='Round'"), ("square-bevel", "Cap='Square' Join='Bevel'"), ("miter-seven", "Cap='Butt' Join='Miter' MiterLimit='7'") })
        {
            var bytes = await Fixture(attributes); var (svg, pdf) = await Export(bytes);
            await SaveArtifacts(name, bytes, svg, pdf);
            using var pig = UglyToad.PdfPig.PdfDocument.Open(pdf); Assert.Equal(1, pig.NumberOfPages);
        }
    }

    private static async Task SaveArtifacts(string name, byte[] ofd, XDocument svg, byte[] pdf)
    {
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is not { Length: > 0 } directory) return;
        Directory.CreateDirectory(directory); await File.WriteAllBytesAsync(Path.Combine(directory, name + ".ofd"), ofd);
        svg.Save(Path.Combine(directory, name + ".svg")); await File.WriteAllBytesAsync(Path.Combine(directory, name + ".pdf"), pdf);
    }

    internal static async Task<byte[]> Fixture(string? attributes, int colorAlpha = 255)
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 70 };
        page.Elements.Add(new OfdPathElement { XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 50,
            AbbreviatedData = "M 5 25 L 35 5 L 65 25 L 65 40 L 5 40 C M 5 47 L 65 47 M 68 40 L 73 3 L 78 40", Fill = true, Stroke = true, LineWidthMillimeters = 2,
            StrokeColor = new OfdColor(0, 40, 180, colorAlpha), FillColor = new OfdColor(220, 40, 20, colorAlpha) });
        package.Pages.Add(page); using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd);
        if (attributes is not null)
        {
            using var zip = new ZipArchive(ofd, ZipArchiveMode.Update, true); var entry = zip.GetEntry("Doc_0/Pages/Page_0/Content.xml")!;
            XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            var node = xml.Descendants().Single(n => n.Name.LocalName == "PathObject");
            foreach (var attribute in XElement.Parse($"<style {attributes}/>").Attributes()) node.SetAttributeValue(attribute.Name, attribute.Value);
            entry.Delete(); using var output = zip.CreateEntry("Doc_0/Pages/Page_0/Content.xml").Open(); xml.Save(output);
        }
        return ofd.ToArray();
    }

    internal static async Task<(XDocument Svg, byte[] Pdf)> Export(byte[] bytes)
    {
        using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(new MemoryStream(bytes), svg);
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(new MemoryStream(bytes), pdf);
        return (XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())), pdf.ToArray());
    }
    private static XElement PaintedPath(XDocument document) => document.Descendants().Single(n => n.Name.LocalName == "path");
    private static double Number(string value) => double.Parse(value, CultureInfo.InvariantCulture);
    private static double Opacity(XElement node, string name) => Number(node.Attribute(name)?.Value ?? "1");
    private static string Content(byte[] bytes)
    {
        using var parsed = PdfReader.Open(new MemoryStream(bytes), PdfDocumentOpenMode.Import);
        return Encoding.ASCII.GetString(parsed.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
    }
}
