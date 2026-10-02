using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class GraphicsEmphasisAndNumberTests
{
    private static readonly XName Marker = XName.Get("FauxItalicMatrixV1", "https://ofdrw.net/style-hints");
    private static (OfdDocumentPackage Package, OfdGraphics Graphics) Create()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        return (package, new OfdGraphics(package, page));
    }

    [Fact]
    public async Task PathAndAdvances_WritePlainDecimalLiteralsAcrossMagnitudeAndSign()
    {
        var (package, graphics) = Create();
        var path = new OfdGraphicsPath().MoveTo(1e-20, -1e20).LineTo(2e-20, -1e20 + 1e10);
        graphics.DrawPath(path, new OfdPen(OfdColor.Black));
        var deltas = new[] { 1e-20, -1e-20, double.Epsilon, 1e20 };
        graphics.DrawString("ABCDE", new OfdFont("Plain", 3.333), new OfdBrush(OfdColor.Black), 1, 10, deltas);
        var saved = await RoundTrip(package);
        var savedPath = Assert.IsType<OfdPathElement>(saved.Pages[0].Elements[0]);
        var savedText = Assert.IsType<OfdTextElement>(saved.Pages[0].Elements[1]);
        Assert.DoesNotContain('E', savedPath.AbbreviatedData);
        Assert.DoesNotContain('e', savedPath.AbbreviatedData);
        Assert.Contains("0.00000000000000000001", savedPath.AbbreviatedData);
        Assert.Contains("-100000000000000000000", savedPath.AbbreviatedData);
        var literals = Assert.Single(savedText.Runs).DeltaX!.Split(' ');
        Assert.Equal(deltas.Length, literals.Length);
        Assert.All(literals, literal => { Assert.DoesNotContain('E', literal); Assert.DoesNotContain('e', literal); });
        Assert.StartsWith("0.", literals[2]);
        Assert.Equal(deltas, literals.Select(literal => double.Parse(literal, CultureInfo.InvariantCulture)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NameOnlyItalic_ComposesOneFactorAndKeepsBaselineAcrossResave(bool transformed)
    {
        var (package, graphics) = Create();
        if (transformed)
        {
            graphics.Rotate(23);
            graphics.Scale(1.4, 0.8);
            graphics.MultiplyTransform(new OfdMatrix(1, 0.15, 0.25, 1, 0, 0));
        }
        var user = graphics.Transform;
        const double size = 3.333, x = 21.2, baseline = 40.4;
        graphics.DrawString("斜体 I", new OfdFont("SimSun", size, italic: true), new OfdBrush(OfdColor.Black), x, baseline);
        var once = await RoundTrip(package);
        var text = Assert.IsType<OfdTextElement>(Assert.Single(once.Pages[0].Elements));
        var xml = XElement.Parse(text.SourceXml!);
        var factor = Numbers(xml.Attribute(Marker)!.Value);
        var sizeOnWire = double.Parse(xml.Attribute("Size")!.Value, CultureInfo.InvariantCulture);
        Assert.Equal(new[] { 1d, 0, -0.2, 1, 0.2 * sizeOnWire, 0 }, factor);
        var expected = user.Multiply(OfdMatrix.Translation(x, baseline - size))
            .Multiply(new OfdMatrix(1, 0, -0.2, 1, 0.2 * sizeOnWire, 0));
        var actual = Numbers(xml.Attribute("CTM")!.Value);
        AssertClose(expected.A, actual[0]); AssertClose(expected.B, actual[1]);
        AssertClose(expected.C, actual[2]); AssertClose(expected.D, actual[3]);
        AssertClose(expected.E, actual[4]); AssertClose(expected.F, actual[5]);
        var run = Assert.Single(text.Runs);
        var anchorX = actual[0] * run.XMillimeters + actual[2] * run.YMillimeters + actual[4];
        var anchorY = actual[1] * run.XMillimeters + actual[3] * run.YMillimeters + actual[5];
        var target = user.TransformPoint(x, baseline);
        Assert.InRange(Math.Abs(anchorX - target.X), 0, 0.002);
        Assert.InRange(Math.Abs(anchorY - target.Y), 0, 0.002);
        var twice = await RoundTrip(once);
        var twiceXml = XElement.Parse(Assert.IsType<OfdTextElement>(Assert.Single(twice.Pages[0].Elements)).SourceXml!);
        Assert.Equal(xml.Attribute("CTM")!.Value, twiceXml.Attribute("CTM")!.Value);
        Assert.Equal(xml.Attribute(Marker)!.Value, twiceXml.Attribute(Marker)!.Value);
    }

    [Fact]
    public async Task FractionalFontSize_UnderLargeScale_KeepsWrittenBaselineFixed()
    {
        var (package, graphics) = Create();
        graphics.Scale(1000, 1000);
        graphics.DrawString("I", new OfdFont("SimSun", 3.3334, italic: true),
            new OfdBrush(OfdColor.Black), 0.01, 0.02);
        var saved = await RoundTrip(package);
        var text = Assert.IsType<OfdTextElement>(Assert.Single(saved.Pages[0].Elements));
        var xml = XElement.Parse(text.SourceXml!);
        var matrix = Numbers(xml.Attribute("CTM")!.Value);
        var run = Assert.Single(text.Runs);
        var anchorX = matrix[0] * run.XMillimeters + matrix[2] * run.YMillimeters + matrix[4];
        var anchorY = matrix[1] * run.XMillimeters + matrix[3] * run.YMillimeters + matrix[5];
        Assert.InRange(Math.Abs(anchorX - 10), 0, 0.002);
        Assert.InRange(Math.Abs(anchorY - 20), 0, 0.002);
    }

    [Fact]
    public async Task ResourceItalicAndEmbeddedOrTransparentText_KeepTheirBindingAndSkipFauxAsAppropriate()
    {
        var (package, graphics) = Create();
        package.Fonts.Add(new OfdFontResource { Id = "20", FontName = "Named Italic", Italic = true });
        package.Fonts.Add(new OfdFontResource { Id = "21", FontName = "Embedded Italic", Italic = true, Data = new byte[] { 1, 2, 3 } });
        var black = new OfdBrush(OfdColor.Black);
        graphics.DrawString("A", new OfdFont("Alias", 3.333, resourceId: "20"), black, 10, 10);
        graphics.DrawString("B", new OfdFont("Alias", 3.333, resourceId: "21"), black, 10, 20);
        graphics.DrawString("C", new OfdFont("SimSun", 3.333, italic: true), new OfdBrush(new OfdColor(0, 0, 0, 0)), 10, 30);
        var saved = await RoundTrip(package);
        var texts = saved.Pages[0].Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(3, texts.Length);
        Assert.Equal("Named Italic", texts[0].FontName);
        Assert.NotNull(XElement.Parse(texts[0].SourceXml!).Attribute(Marker));
        Assert.Equal("Embedded Italic", texts[1].FontName);
        Assert.Null(XElement.Parse(texts[1].SourceXml!).Attribute(Marker));
        Assert.Null(XElement.Parse(texts[2].SourceXml!).Attribute(Marker));
    }

    private static double[] Numbers(string value) => value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)
        .Select(token => double.Parse(token, CultureInfo.InvariantCulture)).ToArray();
    private static void AssertClose(double expected, double actual) => Assert.InRange(Math.Abs(expected - actual), 0, 0.002);
    private static async Task<OfdDocumentPackage> RoundTrip(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        return await new OfdReader().ReadAsync(stream);
    }
}
