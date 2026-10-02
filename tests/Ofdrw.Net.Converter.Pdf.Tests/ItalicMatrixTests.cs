using System.Text;
using System.Xml.Linq;
using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class ItalicMatrixTests
{
    public ItalicMatrixTests() => PdfFontRegistry.EnsureInstalled();

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnmarkedShear_ChangesGlyphSlopeEvenWhenEveryBaselineAnchorIsFixed(bool italic)
    {
        // Original MIT rectangle glyphs isolate the matrix from host font selection.
        // The Issue 02 matrix fixes y=Size, so an anchor-only renderer cannot pass.
        async Task<(byte[] Pixels, int Width, int Height)> Render(bool shear)
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Ofdrw Test Face",
                Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts/narrow.ttf")) });
            var page = new OfdPage { WidthMillimeters = 70, HeightMillimeters = 40 };
            var text = new OfdTextElement { Text = "I", FontName = "Ofdrw Test Face", FontResourceId = "10",
                FontSizeMillimeters = 8, XMillimeters = 15, YMillimeters = 15, Italic = italic,
                FillColor = new OfdColor(128, 0, 128), Transform = shear ? [1, 0, -0.2, 1, 1.6, 0] : [1, 0, 0, 1, 0, 0] };
            text.Runs.Add(new OfdTextRun { Text = "I", YMillimeters = 8 });
            page.Elements.Add(text);
            var control = new OfdTextElement { Text = "I", FontName = "Ofdrw Test Face", FontResourceId = "10",
                FontSizeMillimeters = 8, XMillimeters = 45, YMillimeters = 15 };
            control.Runs.Add(new OfdTextRun { Text = "I", YMillimeters = 8 });
            page.Elements.Add(control); package.Pages.Add(page);
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            var saved = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd);
            Assert.DoesNotContain(XElement.Parse(((OfdTextElement)saved.Pages[0].Elements[0]).SourceXml!).Attributes(),
                attribute => attribute.Name.LocalName == "FauxItalicMatrixV1");
            ofd.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(3d)); using var raster = doc.GetPageReader(0);
            return (raster.GetImage(), raster.GetPageWidth(), raster.GetPageHeight());
        }
        (double Slope, int Height, int Ink) Measure((byte[] Pixels, int Width, int Height) raster)
        {
            var rows = new Dictionary<int, int>(); var ink = 0;
            for (var y = 0; y < raster.Height; y++) for (var x = 0; x < raster.Width; x++)
            {
                var offset = (y * raster.Width + x) * 4;
                if (raster.Pixels[offset + 3] < 150 || raster.Pixels[offset + 1] > 50 ||
                    raster.Pixels[offset] < 70 || raster.Pixels[offset + 2] < 70) continue;
                rows.TryAdd(y, x); ink++;
            }
            Assert.NotEmpty(rows);
            var top = rows.Keys.Min() + 5; var bottom = rows.Keys.Max() - 5;
            Assert.True(bottom > top);
            return ((rows[top] - rows[bottom]) / (double)(bottom - top), rows.Count, ink);
        }
        var plain = await Render(false); var transformed = await Render(true);
        var before = Measure(plain); var after = Measure(transformed);
        Assert.InRange(after.Slope - before.Slope, 0.15, 0.25);
        Assert.InRange(after.Height, before.Height - 2, before.Height + 2);
        Assert.InRange(after.Ink / (double)before.Ink, 0.92, 1.08); // thresholded antialiased edge pixels vary with slope
        // Local emphasis/CTM must not alter the adjacent upright control.
        for (var y = 0; y < plain.Height; y++)
            Assert.Equal(plain.Pixels.AsSpan((y * plain.Width + plain.Width * 4 / 7) * 4, plain.Width * 3 / 7 * 4).ToArray(),
                transformed.Pixels.AsSpan((y * transformed.Width + transformed.Width * 4 / 7) * 4, transformed.Width * 3 / 7 * 4).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task GeneratedItalicFactor_IsAppliedOnlyOnceAndUnmarkedUserCtmKeepsItsShear(bool rotate)
    {
        async Task<(byte[] Pixels, XElement Text)> Render(bool factor, bool marked)
        {
            var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
            var matrix = rotate ? (factor ? "0 1 -1 -0.2 0 1.6" : "0 1 -1 0 0 0") : (factor ? "1 0 -0.2 1 1.6 0" : "1 0 0 1 0 0");
            var text = new OfdTextElement { Text = "ABCD", FontName = "Arial", Italic = true, FontSizeMillimeters = 8, XMillimeters = 25, YMillimeters = 20 };
            text.SourceXml = $"<TextObject Boundary='25 20 40 15' Size='8' Italic='true' CTM='{matrix}' {(marked ? "xmlns:h='https://ofdrw.net/style-hints' h:FauxItalicMatrixV1='1 0 -0.2 1 1.6 0'" : "")}><TextCode X='0' Y='8'>ABCD</TextCode></TextObject>";
            page.Elements.Add(text); package.Pages.Add(page);
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            using var doc = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(3d)); using var raster = doc.GetPageReader(0); var pixels = raster.GetImage();
            ofd.Position = 0; using var svg = new MemoryStream(); await new OfdToSvgConverter().ConvertAsync(ofd, svg);
            var node = XDocument.Parse(Encoding.UTF8.GetString(svg.ToArray())).Descendants().Single(node => node.Name.LocalName == "text");
            return (pixels, node);
        }
        var single = await Render(false, false); var generated = await Render(true, true); var user = await Render(true, false);
        Assert.Equal(single.Pixels, generated.Pixels);
        Assert.False(single.Pixels.SequenceEqual(user.Pixels));
        Assert.Equal(single.Text.Attribute("transform")!.Value, generated.Text.Attribute("transform")!.Value);
        Assert.Equal(single.Text.Attribute("x")!.Value, generated.Text.Attribute("x")!.Value);
        Assert.Equal(single.Text.Attribute("y")!.Value, generated.Text.Attribute("y")!.Value);
        Assert.NotEqual(single.Text.Attribute("transform")!.Value, user.Text.Attribute("transform")!.Value);
    }

    [Fact]
    public async Task Writer_RecordsActualFactorOnlyWhenItCreatesNameOnlyItalicCtm()
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdTextElement { Text = "AUTO", Italic = true, FontName = "Arial", FontSizeMillimeters = 5 });
        page.Elements.Add(new OfdTextElement { Text = "USER", Italic = true, FontName = "Arial", Transform = [1,0,-0.2,1,1,0] }); package.Pages.Add(page);
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
        var read = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd);
        var auto = XElement.Parse(((OfdTextElement)read.Pages[0].Elements[0]).SourceXml!); var user = XElement.Parse(((OfdTextElement)read.Pages[0].Elements[1]).SourceXml!);
        Assert.Equal("1 0 -0.2 1 1 0", auto.Attributes().Single(attribute => attribute.Name.LocalName == "FauxItalicMatrixV1").Value);
        Assert.DoesNotContain(user.Attributes(), attribute => attribute.Name.LocalName == "FauxItalicMatrixV1");
    }
}
