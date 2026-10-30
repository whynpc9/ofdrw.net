using System.Globalization;
using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class FilledStrokeCoverageTests
{
    [Theory]
    [InlineData(400)]
    [InlineData(700)]
    public async Task OpaqueSameColorStrokeCoversTheOriginalCjkFilledContour(int weight)
    {
        // Reuse the pinned shared fixture; do not depend on installed OS fonts.
        using var archive = ZipFile.OpenRead(Path.Combine(AppContext.BaseDirectory, "fonts", "trusted-fonts.zip"));
        using var entry = archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!.Open();
        using var bytes = new MemoryStream();
        await entry.CopyToAsync(bytes);
        const string value = "文档转换视觉回归样例";
        var font = TrueTypeSubset.Create(new OpenTypeFace(bytes.ToArray()), value.Select(c => (int)c),
            CancellationToken.None, 1_000_000, out _, out _);
        const double size = 6.35, width = 0.0635;
        var actual = new OfdDocumentPackage();
        actual.Fonts.Add(new OfdFontResource { Id = "10", FontName = "pinned CJK", Data = font });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50 };
        page.Elements.Add(new OfdTextElement
        {
            Text = value, FontName = "pinned CJK", FontResourceId = "10", FontSizeMillimeters = size,
            Weight = weight, XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 20,
            SourceXml = $"<TextObject xmlns='http://www.ofdspec.org' Boundary='10 10 80 20' Font='10' Size='6.35' " +
                $"Weight='{weight}' Stroke='true' Fill='true' LineWidth='0.0635'>" +
                "<FillColor Value='31 78 120'/><StrokeColor Value='31 78 120'/>" +
                "<TextCode X='0' Y='10' DeltaX='6.35 6.35 6.35 6.35 6.35 6.35 6.35 6.35 6.35'>" + value + "</TextCode></TextObject>"
        });
        actual.Pages.Add(page);

        // An independent path-only reference paints the original glyph contours.
        // The assertion is visible coverage, not the converter's PDF operators.
        var reference = new OfdDocumentPackage();
        var referencePage = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50 };
        var outlines = new TrueTypeGlyphOutlines(font);
        var cmap = new OpenTypeCmap(new OpenTypeFace(font));
        for (var index = 0; index < value.Length; index++)
        {
            var data = new StringBuilder();
            var scale = size / outlines.UnitsPerEm;
            string X(double x) => (index * size + x * scale).ToString("R", CultureInfo.InvariantCulture);
            string Y(double y) => (10 - y * scale).ToString("R", CultureInfo.InvariantCulture);
            foreach (var command in outlines.Read(cmap.DirectGlyph(value[index])).ToPathCommands())
            {
                data.Append(command.Verb switch
                {
                    TrueTypePathVerb.Move => $"M {X(command.X)} {Y(command.Y)} ",
                    TrueTypePathVerb.Line => $"L {X(command.X)} {Y(command.Y)} ",
                    TrueTypePathVerb.Quadratic => $"Q {X(command.ControlX)} {Y(command.ControlY)} {X(command.X)} {Y(command.Y)} ",
                    TrueTypePathVerb.Close => "C ",
                    _ => throw new InvalidOperationException()
                });
            }
            referencePage.Elements.Add(new OfdPathElement
            {
                XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 20,
                AbbreviatedData = data.ToString(), Fill = true, Stroke = true,
                FillColor = new OfdColor(31, 78, 120), StrokeColor = new OfdColor(31, 78, 120),
                LineWidthMillimeters = width
            });
        }
        reference.Pages.Add(referencePage);
        var actualPdf = await Export(actual);
        var expectedPdf = await Export(reference);
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(actualPdf);
        Assert.Equal(value, semantic.GetPage(1).Text);
        Assert.Equal(value.Length, semantic.GetPage(1).Letters.Count);
        using var actualReader = DocLib.Instance.GetDocReader(actualPdf, new PageDimensions(2d));
        using var expectedReader = DocLib.Instance.GetDocReader(expectedPdf, new PageDimensions(2d));
        using var actualPage = actualReader.GetPageReader(0);
        using var expectedPage = expectedReader.GetPageReader(0);
        var actualPixels = actualPage.GetImage();
        var expectedPixels = expectedPage.GetImage();
        Assert.Equal(expectedPixels.Length, actualPixels.Length);
        var missing = 0;
        for (var offset = 0; offset < expectedPixels.Length; offset += 4)
        {
            // Interior reference pixels must not turn into light seams. Ignore
            // anti-aliased boundary pixels and any extra hinted text coverage.
            if (expectedPixels[offset + 3] >= 250 && expectedPixels[offset] <= 125 &&
                expectedPixels[offset + 1] <= 85 && expectedPixels[offset + 2] <= 38 &&
                (actualPixels[offset + 3] < 230 || actualPixels[offset] > 160 ||
                 actualPixels[offset + 1] > 130 || actualPixels[offset + 2] > 95)) missing++;
        }
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
        {
            Directory.CreateDirectory(artifacts);
            await File.WriteAllBytesAsync(Path.Combine(artifacts, $"cjk-filled-stroke-{weight}.pdf"), actualPdf);
            await File.WriteAllBytesAsync(Path.Combine(artifacts, "cjk-filled-stroke-reference.pdf"), expectedPdf);
        }
        Assert.True(missing == 0, $"{missing} opaque original-contour pixels became light gaps.");
    }

    [Theory]
    [InlineData("translucent", false)]
    [InlineData("translucent", true)]
    [InlineData("different-color", false)]
    [InlineData("different-color", true)]
    [InlineData("outline-only", false)]
    public async Task GuardedPaintsKeepTheirInteriorCompositing(string mode, bool inherited)
    {
        const double size = 6.35;
        var alpha = mode == "translucent" ? 128 : 255;
        var fill = mode != "outline-only";
        var stroke = mode == "different-color" ? "180 0 0" : "31 78 120";
        async Task<byte[]> Convert(bool withStroke)
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "rectangle fixture",
                Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "narrow.ttf")) });
            package.Pages.Add(new OfdPage { WidthMillimeters = 40, HeightMillimeters = 40,
                Elements = { new OfdTextElement
                {
                    Text = "A", FontName = "rectangle fixture", FontResourceId = "10", FontSizeMillimeters = size,
                    XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 20, HeightMillimeters = 20,
                    SourceXml = "<TextObject xmlns='http://www.ofdspec.org' Boundary='10 10 20 20' " +
                        $"Font='10' Size='6.35' Fill='{fill.ToString().ToLowerInvariant()}' " +
                        $"Stroke='{withStroke.ToString().ToLowerInvariant()}' LineWidth='0.0635'>" +
                        $"<FillColor Value='31 78 120' Alpha='{alpha}'/><StrokeColor Value='{stroke}' Alpha='{alpha}'/>" +
                        "<TextCode X='0' Y='10'>A</TextCode></TextObject>"
                } } });
            return await Export(package, inherited ? archive =>
            {
                Rewrite(archive, "Doc_0/PublicRes.xml", document =>
                {
                    var ns = document.Root!.Name.Namespace;
                    document.Root.Add(new XElement(ns + "DrawParams",
                        new XElement(ns + "DrawParam", new XAttribute("ID", "44"),
                            new XElement(ns + "FillColor", new XAttribute("Value", "31 78 120"), new XAttribute("Alpha", alpha)),
                            new XElement(ns + "StrokeColor", new XAttribute("Value", stroke), new XAttribute("Alpha", alpha))),
                        new XElement(ns + "DrawParam", new XAttribute("ID", "45"),
                            new XAttribute("Relative", "44"), new XAttribute("LineWidth", "0.0635"))));
                });
                Rewrite(archive, "Doc_0/Pages/Page_0/Content.xml", document =>
                {
                    var text = document.Descendants().Single(node => node.Name.LocalName == "TextObject");
                    text.Elements().Where(node => node.Name.LocalName is "FillColor" or "StrokeColor").Remove();
                    text.Attribute("LineWidth")!.Remove();
                    text.SetAttributeValue("DrawParam", "45");
                });
            } : null);
        }
        var reference = await Convert(false);
        var actual = await Convert(true);
        using var referenceReader = DocLib.Instance.GetDocReader(reference, new PageDimensions(2d));
        using var actualReader = DocLib.Instance.GetDocReader(actual, new PageDimensions(2d));
        using var referencePage = referenceReader.GetPageReader(0);
        using var actualPage = actualReader.GetPageReader(0);
        var before = referencePage.GetImage();
        var after = actualPage.GetImage();
        // The pinned glyph is a filled rectangle: this point is well inside it,
        // far from the independent boundary stroke. It must neither darken from
        // an extra translucent fill nor adopt the stroke color or fill a hollow.
        var x = (int)Math.Round((10 + size * .25) * 72 / 25.4 * 2);
        var y = (int)Math.Round((20 - size * .35) * 72 / 25.4 * 2);
        var offset = (y * actualPage.GetPageWidth() + x) * 4;
        Assert.Equal(before.Skip(offset).Take(4), after.Skip(offset).Take(4));
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(actual);
        Assert.Equal("A", semantic.GetPage(1).Text);
        Assert.Single(semantic.GetPage(1).Letters);
    }

    private static void Rewrite(ZipArchive archive, string name, Action<XDocument> change)
    {
        var entry = archive.GetEntry(name)!;
        XDocument document;
        using (var input = entry.Open()) document = XDocument.Load(input);
        change(document);
        entry.Delete();
        using var output = archive.CreateEntry(name).Open();
        document.Save(output);
    }

    private static async Task<byte[]> Export(OfdDocumentPackage package, Action<ZipArchive>? change = null)
    {
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        if (change is not null)
        {
            ofd.Position = 0;
            using var archive = new ZipArchive(ofd, ZipArchiveMode.Update, leaveOpen: true);
            change(archive);
        }
        ofd.Position = 0;
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        return pdf.ToArray();
    }
}
