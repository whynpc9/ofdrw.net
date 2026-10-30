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
    [InlineData(400, false)]
    [InlineData(700, false)]
    [InlineData(700, true)]
    public async Task OpaqueSameColorStrokeCoversTheOriginalCjkFilledContour(int weight, bool transformed)
    {
        const string value = "文档转换视觉回归样例";
        var font = await CjkFont(value);
        const double size = 6.35, width = 0.0635;
        var transform = transformed ? new[] { 1.15, .2, .15, .9, 1d, 0d } : null;
        var clip = transformed ? "<Clips xmlns='http://www.ofdspec.org'><Clip><Area><Path><AbbreviatedData>" +
            "M 0 0 L 45 0 L 45 20 L 0 20 C</AbbreviatedData></Path></Area></Clip></Clips>" : null;
        var actual = new OfdDocumentPackage();
        actual.Fonts.Add(new OfdFontResource { Id = "10", FontName = "pinned CJK", Data = font });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50 };
        page.Elements.Add(new OfdTextElement
        {
            Text = value, FontName = "pinned CJK", FontResourceId = "10", FontSizeMillimeters = size,
            Weight = weight, XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 20,
            Transform = transform, ClippingXml = clip,
            SourceXml = $"<TextObject xmlns='http://www.ofdspec.org' Boundary='10 10 80 20' Font='10' Size='6.35' " +
                (transformed ? "CTM='1.15 0.2 0.15 0.9 1 0' " : "") +
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
                LineWidthMillimeters = width, Transform = transform, ClippingXml = clip
            });
        }
        reference.Pages.Add(referencePage);
        var actualPdf = await Export(actual);
        var expectedPdf = await Export(reference);
        foreach (var element in referencePage.Elements.OfType<OfdPathElement>()) element.Stroke = false;
        var fillOnlyPdf = await Export(reference);
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(actualPdf);
        Assert.Equal(value, semantic.GetPage(1).Text);
        Assert.Equal(value.Length, semantic.GetPage(1).Letters.Count);
        using var actualReader = DocLib.Instance.GetDocReader(actualPdf, new PageDimensions(2d));
        using var expectedReader = DocLib.Instance.GetDocReader(fillOnlyPdf, new PageDimensions(2d));
        using var actualPage = actualReader.GetPageReader(0);
        using var expectedPage = expectedReader.GetPageReader(0);
        var actualPixels = actualPage.GetImage();
        var expectedPixels = expectedPage.GetImage();
        Assert.Equal(expectedPixels.Length, actualPixels.Length);
        var missing = 0;
        for (var offset = 0; offset < expectedPixels.Length; offset += 4)
        {
            // Use the original fill for the interior mask: a low-resolution
            // native reference stroke can expand to a pixel and mark exterior
            // boundary pixels as opaque interior. Check the stroke separately
            // at a resolution where its true width exceeds one pixel.
            if (expectedPixels[offset + 3] >= 250 && expectedPixels[offset] <= 125 &&
                expectedPixels[offset + 1] <= 85 && expectedPixels[offset + 2] <= 38 &&
                (actualPixels[offset + 3] < 230 || actualPixels[offset] > 160 ||
                 actualPixels[offset + 1] > 130 || actualPixels[offset + 2] > 95)) missing++;
        }
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
        {
            Directory.CreateDirectory(artifacts);
            var suffix = weight.ToString(CultureInfo.InvariantCulture) + (transformed ? "-affine-clip" : "");
            await File.WriteAllBytesAsync(Path.Combine(artifacts, $"cjk-filled-stroke-{suffix}.pdf"), actualPdf);
            await File.WriteAllBytesAsync(Path.Combine(artifacts, $"cjk-stroke-reference-{suffix}.pdf"), expectedPdf);
            await File.WriteAllBytesAsync(Path.Combine(artifacts, $"cjk-fill-reference-{suffix}.pdf"), fillOnlyPdf);
        }
        Assert.True(missing == 0, $"{missing} opaque original-contour pixels became light gaps.");
        using var actualDetailReader = DocLib.Instance.GetDocReader(actualPdf, new PageDimensions(8d));
        using var expectedDetailReader = DocLib.Instance.GetDocReader(expectedPdf, new PageDimensions(8d));
        using var actualDetail = actualDetailReader.GetPageReader(0);
        using var expectedDetail = expectedDetailReader.GetPageReader(0);
        var expectedInk = AssertStrokeCoverageWithinOnePixel(expectedDetail.GetImage(), actualDetail.GetImage(),
            expectedDetail.GetPageWidth());
        var actualInk = 0d;
        for (var offset = 3; offset < actualPixels.Length; offset += 4) actualInk += actualPixels[offset];
        actualInk /= actualPixels.Length / 4;
        // A renderer can turn a subpixel outline into a full dark pixel. The
        // integrated area must still agree with the independently rendered true
        // stroke above; this rejects the old light-track/double-boundary output.
        Assert.InRange(actualInk / expectedInk, .98, 1.02);
    }

    private static double AssertStrokeCoverageWithinOnePixel(byte[] expected, byte[] actual, int width)
    {
        Assert.Equal(expected.Length, actual.Length);
        var height = expected.Length / 4 / width;
        bool Neighbor(byte[] pixels, int x, int y)
        {
            for (var dy = -1; dy <= 1; dy++)
                for (var dx = -1; dx <= 1; dx++)
                {
                    var column = x + dx; var row = y + dy;
                    if (column >= 0 && column < width && row >= 0 && row < height &&
                        pixels[(row * width + column) * 4 + 3] >= 230) return true;
                }
            return false;
        }
        var missing = 0; var extra = 0; var expectedForeground = 0; var actualForeground = 0;
        var expectedInk = 0d;
        for (var y = 0; y < height; y++)
            for (var x = 0; x < width; x++)
            {
                var offset = (y * width + x) * 4;
                var before = expected[offset + 3] >= 230;
                var after = actual[offset + 3] >= 230;
                expectedInk += expected[offset + 3];
                if (before) expectedForeground++;
                if (after) actualForeground++;
                if (before && !after && !Neighbor(actual, x, y)) missing++;
                if (after && !before && !Neighbor(expected, x, y)) extra++;
            }
        Assert.True(missing == 0 && extra == 0,
            $"Original stroke coverage differs beyond one raster boundary pixel: missing={missing}, extra={extra}.");
        Assert.True(expectedForeground > 1000 && actualForeground > 1000,
            $"Expected a nonempty glyph foreground: reference={expectedForeground}, actual={actualForeground}.");
        return expectedInk / (width * height);
    }

    [Fact]
    public async Task DenseCjkCountersRemainReadableInTheDefaultPublicImageExport()
    {
        const string value = "数量";
        const double size = 3.704, x = 79.117, y = 71.091, advance = 3.704167;
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "pinned CJK", Data = await CjkFont(value) });
        package.Pages.Add(new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdTextElement
            {
                Text = value, FontName = "pinned CJK", FontResourceId = "10", FontSizeMillimeters = size,
                Weight = 700, XMillimeters = x, YMillimeters = y, WidthMillimeters = 7.408, HeightMillimeters = 4.815,
                SourceXml = "<TextObject xmlns='http://www.ofdspec.org' Boundary='79.117 71.091 7.408 4.815' " +
                    "Font='10' Size='3.704' Weight='700' Stroke='true' Fill='true' LineWidth='0.037'>" +
                    "<StrokeColor Value='0 0 0'/><TextCode X='0' Y='3.704' DeltaX='3.704167'>数量</TextCode></TextObject>"
            } } });
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var png = new MemoryStream();
        await new OfdToImageConverter().ConvertAsync(ofd, png);
        using var image = SixLabors.ImageSharp.Image.Load<SixLabors.ImageSharp.PixelFormats.Rgb24>(png.ToArray());
        const double ppm = 144d / 25.4;
        var left = (int)Math.Ceiling((x + advance + size * .300) * ppm);
        var right = (int)Math.Floor((x + advance + size * .700) * ppm);
        double BrightestCounterRow(double bottom, double top)
        {
            var first = (int)Math.Floor((y + size - size * top / 1000) * ppm);
            var last = (int)Math.Ceiling((y + size - size * bottom / 1000) * ppm);
            var brightest = 0d;
            for (var row = first; row < last; row++)
            {
                var total = 0d;
                for (var column = left; column <= right; column++) total += image[column, row].R;
                brightest = Math.Max(brightest, total / (right - left + 1));
            }
            return brightest;
        }
        // Independent original-fixture geometry: the two upper apertures in 量
        // span x=250..747, y=610..665 and709..763 at1000 UPM. Require a visibly
        // light row across their interior, not merely a mathematically open path.
        var lowerCounter = BrightestCounterRow(610, 665);
        var upperCounter = BrightestCounterRow(709, 763);
        var inkRows = new[] { 780d, 790d, 800d }.Select(designY =>
            (int)Math.Round((y + size - size * designY / 1000) * ppm)).Distinct();
        Assert.True(inkRows.Any(row => Enumerable.Range(left, right - left + 1).Count(column => image[column, row].R < 150) >= 4),
            "The original upper bar must still paint; blank output cannot satisfy counter clearance.");
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
        {
            Directory.CreateDirectory(artifacts);
            await File.WriteAllBytesAsync(Path.Combine(artifacts, "dense-cjk-default-image.png"), png.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifacts, "dense-cjk-default-image.ofd"), ofd.ToArray());
        }
        Assert.True(lowerCounter > 100 && upperCounter > 100,
            $"Dense CJK apertures became dark at default144 DPI: lower={lowerCounter:F1}, upper={upperCounter:F1}.");
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(await Export(package));
        Assert.Equal(value, semantic.GetPage(1).Text);
        Assert.Equal(2, semantic.GetPage(1).Letters.Count);
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

    private static async Task<byte[]> CjkFont(string value)
    {
        using var archive = ZipFile.OpenRead(Path.Combine(AppContext.BaseDirectory, "fonts", "trusted-fonts.zip"));
        using var entry = archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!.Open();
        using var bytes = new MemoryStream();
        await entry.CopyToAsync(bytes);
        return TrueTypeSubset.Create(new OpenTypeFace(bytes.ToArray()), value.Select(c => (int)c),
            CancellationToken.None, 1_000_000, out _, out _);
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
