using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class MergeVisualTests
{
    public MergeVisualTests() => PdfFontRegistry.EnsureInstalled();

    [Fact]
    public async Task TransparentSemanticTextAfterImage_ShouldRemainExtractableWithoutPaintingPixels()
    {
        using var pixels = new Image<Rgba32>(80, 40, new Rgba32(180, 220, 255));
        using var png = new MemoryStream(); pixels.SaveAsPng(png);
        async Task<byte[]> Convert(bool semanticText)
        {
            var package = new OfdDocumentPackage();
            var page = new OfdPage { WidthMillimeters = 60, HeightMillimeters = 40 };
            page.Elements.Add(new OfdImageElement { Data = png.ToArray(), WidthMillimeters = 60, HeightMillimeters = 40 });
            if (semanticText)
                page.Elements.Add(new OfdTextElement { Text = "ORIGINAL SEMANTIC TEXT", FontName = "Arial", FontSizeMillimeters = 4,
                    XMillimeters = 2, YMillimeters = 2, FillColor = new OfdColor(0, 0, 0, 0) });
            page.Elements.Add(new OfdTextElement { Text = "Visible text", FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 2, YMillimeters = 20 });
            package.Pages.Add(page);
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf); return pdf.ToArray();
        }
        var baseline = await Convert(false); var actual = await Convert(true);
        using var expectedReader = DocLib.Instance.GetDocReader(baseline, new PageDimensions(2d));
        using var actualReader = DocLib.Instance.GetDocReader(actual, new PageDimensions(2d));
        using var expectedPage = expectedReader.GetPageReader(0); using var actualPage = actualReader.GetPageReader(0);
        Assert.Equal(expectedPage.GetImage(), actualPage.GetImage());
        using var textReader = UglyToad.PdfPig.PdfDocument.Open(actual);
        Assert.Contains("ORIGINAL SEMANTIC TEXT", textReader.GetPage(1).Text);
        Assert.Contains("Visible text", textReader.GetPage(1).Text);
    }

    [Fact]
    public async Task Merge_ShouldRenderDistinctFontsAndRotatedClippedTransparentImages()
    {
        using var pixels = new Image<Rgba32>(80, 40);
        for (var y = 0; y < 40; y++)
        for (var x = 0; x < 80; x++) pixels[x, y] = y < 20
            ? x < 40 ? new Rgba32(255, 0, 0) : new Rgba32(0, 0, 255)
            : x < 40 ? new Rgba32(255, 255, 0) : new Rgba32(0, 255, 0);
        using var png = new MemoryStream(); pixels.SaveAsPng(png);
        OfdDocumentPackage Source(string variant)
        {
            var source = new OfdDocumentPackage();
            source.Fonts.Add(new OfdFontResource { Id = "20", FontName = "Ofdrw Test Face", Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", variant + ".ttf")) });
            var page = new OfdPage { WidthMillimeters = 148, HeightMillimeters = 210 };
            page.Elements.Add(new OfdTextElement { Text = variant + " font / rotated image", FontName = "Arial", FontSizeMillimeters = 5, XMillimeters = 12, YMillimeters = 12 });
            page.Elements.Add(new OfdTextElement { Text = "ABCD", FontName = "Ofdrw Test Face", FontResourceId = "20", FontSizeMillimeters = 8, XMillimeters = 12, YMillimeters = 28 });
            page.Elements.Add(new OfdImageElement
            {
                Data = png.ToArray(), XMillimeters = 20, YMillimeters = 60, WidthMillimeters = 40, HeightMillimeters = 80,
                Transform = [0, 80, -40, 0, 40, 0], Alpha = 128,
                ClipsXml = "<Clips><Clip><Area><Path><AbbreviatedData>M 5 10 L 35 10 L 35 70 L 5 70 C</AbbreviatedData></Path></Area></Clip></Clips>"
            });
            source.Pages.Add(page); return source;
        }
        var merged = OfdDocumentMerger.Merge([Source("narrow"), Source("wide")]);
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(merged, ofd); ofd.Position = 0;
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        var directory = Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS");
        if (!string.IsNullOrWhiteSpace(directory))
        {
            Directory.CreateDirectory(directory);
            await File.WriteAllBytesAsync(Path.Combine(directory, "merged-styles.ofd"), ofd.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(directory, "merged-styles.pdf"), pdf.ToArray());
            for (var page = 0; page < 2; page++)
            {
                ofd.Position = 0;
                using var svg = File.Create(Path.Combine(directory, $"merged-styles-{page + 1}.svg"));
                await new OfdToSvgConverter().ConvertAsync(ofd, svg, page);
            }
        }
        using var reader = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d));
        for (var index = 0; index < 2; index++)
        {
            using var page = reader.GetPageReader(index);
            var raster = page.GetImage();
            void ColorAt(double x, double y, int r, int g, int b)
            {
                var offset = ((int)Math.Round(y * 72d / 25.4d * 2) * page.GetPageWidth() + (int)Math.Round(x * 72d / 25.4d * 2)) * 4;
                // PDFium returns uncomposited RGBA. Compare the same white
                // background used by Preview, while checking source opacity.
                var alpha = raster[offset + 3];
                int OnWhite(byte channel) => (int)Math.Round(channel * alpha / 255d + 255 - alpha);
                Assert.InRange(OnWhite(raster[offset + 2]), Math.Max(0, r - 8), Math.Min(255, r + 8));
                Assert.InRange(OnWhite(raster[offset + 1]), Math.Max(0, g - 8), Math.Min(255, g + 8));
                Assert.InRange(OnWhite(raster[offset]), Math.Max(0, b - 8), Math.Min(255, b + 8));
                if (x >= 25 && y >= 70) Assert.InRange((int)alpha, 120, 136);
            }
            ColorAt(30, 80, 255, 255, 127); // rotated bottom-left source quadrant
            ColorAt(50, 80, 255, 127, 127);
            ColorAt(30, 120, 127, 255, 127);
            ColorAt(50, 120, 127, 127, 255);
            ColorAt(22, 62, 255, 255, 255); // outside the explicit clip
        }

    }
}
