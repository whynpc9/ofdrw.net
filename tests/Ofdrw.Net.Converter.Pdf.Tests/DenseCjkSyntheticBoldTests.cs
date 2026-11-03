using System.IO.Compression;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

// Issue #107: native DOCX bold on a Regular-only Han face is written as a
// 1%-em OFD stroke. Viewers widen that sub-pixel stroke to a one-pixel
// hairline, which closed the counters of 量 at the default 144 DPI export.
public sealed class DenseCjkSyntheticBoldTests
{
    private const string Value = "数量";
    private const double Size = 3.704, X = 79.117, Y = 71.091, Advance = 3.704167;
    private const double Ppm = 144d / 25.4;

    [Fact]
    public async Task RegularOnlyHanBoldKeepsDenseCountersOpenInTheDefaultPublicImageExport()
    {
        var regular = Measure(await Render(OfdTextElement.DefaultWeight, null));
        var boldPackage = Package(OfdTextElement.BoldWeight, "0.037");
        var bold = Measure(await Render(boldPackage));
        // The same windows and threshold as the original PR78 regression:
        // the two upper apertures of 量 at x=300..700, y=610..665 and 709..763
        // (1000 UPM) must keep a visibly light row.
        Assert.True(bold.Lower > 100 && bold.Upper > 100,
            $"Dense CJK apertures became dark at default 144 DPI: lower={bold.Lower:F1}, upper={bold.Upper:F1}; " +
            $"regular lower={regular.Lower:F1}, upper={regular.Upper:F1}.");
        Assert.True(bold.TopBarPainted, "The original upper bar must still paint; blank output cannot satisfy counter clearance.");
        // Emphasis must stay visibly heavier than the same Regular face
        // (measured 1.23x on the pinned fixture); this is a regression floor.
        Assert.True(bold.Ink >= regular.Ink * 1.15, $"Synthetic bold lost its emphasis: bold={bold.Ink:F1}, regular={regular.Ink:F1}.");

        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(boldPackage, ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Equal(Value, semantic.GetPage(1).Text);
        Assert.Equal(2, semantic.GetPage(1).Letters.Count);
    }

    [Fact]
    public void HorizontalEmboldenOverlayNeverPaintsTheGapBetweenStackedStrokes()
    {
        // Two horizontal bars (y=0..2 and y=3..5) joined by a left stem (x=0..1).
        // The gap y=2..3 right of the stem is a counter.
        var contour = new List<(double X, double Y)>
        {
            (0, 0), (10, 0), (10, 2), (1, 2), (1, 3), (10, 3), (10, 5), (0, 5)
        };
        const double h = .25;
        var quads = HorizontalEmboldenOverlay.Build(new[] { contour }, h);
        Assert.NotEmpty(quads);
        foreach (var quad in quads)
        {
            var top = quad.Min(point => point.Y);
            var bottom = quad.Max(point => point.Y);
            var left = quad.Min(point => point.X);
            var right = quad.Max(point => point.X);
            // Only horizontal growth: nothing above or below the glyph.
            Assert.InRange(top, 0, 5);
            Assert.InRange(bottom, 0, 5);
            Assert.InRange(left, -h - 1e-9, 10 + h + 1e-9);
            Assert.InRange(right, -h - 1e-9, 10 + h + 1e-9);
            // In the gap rows, only the stem's edges may grow.
            if (top >= 2 - 1e-9 && bottom <= 3 + 1e-9)
                Assert.True(right <= 1 + h + 1e-9, $"Overlay entered the counter: {left}..{right} at {top}..{bottom}.");
        }
        // The stem's inner edge does grow into the counter by exactly h.
        Assert.Contains(quads, quad => quad.Min(p => p.Y) >= 2 - 1e-9 && quad.Max(p => p.Y) <= 3 + 1e-9 &&
            Math.Abs(quad.Max(p => p.X) - (1 + h)) < 1e-9);
        // Outer sides grow by h.
        Assert.Contains(quads, quad => Math.Abs(quad.Min(p => p.X) + h) < 1e-9);
        Assert.Contains(quads, quad => Math.Abs(quad.Max(p => p.X) - (10 + h)) < 1e-9);
    }

    private static OfdDocumentPackage Package(int weight, string? lineWidth)
    {
        var stroke = lineWidth is null ? "" : $" Stroke='true' Fill='true' LineWidth='{lineWidth}'";
        var strokeColor = lineWidth is null ? "" : "<StrokeColor Value='0 0 0'/>";
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "pinned CJK", Data = CjkFont(Value) });
        package.Pages.Add(new OfdPage
        {
            WidthMillimeters = 210, HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = Value, FontName = "pinned CJK", FontResourceId = "10", FontSizeMillimeters = Size,
                    Weight = weight, XMillimeters = X, YMillimeters = Y, WidthMillimeters = 7.408, HeightMillimeters = 4.815,
                    SourceXml = "<TextObject xmlns='http://www.ofdspec.org' Boundary='79.117 71.091 7.408 4.815' " +
                        $"Font='10' Size='3.704' Weight='{weight}'{stroke}>{strokeColor}" +
                        "<TextCode X='0' Y='3.704' DeltaX='3.704167'>数量</TextCode></TextObject>"
                }
            }
        });
        return package;
    }

    private static Task<byte[]> Render(int weight, string? lineWidth) => Render(Package(weight, lineWidth));

    private static async Task<byte[]> Render(OfdDocumentPackage package)
    {
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var png = new MemoryStream();
        await new OfdToImageConverter().ConvertAsync(ofd, png);
        return png.ToArray();
    }

    private static (double Lower, double Upper, double Ink, bool TopBarPainted) Measure(byte[] png)
    {
        using var image = Image.Load<Rgb24>(png);
        var left = (int)Math.Ceiling((X + Advance + Size * .300) * Ppm);
        var right = (int)Math.Floor((X + Advance + Size * .700) * Ppm);
        double BrightestCounterRow(double bottom, double top)
        {
            var first = (int)Math.Floor((Y + Size - Size * top / 1000) * Ppm);
            var last = (int)Math.Ceiling((Y + Size - Size * bottom / 1000) * Ppm);
            var brightest = 0d;
            for (var row = first; row < last; row++)
            {
                var total = 0d;
                for (var column = left; column <= right; column++) total += image[column, row].R;
                brightest = Math.Max(brightest, total / (right - left + 1));
            }
            return brightest;
        }
        var inkRows = new[] { 780d, 790d, 800d }.Select(designY => (int)Math.Round((Y + Size - Size * designY / 1000) * Ppm)).Distinct();
        var topBar = inkRows.Any(row => Enumerable.Range(left, right - left + 1).Count(column => image[column, row].R < 150) >= 4);
        // Total ink over both glyph cells, for the relative emphasis check.
        var x0 = (int)Math.Floor(X * Ppm) - 2;
        var x1 = (int)Math.Ceiling((X + 2 * Advance) * Ppm) + 2;
        var y0 = (int)Math.Floor((Y - .5) * Ppm);
        var y1 = (int)Math.Ceiling((Y + Size + 1) * Ppm);
        var ink = 0d;
        for (var row = y0; row <= y1; row++)
            for (var column = x0; column <= x1; column++) ink += 255 - image[column, row].R;
        return (BrightestCounterRow(610, 665), BrightestCounterRow(709, 763), ink / 1000, topBar);
    }

    private static byte[] CjkFont(string value)
    {
        using var archive = ZipFile.OpenRead(Path.Combine(AppContext.BaseDirectory, "fonts", "trusted-fonts.zip"));
        using var entry = archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!.Open();
        using var bytes = new MemoryStream();
        entry.CopyTo(bytes);
        return TrueTypeSubset.Create(new OpenTypeFace(bytes.ToArray()), value.Select(c => (int)c),
            CancellationToken.None, 1_000_000, out _, out _);
    }
}
