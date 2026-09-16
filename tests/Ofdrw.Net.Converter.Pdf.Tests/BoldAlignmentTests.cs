using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class BoldAlignmentTests
{
    [Theory]
    [InlineData(0, false)]
    [InlineData(1, false)]
    [InlineData(2, false)]
    [InlineData(0, true)]
    [InlineData(1, true)]
    [InlineData(2, true)]
    public async Task SimulatedBold_ShouldThickenAtTheOriginalBaselineWithoutDuplicateText(int positioning, bool italic)
    {
        async Task<byte[]> Convert(bool bold)
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "fixture", Bold = bold, Italic = italic,
                Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "style-metrics.ttf")) });
            var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
            var text = new OfdTextElement { Text = "AB中文", FontName = "fixture", FontResourceId = "10",
                FontSizeMillimeters = 8, XMillimeters = 15, YMillimeters = 20 };
            if (positioning != 0)
                text.Runs.Add(new OfdTextRun { Text = text.Text, YMillimeters = 8,
                    DeltaX = positioning == 2 ? "5 5 5" : null });
            page.Elements.Add(text); package.Pages.Add(page);
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            return pdf.ToArray();
        }
        var regular = await Convert(false); var bold = await Convert(true);
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(bold);
        Assert.Equal("AB中文", semantic.GetPage(1).Text);
        using var regularReader = DocLib.Instance.GetDocReader(regular, new PageDimensions(4d));
        using var boldReader = DocLib.Instance.GetDocReader(bold, new PageDimensions(4d));
        using var regularPage = regularReader.GetPageReader(0); using var boldPage = boldReader.GetPageReader(0);
        var expected = Ink(regularPage.GetImage(), regularPage.GetPageWidth());
        var actual = Ink(boldPage.GetImage(), boldPage.GetPageWidth());
        Assert.True(actual.Count > expected.Count * 1.025, "Requested bold must visibly increase ink coverage.");
        // At 4 pixels/point the 0.025em stroke expands by ~1.2 pixels per side.
        // A second, displaced outline instead expands the bounds by ~23 pixels.
        Assert.InRange(Math.Abs(actual.Top - expected.Top), 0, 3);
        Assert.InRange(Math.Abs(actual.Bottom - expected.Bottom), 0, 3);
        Assert.InRange(Math.Abs(actual.Left - expected.Left), 0, 3);
        Assert.InRange(Math.Abs(actual.Right - expected.Right), 0, 3);
    }

    private static (int Left, int Top, int Right, int Bottom, int Count) Ink(byte[] pixels, int width)
    {
        var left = width; var top = int.MaxValue; var right = 0; var bottom = 0; var count = 0;
        for (var i = 0; i < pixels.Length; i += 4)
        {
            if (pixels[i + 3] < 128 || pixels[i] > 128 || pixels[i + 1] > 128 || pixels[i + 2] > 128) continue;
            var x = i / 4 % width; var y = i / 4 / width;
            left = Math.Min(left, x); right = Math.Max(right, x);
            top = Math.Min(top, y); bottom = Math.Max(bottom, y); count++;
        }
        Assert.True(count > 0);
        return (left, top, right, bottom, count);
    }
}
