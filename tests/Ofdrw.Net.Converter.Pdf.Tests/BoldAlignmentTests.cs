using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class BoldAlignmentTests
{
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task DecomposedLatinI_ShouldPaintTheCompositeWithoutExtraDotOrDuplicateText(bool bold, bool italic)
    {
        async Task<byte[]> Convert(string value)
        {
            var package = new OfdDocumentPackage();
            var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
            var text = new OfdTextElement { Text = value, FontName = "Arial", FontSizeMillimeters = 10,
                Weight = bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight, Italic = italic,
                XMillimeters = 15, YMillimeters = 20, WidthMillimeters = 20, HeightMillimeters = 15 };
            text.Runs.Add(new OfdTextRun { Text = value, YMillimeters = 10, DeltaX = "3" });
            page.Elements.Add(text);package.Pages.Add(page);
            using var ofd = new MemoryStream();await new OfdPackageWriter().WriteAsync(package, ofd);ofd.Position = 0;
            var reread = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd);
            Assert.Equal(value, Assert.Single(Assert.Single(reread.Pages).Elements.OfType<OfdTextElement>()).Text);
            ofd.Position = 0;using var pdf = new MemoryStream();
            await new OfdToPdfConverter().ConvertAsync(ofd, pdf);return pdf.ToArray();
        }
        var composed = await Convert("íB");var decomposed = await Convert("i\u0301B");
        using var expectedReader = DocLib.Instance.GetDocReader(composed, new PageDimensions(4d));
        using var actualReader = DocLib.Instance.GetDocReader(decomposed, new PageDimensions(4d));
        using var expectedPage = expectedReader.GetPageReader(0);using var actualPage = actualReader.GetPageReader(0);
        Assert.Equal(expectedPage.GetImage(), actualPage.GetImage());
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(decomposed);
        Assert.Equal("íB", semantic.GetPage(1).Text);
        Assert.Equal(2, semantic.GetPage(1).Letters.Count);
    }

    [Theory]
    [InlineData("\u00A0", false)]
    [InlineData("\u202F", false)]
    [InlineData("\u2007", false)]
    [InlineData("\u00A0", true)]
    [InlineData("\u202F", true)]
    [InlineData("\u2007", true)]
    public async Task PositionedNonbreakingWhitespace_ShouldNotPaintMissingGlyphBox(string separator, bool bold)
    {
        async Task<byte[]> Convert(bool withSeparator)
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "fixture", Bold = bold,
                Data = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "style-metrics.ttf")) });
            var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
            void Add(string value, double x, string? deltaX)
            {
                var text = new OfdTextElement { Text = value, FontName = "fixture", FontResourceId = "10",
                    FontSizeMillimeters = 8, XMillimeters = x, YMillimeters = 20,
                    WidthMillimeters = 30, HeightMillimeters = 15 };
                text.Runs.Add(new OfdTextRun { Text = value, YMillimeters = 8, DeltaX = deltaX });
                page.Elements.Add(text);
            }
            if (withSeparator) Add("A" + separator + "B", 20, "8 8");
            else { Add("A", 20, null); Add("B", 36, null); }
            package.Pages.Add(page);
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            return pdf.ToArray();
        }
        var baseline = await Convert(false); var actual = await Convert(true);
        using var expectedReader = DocLib.Instance.GetDocReader(baseline, new PageDimensions(4d));
        using var actualReader = DocLib.Instance.GetDocReader(actual, new PageDimensions(4d));
        using var expectedPage = expectedReader.GetPageReader(0); using var actualPage = actualReader.GetPageReader(0);
        Assert.Equal(expectedPage.GetImage(), actualPage.GetImage());
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(actual);
        Assert.Contains(separator, semantic.GetPage(1).Text);
        Assert.Equal(1, semantic.GetPage(1).Letters.Count(letter => letter.Value == "A"));
        Assert.Equal(1, semantic.GetPage(1).Letters.Count(letter => letter.Value == "B"));
    }

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
