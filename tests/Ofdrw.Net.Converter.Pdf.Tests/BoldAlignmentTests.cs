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
        if (separator == "\u202F" && Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory);
            await File.WriteAllBytesAsync(Path.Combine(directory, $"positioned-narrow-space-{bold}.pdf"), actual);
        }
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

    [Theory]
    [InlineData("OO", "NotoSans-Regular.ttf", "0.2", "0 0 0", true)]
    [InlineData("OO", "NotoSans-Regular.ttf", "0.1", "0 0 0", false)]
    [InlineData("OO", "NotoSans-Regular.ttf", "0.2", "180 0 0", false)]
    [InlineData("O中", "LXGWWenKai-Regular.ttf", "0.08", "0 0 0", true)]
    [InlineData("数量", "LXGWWenKai-Regular.ttf", "0.2", "0 0 0", true)]
    [InlineData("数量", "LXGWWenKai-Regular.ttf", "0.16", "0 0 0", true)]
    public async Task ExplicitFilledStrokeOnRegularFaceDoesNotPaintASecondBoldOutline(
        string value, string fontFile, string lineWidth, string strokeColor, bool sameVisibleStroke)
    {
        async Task<byte[]> Convert(int weight)
        {
            var root = new DirectoryInfo(AppContext.BaseDirectory);
            while (root is not null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
            var fontPath = Path.Combine(root!.FullName,
                "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts", fontFile);
            var fontName = fontFile == "LXGWWenKai-Regular.ttf" ? "LXGW WenKai" : "Noto Sans";
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = fontName,
                Data = File.ReadAllBytes(fontPath) });
            package.Pages.Add(new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60,
                Elements = { new OfdTextElement
                {
                    Text = value, FontName = fontName, FontResourceId = "10",
                    Weight = weight, FontSizeMillimeters = 8,
                    XMillimeters = 15, YMillimeters = 20, WidthMillimeters = 30, HeightMillimeters = 15,
                    SourceXml = $"<TextObject xmlns='http://www.ofdspec.org' Boundary='15 20 30 15' " +
                        $"Font='10' Size='8' Weight='{weight}' Stroke='true' Fill='true' LineWidth='{lineWidth}'>" +
                        $"<StrokeColor Value='{strokeColor}'/><TextCode X='2' Y='10'" +
                        (value == "O中" ? " DeltaX='8'" : "") + $">{value}</TextCode></TextObject>"
                } } });
            using var ofd = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, ofd);
            ofd.Position = 0;
            using var pdf = new MemoryStream();
            await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            return pdf.ToArray();
        }

        var explicitStroke = await Convert(OfdTextElement.DefaultWeight);
        var explicitStrokeWithBoldWeight = await Convert(OfdTextElement.BoldWeight);
        using var regularReader = DocLib.Instance.GetDocReader(explicitStroke, new PageDimensions(4d));
        using var boldReader = DocLib.Instance.GetDocReader(explicitStrokeWithBoldWeight, new PageDimensions(4d));
        using var regularPage = regularReader.GetPageReader(0);
        using var boldPage = boldReader.GetPageReader(0);
        // A same-color explicit stroke wider than synthetic emphasis already
        // covers it, even when its width is neither historical 1% nor 2.5% em.
        if (sameVisibleStroke) Assert.Equal(regularPage.GetImage(), boldPage.GetImage());
        else Assert.NotEqual(regularPage.GetImage(), boldPage.GetImage());
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(explicitStrokeWithBoldWeight);
        Assert.Equal(value, semantic.GetPage(1).Text);
        Assert.Equal(2, semantic.GetPage(1).Letters.Count);
    }

    [Fact]
    public async Task OutlineOnlyTextKeepsOpenCounterWhenBoldIsSimulated()
    {
        async Task<byte[]> Convert(bool bold)
        {
            var dir = new DirectoryInfo(AppContext.BaseDirectory);
            while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "Ofdrw.Net.sln"))) dir = dir.Parent;
            var regularFace = Path.Combine(dir!.FullName,
                "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf");
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans",
                Data = File.ReadAllBytes(regularFace) });
            var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
            page.Elements.Add(new OfdTextElement
            {
                Text = "O", FontName = "Noto Sans", FontResourceId = "10", FontSizeMillimeters = 8,
                Weight = bold ? OfdTextElement.BoldWeight : OfdTextElement.DefaultWeight,
                XMillimeters = 15, YMillimeters = 20, WidthMillimeters = 30, HeightMillimeters = 15,
                SourceXml = "<TextObject xmlns='http://www.ofdspec.org' Boundary='15 20 30 15' Size='8' " +
                    "Stroke='true' Fill='false' LineWidth='0.1'><StrokeColor Value='0 0 0'/>" +
                    "<TextCode X='2' Y='10'>O</TextCode></TextObject>"
            });
            package.Pages.Add(page);
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd);
            ofd.Position = 0;
            using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } directory)
            {
                Directory.CreateDirectory(directory);
                var name = bold ? "outline-only-simulated-bold" : "outline-only-regular";
                await File.WriteAllBytesAsync(Path.Combine(directory, name + ".ofd"), ofd.ToArray());
                await File.WriteAllBytesAsync(Path.Combine(directory, name + ".pdf"), pdf.ToArray());
            }
            return pdf.ToArray();
        }
        var regular = await Convert(false); var bold = await Convert(true);
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(bold);
        Assert.Equal("O", semantic.GetPage(1).Text);
        using var regularReader = DocLib.Instance.GetDocReader(regular, new PageDimensions(4d));
        using var boldReader = DocLib.Instance.GetDocReader(bold, new PageDimensions(4d));
        using var regularPage = regularReader.GetPageReader(0); using var boldPage = boldReader.GetPageReader(0);
        var regularInk = Ink(regularPage.GetImage(), regularPage.GetPageWidth());
        var boldPixels = boldPage.GetImage(); var width = boldPage.GetPageWidth();
        var boldInk = Ink(boldPixels, width);
        Assert.True(boldInk.Count > regularInk.Count * 1.1, "Simulated bold must visibly thicken the unfilled outline.");
        Assert.InRange(Math.Abs(boldInk.Left - regularInk.Left), 0, 5);
        Assert.InRange(Math.Abs(boldInk.Right - regularInk.Right), 0, 5);
        var clearCounterPixels = 0;
        for (var y = boldInk.Top + (boldInk.Bottom - boldInk.Top) / 3;
            y <= boldInk.Bottom - (boldInk.Bottom - boldInk.Top) / 3; y++)
        for (var x = boldInk.Left + (boldInk.Right - boldInk.Left) / 3;
            x <= boldInk.Right - (boldInk.Right - boldInk.Left) / 3; x++)
        {
            var offset = (y * width + x) * 4;
            if (boldPixels[offset + 3] < 128 ||
                boldPixels[offset] > 200 && boldPixels[offset + 1] > 200 && boldPixels[offset + 2] > 200)
                clearCounterPixels++;
        }
        Assert.True(clearCounterPixels > 0, $"The O counter must remain unfilled: bounds={boldInk}.");
    }

    [Theory]
    [InlineData("bold.ttf", true, false)]
    [InlineData("italic.ttf", false, true)]
    public async Task OutlineOnlyTextUsesPhysicalFaceStyle(string fileName, bool bold, bool italic)
    {
        var dir = new DirectoryInfo(AppContext.BaseDirectory);
        while (dir is not null && !File.Exists(Path.Combine(dir.FullName, "Ofdrw.Net.sln"))) dir = dir.Parent;
        var fontPath = Path.Combine(dir!.FullName, "e2e/Ofdrw.Net.SkiaSharp.E2E/testdata/fonts", fileName);
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "physical fixture", Bold = bold,
            Italic = italic, Data = File.ReadAllBytes(fontPath) });
        var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 60 };
        page.Elements.Add(new OfdTextElement
        {
            Text = "O", FontName = "physical fixture", FontResourceId = "10", FontSizeMillimeters = 8,
            XMillimeters = 15, YMillimeters = 20, WidthMillimeters = 30, HeightMillimeters = 15,
            SourceXml = "<TextObject xmlns='http://www.ofdspec.org' Boundary='15 20 30 15' Size='8' " +
                "Stroke='true' Fill='false' LineWidth='0.1'><StrokeColor Value='0 0 0'/>" +
                "<TextCode X='2' Y='10'>O</TextCode></TextObject>"
        });
        package.Pages.Add(page);
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var semantic = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Equal("O", semantic.GetPage(1).Text);
        Assert.Single(semantic.GetPage(1).Letters);
        using var reader = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(4d));
        using var rendered = reader.GetPageReader(0);
        Assert.True(Ink(rendered.GetImage(), rendered.GetPageWidth()).Count > 0);
        if (Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS") is { Length: > 0 } artifacts)
        {
            Directory.CreateDirectory(artifacts);
            var name = bold ? "outline-only-physical-bold" : "outline-only-physical-italic";
            await File.WriteAllBytesAsync(Path.Combine(artifacts, name + ".ofd"), ofd.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(artifacts, name + ".pdf"), pdf.ToArray());
        }
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
