using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Cli.Tests;

public sealed class ImageCommandTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-image-cli-" + Guid.NewGuid().ToString("N"));
    public ImageCommandTests() => Directory.CreateDirectory(_directory);
    private string PathFor(string name) => Path.Combine(_directory, name);

    [Fact]
    public async Task CommandsImportMultipleInputsAndSelectOneBasedPageWithExplicitFormat()
    {
        using var red = new Image<Rgb24>(40, 20, new Rgb24(255, 0, 0));
        using var blue = new Image<Rgb24>(20, 40, new Rgb24(0, 0, 255));
        // Content magic, rather than suffix, determines the input encoding.
        red.SaveAsPng(PathFor("red.bin")); blue.SaveAsJpeg(PathFor("blue.bin"));
        var ofd = PathFor("images.ofd");
        Assert.Equal(0, await global::Cli.RunAsync(["image-to-ofd", PathFor("red.bin"), PathFor("blue.bin"), "--output", ofd,
            "--ppm", "2", "--page-width", "30", "--page-height", "30"]));
        using (var input = File.OpenRead(ofd))
        {
            var package = await new OfdReader().ReadAsync(input); Assert.Equal(2, package.Pages.Count);
            var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[1].Elements));
            Assert.Equal(10, image.XMillimeters); Assert.Equal(5, image.YMillimeters);
        }
        var jpeg = PathFor("page.jpg");
        Assert.Equal(0, await global::Cli.RunAsync(["ofd-to-image", ofd, jpeg, "--pages", "2", "--ppm", "2", "--format", "jpeg"]));
        Assert.Equal(0xff, File.ReadAllBytes(jpeg)[0]);
        using (var image = Image.Load<Rgb24>(jpeg)) { Assert.Equal(60, image.Width); Assert.True(image[30, 30].B > 220); Assert.True(image[30, 30].R < 30); }
        var png = PathFor("page.png");
        Assert.Equal(0, await global::Cli.RunAsync(["ofd-to-image", ofd, png, "--ppm", "1"]));
        Assert.Equal(0x89, File.ReadAllBytes(png)[0]);
        using (var image = Image.Load<Rgb24>(png)) Assert.True(image[15, 15].R > 220);
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("--pages", "0")] [InlineData("--pages", "2")] [InlineData("--pages", "1,2")]
    [InlineData("--pages", "1-2147483647")] [InlineData("--ppm", "NaN")] [InlineData("--ppm", "Infinity")]
    [InlineData("--format", "gif")] [InlineData("--jpeg-quality", "101")]
    [InlineData("--max-pixels", "1")] [InlineData("--max-input-bytes", "1")]
    [InlineData("--max-output-bytes", "1")] [InlineData("--ppm", "1e100")]
    public async Task ExportRejectsInvalidParametersAndBudgetsWithoutPublishing(string option, string value)
    {
        var source = new OfdDocumentPackage(); source.Pages.Add(new OfdPage { WidthMillimeters = 20, HeightMillimeters = 20 });
        var input = PathFor("input.ofd"); using (var file = File.Create(input)) await new OfdPackageWriter().WriteAsync(source, file);
        var output = PathFor("old.png"); await File.WriteAllTextAsync(output, "preserved");
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-image", input, output, option, value]));
        Assert.Equal("preserved", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("--page-width", "30")] [InlineData("--page-width", "NaN")]
    [InlineData("--format", "png")] [InlineData("--pages", "1")]
    [InlineData("--max-pixels", "1")] [InlineData("--max-output-bytes", "1")]
    [InlineData("--max-entries", "1")] [InlineData("--max-total-input-bytes", "1")]
    public async Task ImportRejectsInvalidParametersAndBudgetsWithoutPublishing(string option, string value)
    {
        var input = PathFor("input.png"); using (var image = new Image<Rgb24>(10, 10)) image.SaveAsPng(input);
        var output = PathFor("old.ofd"); await File.WriteAllTextAsync(output, "preserved");
        Assert.Equal(1, await global::Cli.RunAsync(["image-to-ofd", input, output, option, value]));
        Assert.Equal("preserved", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task BadSecondInputAndSameDestinationFailWithoutPublishing()
    {
        var png = PathFor("input.png"); using (var image = new Image<Rgb24>(10, 10)) image.SaveAsPng(png);
        var bad = PathFor("bad.jpg"); await File.WriteAllTextAsync(bad, "not JPEG");
        var output = PathFor("old.ofd"); await File.WriteAllTextAsync(output, "preserved");
        Assert.Equal(1, await global::Cli.RunAsync(["image-to-ofd", png, bad, "--output", output]));
        Assert.Equal("preserved", await File.ReadAllTextAsync(output));
        var original = File.ReadAllBytes(png);
        Assert.Equal(1, await global::Cli.RunAsync(["image-to-ofd", png, png])); Assert.Equal(original, File.ReadAllBytes(png));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    [Fact]
    public async Task MissingMultiImageOutputNeverOverwritesLastImageAndMixedInputsKeepOrder()
    {
        var first = PathFor("red.png"); var second = PathFor("blue.jpg");
        using (var image = new Image<Rgb24>(20, 10, new Rgb24(255, 0, 0))) image.SaveAsPng(first);
        using (var image = new Image<Rgb24>(10, 20, new Rgb24(0, 0, 255))) image.SaveAsJpeg(second);
        var originalFirst = File.ReadAllBytes(first); var originalSecond = File.ReadAllBytes(second);
        Assert.Equal(1, await global::Cli.RunAsync(["image-to-ofd", first, second]));
        Assert.Equal(originalFirst, File.ReadAllBytes(first)); Assert.Equal(originalSecond, File.ReadAllBytes(second));
        var output = PathFor("mixed.ofd");
        Assert.Equal(0, await global::Cli.RunAsync(["image-to-ofd", first, "--input", second, "--output", output, "--ppm", "2"]));
        using var input = File.OpenRead(output); var pages = (await new OfdReader().ReadAsync(input)).Pages;
        Assert.Equal(10, pages[0].WidthMillimeters); Assert.Equal(5, pages[0].HeightMillimeters);
        Assert.Equal(5, pages[1].WidthMillimeters); Assert.Equal(10, pages[1].HeightMillimeters);
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
