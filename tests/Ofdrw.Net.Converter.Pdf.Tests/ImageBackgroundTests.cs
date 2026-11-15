using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Png;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class ImageBackgroundTests
{
    [Fact]
    public async Task DefaultStaysOpaqueWhiteRgb()
    {
        var (image, bytes) = await Export(new OfdToImageOptions(), artifactName: "background-default.png");
        using (image)
        {
            Assert.Equal(400, image.Width);
            Assert.Equal(400, image.Height);
            Assert.Equal(new Rgba32(255, 255, 255, 255), image[350, 380]);
            Assert.Equal(new Rgba32(255, 0, 0, 255), image[100, 100]);
            var halfBlue = image[300, 100];
            Assert.InRange(halfBlue.R, 126, 130);
            Assert.InRange(halfBlue.G, 126, 130);
            Assert.Equal(255, halfBlue.B);
            Assert.Equal(255, halfBlue.A);
            Assert.Equal(PngColorType.Rgb, Image.Identify(bytes).Metadata.GetPngMetadata().ColorType);
        }
    }

    [Fact]
    public async Task TransparentPngKeepsThePagesAlpha()
    {
        var (image, bytes) = await Export(
            new OfdToImageOptions { Background = OfdImageBackground.Transparent },
            artifactName: "background-transparent.png");
        using (image)
        {
            Assert.Equal(0, image[350, 380].A);
            Assert.Equal(new Rgba32(255, 0, 0, 255), image[100, 100]);
            var halfBlue = image[300, 100];
            Assert.Equal(0, halfBlue.R);
            Assert.Equal(0, halfBlue.G);
            Assert.Equal(255, halfBlue.B);
            Assert.InRange(halfBlue.A, 125, 130);
            var overlap = image[150, 300];
            Assert.InRange(overlap.R, 126, 130);
            Assert.Equal(0, overlap.G);
            Assert.InRange(overlap.B, 125, 129);
            Assert.Equal(255, overlap.A);
            Assert.Equal(PngColorType.RgbWithAlpha, Image.Identify(bytes).Metadata.GetPngMetadata().ColorType);
        }
    }

    [Fact]
    public async Task OpaqueColourIsUsedAsBackground()
    {
        var (image, bytes) = await Export(
            new OfdToImageOptions { Background = OfdImageBackground.Opaque(0, 255, 0) },
            artifactName: "background-green.png");
        using (image)
        {
            Assert.Equal(new Rgba32(0, 255, 0, 255), image[350, 380]);
            Assert.Equal(new Rgba32(255, 0, 0, 255), image[100, 100]);
            var halfBlue = image[300, 100];
            Assert.Equal(0, halfBlue.R);
            Assert.InRange(halfBlue.G, 126, 130);
            Assert.InRange(halfBlue.B, 125, 129);
            Assert.Equal(255, halfBlue.A);
            Assert.Equal(PngColorType.Rgb, Image.Identify(bytes).Metadata.GetPngMetadata().ColorType);
        }
    }

    [Fact]
    public async Task JpegUsesTheOpaqueBackground()
    {
        var (image, _) = await Export(new OfdToImageOptions
        {
            Format = OfdImageFormat.Jpeg,
            Background = OfdImageBackground.Opaque(0, 0, 0)
        }, artifactName: "background-black.jpg");
        using (image)
        {
            var blank = image[350, 380];
            Assert.True(blank.R < 16 && blank.G < 16 && blank.B < 16, $"Expected near black, found {blank}.");
        }
    }

    [Fact]
    public void JpegWithTransparentBackgroundIsRejected()
    {
        Assert.Throws<ArgumentException>("options", () => new OfdToImageConverter(new OfdToImageOptions
        {
            Format = OfdImageFormat.Jpeg,
            Background = OfdImageBackground.Transparent
        }));
    }

    [Fact]
    public void NullBackgroundIsRejected()
    {
        Assert.Throws<ArgumentException>("options", () => new OfdToImageConverter(new OfdToImageOptions
        {
            Background = null!
        }));
    }

    [Fact]
    public async Task OptionsAreCopiedAtConstruction()
    {
        var options = new OfdToImageOptions
        {
            PixelsPerMillimeter = 10,
            Background = OfdImageBackground.Transparent
        };
        var converter = new OfdToImageConverter(options);
        options.Background = OfdImageBackground.White;

        using var input = await Fixture();
        using var output = new MemoryStream();
        await converter.ConvertAsync(input, output);
        using var image = Image.Load<Rgba32>(output.ToArray());
        Assert.Equal(0, image[350, 380].A);
    }

    [Fact]
    public async Task ResizedTransparentPageKeepsItsEdges()
    {
        var (image, _) = await Export(
            new OfdToImageOptions { Background = OfdImageBackground.Transparent },
            size: 40.03, artifactName: "background-transparent-resized.png");
        using (image)
        {
            Assert.Equal(401, image.Width);
            Assert.Equal(401, image.Height);
            Assert.Equal(0, image[395, 395].A);
            Assert.Equal(new Rgba32(255, 0, 0, 255), image[100, 100]);

            // The bottom edge of the blue rectangle fades to transparent without losing its straight RGB.
            var transition = Enumerable.Range(197, 8).Select(y => image[300, y])
                .Where(pixel => pixel.A is > 0 and < 125).ToArray();
            Assert.NotEmpty(transition);
            Assert.All(transition, pixel =>
            {
                Assert.InRange(pixel.R, 0, 2);
                Assert.InRange(pixel.G, 0, 2);
                Assert.InRange(pixel.B, 253, 255);
            });
        }
    }

    [Theory]
    [InlineData(100L, long.MaxValue)]
    [InlineData(long.MaxValue, 16L * 400 * 400 - 1)]
    public async Task BudgetAppliesToTransparentExport(long maxPixels, long maxWorkingBytes)
    {
        using var input = await Fixture();
        using var output = new MemoryStream();
        var converter = new OfdToImageConverter(new OfdToImageOptions
        {
            PixelsPerMillimeter = 10,
            Background = OfdImageBackground.Transparent,
            MaxPixels = maxPixels,
            MaxRasterWorkingBytes = maxWorkingBytes
        });
        await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    private static async Task<(Image<Rgba32> Image, byte[] Bytes)> Export(
        OfdToImageOptions options, double size = 40, string? artifactName = null)
    {
        options.PixelsPerMillimeter = 10;
        using var input = await Fixture(size);
        using var output = new MemoryStream();
        await new OfdToImageConverter(options).ConvertAsync(input, output);
        var bytes = output.ToArray();
        if (artifactName is not null) await SaveArtifacts(artifactName, input, bytes);
        return (Image.Load<Rgba32>(bytes), bytes);
    }

    private static async Task SaveArtifacts(string fileName, MemoryStream input, byte[] bytes)
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS");
        if (string.IsNullOrWhiteSpace(directory)) return;
        Directory.CreateDirectory(directory);
        await File.WriteAllBytesAsync(Path.Combine(directory, Path.ChangeExtension(fileName, ".ofd")), input.ToArray());
        await File.WriteAllBytesAsync(Path.Combine(directory, fileName), bytes);
    }

    private static OfdPathElement Rect(double x, double y, double w, double h, OfdColor fill) => new()
    {
        XMillimeters = x, YMillimeters = y, WidthMillimeters = w, HeightMillimeters = h,
        AbbreviatedData = FormattableString.Invariant($"M 0 0 L {w} 0 L {w} {h} L 0 {h} C"),
        Stroke = false, Fill = true, FillColor = fill,
    };

    private static async Task<MemoryStream> Fixture(double size = 40)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { Index = 0, WidthMillimeters = size, HeightMillimeters = size };
        page.Elements.Add(Rect(0, 0, 20, 20, new OfdColor(255, 0, 0)));
        page.Elements.Add(Rect(20, 0, 20, 20, new OfdColor(0, 0, 255, 128)));
        page.Elements.Add(Rect(0, 25, 20, 10, new OfdColor(255, 0, 0)));
        page.Elements.Add(Rect(10, 25, 20, 10, new OfdColor(0, 0, 255, 128)));
        package.Pages.Add(page);
        var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        return stream;
    }
}
