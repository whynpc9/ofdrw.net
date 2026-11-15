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
