using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class ImageMainIntegrationTests
{
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task SharedTypedCompositeKeepsBothPlacementsAndEvenOddHoleInImage(bool nested)
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var resource = new OfdCompositeResource { WidthMillimeters = 20, HeightMillimeters = 20 };
        resource.Elements.Add(new OfdPathElement { WidthMillimeters = 20, HeightMillimeters = 20,
            AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C M 5 5 L 15 5 L 15 15 L 5 15 C",
            Fill = true, Stroke = false, FillRule = OfdFillRule.EvenOdd, FillColor = new OfdColor(255, 0, 0) });
        package.CompositeResources.Items.Add(resource);
        var page = new OfdPage { WidthMillimeters = 40, HeightMillimeters = 40 };
        foreach (var offset in new[] { 0, 20 }) page.Elements.Add(new OfdCompositeElement { Resource = resource,
            XMillimeters = offset, YMillimeters = offset, WidthMillimeters = 20, HeightMillimeters = 20 });
        package.Pages.Add(page);
        var bytes = await Write(package);
        if (nested) bytes = await ImageIoTests.WithSeals([bytes], stampsPerRecord: 2, pageCount: 1);
        using var input = new MemoryStream(bytes); using var output = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 8 }).ConvertAsync(input, output);
        var evidence = Environment.GetEnvironmentVariable("OFDRW_IMAGE_EVIDENCE");
        if (!string.IsNullOrWhiteSpace(evidence))
        {
            Directory.CreateDirectory(evidence);
            var name = nested ? "typed-composite-seal" : "typed-composite";
            await File.WriteAllBytesAsync(Path.Combine(evidence, name + ".ofd"), bytes);
            await File.WriteAllBytesAsync(Path.Combine(evidence, name + ".png"), output.ToArray());
            input.Position = 0;
            using var pdf = File.Create(Path.Combine(evidence, name + ".pdf"));
            await new OfdToPdfConverter().ConvertAsync(input, pdf);
        }
        using var image = Image.Load<Rgb24>(output.ToArray());
        void Red(int x, int y) { Assert.True(image[x, y].R > 220); Assert.True(image[x, y].G < 30); }
        if (nested)
        {
            // 40mm nested page shrinks into each 10mm stamp; the second stamp starts 18mm later.
            foreach (var dx in new[] { 0, 144 })
            {
                Red(42 + dx, 82); Red(82 + dx, 122);
                Assert.Equal(Color.White.ToPixel<Rgb24>(), image[60 + dx, 100]);
                Assert.Equal(Color.White.ToPixel<Rgb24>(), image[100 + dx, 140]);
            }
        }
        else
        {
            Red(8, 8); Red(168, 168);
            Assert.Equal(Color.White.ToPixel<Rgb24>(), image[80, 80]);
            Assert.Equal(Color.White.ToPixel<Rgb24>(), image[240, 240]);
            Assert.Equal(Color.White.ToPixel<Rgb24>(), image[200, 20]);
        }
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task UnsupportedTypedClipNeverPublishesPartialImage(bool nested)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 40, HeightMillimeters = 40 };
        var path = new OfdPathElement { WidthMillimeters = 20, HeightMillimeters = 20,
            AbbreviatedData = "M 0 0 L 20 0 L 20 20 C", Fill = true, Clips = new() };
        var clip = new OfdClip();
        for (var i = 0; i < 2; i++) clip.Areas.Add(new OfdClipArea { Path = new OfdPathElement {
            AbbreviatedData = "M 0 0 L 10 0 L 10 10 C", Fill = true, Stroke = false } });
        path.Clips.Items.Add(clip); page.Elements.Add(path); package.Pages.Add(page);
        var bytes = await Write(package);
        if (nested) bytes = await ImageIoTests.WithSeals([bytes]);
        using var input = new MemoryStream(bytes); using var output = new MemoryStream("sentinel"u8.ToArray());
        if (nested) await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter().ConvertAsync(input, output));
        else await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToImageConverter().ConvertAsync(input, output));
        Assert.Equal("sentinel"u8.ToArray(), output.ToArray());
    }

    [Theory]
    [InlineData("missing")] [InlineData("empty")] [InlineData("blank")]
    [InlineData("duplicate")] [InlineData("duplicate-case")]
    public async Task AlreadyLoadedImageEntryStillRejectsAmbiguousPageIdentity(string kind)
    {
        var package = new OfdDocumentPackage();
        var id = kind switch { "missing" => null, "empty" => "", "blank" => " \t", _ => "page-id" };
        package.Pages.Add(new OfdPage { Id = id, WidthMillimeters = 40, HeightMillimeters = 40 });
        package.Pages.Add(new OfdPage { Id = kind == "duplicate-case" ? "PAGE-ID" : id, Index = 1,
            WidthMillimeters = 40, HeightMillimeters = 40 });
        using var output = new MemoryStream("sentinel"u8.ToArray());
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToPdfConverter().ConvertPackageAsync(
            package, output, [0], default, strictAppearanceBudgets: true));
        Assert.Contains("page ID", error.Message);
        Assert.Equal("sentinel"u8.ToArray(), output.ToArray());
    }

    private static async Task<byte[]> Write(OfdDocumentPackage package)
    {
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }
}
