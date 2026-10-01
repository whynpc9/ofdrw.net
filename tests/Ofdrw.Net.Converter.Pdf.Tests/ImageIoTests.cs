using System.IO.Compression;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class ImageIoTests
{
    internal static byte[] ImageBytes(bool jpeg = false)
    {
        using var image = new Image<Rgb24>(80, 40);
        for (var y = 0; y < image.Height; y++)
        for (var x = 0; x < image.Width; x++) image[x, y] = y < 20
            ? x < 40 ? new Rgb24(255, 0, 0) : new Rgb24(0, 0, 255)
            : x < 40 ? new Rgb24(255, 255, 0) : new Rgb24(0, 255, 0);
        using var output = new MemoryStream();
        if (jpeg) image.Save(output, new JpegEncoder { Quality = 95 }); else image.SaveAsPng(output);
        return output.ToArray();
    }

    internal static async Task<byte[]> Source()
    {
        var package = new OfdDocumentPackage();
        package.Options.Metadata.Title = "Deterministic image IO fixture";
        for (var index = 0; index < 2; index++)
        {
            var page = new OfdPage { Index = index, WidthMillimeters = 80.3, HeightMillimeters = 60.4 };
            page.Elements.Add(new OfdTextElement { Text = $"Image I/O 样例 {index + 1}", FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 5, YMillimeters = 3, WidthMillimeters = 70, HeightMillimeters = 6 });
            page.Elements.Add(new OfdTextElement { Text = "Bold proportional", FontName = "Arial", Weight = 700, FontSizeMillimeters = 3,
                XMillimeters = 5, YMillimeters = 12, WidthMillimeters = 40, HeightMillimeters = 5 });
            page.Elements.Add(new OfdTextElement { Text = "Italic color", FontName = "Arial", Italic = true, FillColor = new OfdColor(160, 0, 180), FontSizeMillimeters = 3,
                XMillimeters = 40, YMillimeters = 12, WidthMillimeters = 35, HeightMillimeters = 5 });
            page.Elements.Add(new OfdImageElement { Data = ImageBytes(), FileName = "quadrants.png", XMillimeters = 5, YMillimeters = 22,
                WidthMillimeters = 40, HeightMillimeters = 20 });
            page.Elements.Add(new OfdPathElement { XMillimeters = 52, YMillimeters = 25, WidthMillimeters = 20, HeightMillimeters = 18,
                AbbreviatedData = "M 0 0 L 20 0 L 20 18 L 0 18 C", Fill = true, Stroke = true,
                FillColor = index == 0 ? new OfdColor(220, 50, 20) : new OfdColor(20, 80, 220) });
            package.Pages.Add(page);
        }
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }

    [Theory]
    [InlineData(0, OfdImageFormat.Png, 1.25)]
    [InlineData(1, OfdImageFormat.Jpeg, 1.7)]
    [InlineData(1, OfdImageFormat.Png, 0.5)] // below the old 72 DPI clamp
    [InlineData(0, OfdImageFormat.Png, 13)] // above the old 300 DPI clamp
    public async Task Export_EncodesSelectedTextImageAndPathAtExactResolution(int page, OfdImageFormat format, double ppm)
    {
        using var input = new MemoryStream(await Source()); using var output = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions { Format = format, PixelsPerMillimeter = ppm }).ConvertAsync(input, output, page);
        Assert.True(input.CanRead); Assert.True(output.CanWrite);
        var bytes = output.ToArray();
        Assert.Equal(format == OfdImageFormat.Png ? 0x89 : 0xff, bytes[0]);
        using var decoded = Image.Load<Rgb24>(bytes);
        Assert.Equal((int)Math.Ceiling(80.3 * ppm), decoded.Width);
        Assert.Equal((int)Math.Ceiling(60.4 * ppm), decoded.Height);
        var path = decoded[(int)(60 * ppm), (int)(32 * ppm)];
        if (page == 0) { Assert.True(path.R > 180); Assert.True(path.B < 60); }
        else { Assert.True(path.B > 180); Assert.True(path.R < 60); }
        var image = decoded[(int)(12 * ppm), (int)(27 * ppm)];
        Assert.True(image.R > 200); Assert.True(image.B < 60); // first image quadrant
        Assert.Equal(new Rgb24(255, 255, 255), decoded[decoded.Width - 2, decoded.Height - 2]);
        var ink = 0;
        for (var y = 0; y < (int)(11 * ppm); y++)
        for (var x = 0; x < decoded.Width; x++) if (decoded[x, y].R < 180) ink++;
        Assert.True(ink > 3, "Selected page must contain rasterized title text, not merely a colored path.");
    }

    [Fact]
    public async Task DefaultsAndOptionsSnapshot_ArePngAnd144Dpi()
    {
        var options = new OfdToImageOptions(); var converter = new OfdToImageConverter(options);
        options.Format = (OfdImageFormat)99; options.PixelsPerMillimeter = double.NaN;
        options.PackageLoadOptions.MaxEntryCount = 1;
        using var input = new MemoryStream(await Source()); using var output = new MemoryStream();
        await converter.ConvertAsync(input, output);
        using var image = Image.Load(output.ToArray());
        Assert.Equal(0x89, output.ToArray()[0]); Assert.Equal((int)Math.Ceiling(80.3 * 144 / 25.4), image.Width);
    }

    [Fact]
    public async Task Import_OneImagePerPageKeepsOrderBytesAspectRatioAndCentering()
    {
        var png = ImageBytes(); var jpeg = ImageBytes(true);
        using var first = new MemoryStream(png); using var second = new MemoryStream(jpeg); using var output = new MemoryStream();
        var options = new ImageToOfdOptions { PixelsPerMillimeter = 2, PageSize = new OfdPageSize { WidthMillimeters = 30, HeightMillimeters = 30 } };
        var converter = new ImageToOfdConverter(options); options.PageSize.WidthMillimeters = 1; options.PixelsPerMillimeter = 100;
        await converter.ConvertAsync([first, second], output);
        Assert.True(first.CanRead); Assert.True(second.CanRead); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output); Assert.Equal(2, package.Pages.Count);
        for (var page = 0; page < 2; page++)
        {
            Assert.Equal(30, package.Pages[page].WidthMillimeters); Assert.Equal(30, package.Pages[page].HeightMillimeters);
            var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[page].Elements));
            Assert.Equal(page == 0 ? png : jpeg, image.Data);
            Assert.Equal(page == 0 ? "image/png" : "image/jpeg", image.MediaType);
            Assert.Equal(0, image.XMillimeters); Assert.Equal(7.5, image.YMillimeters);
            Assert.Equal(30, image.WidthMillimeters); Assert.Equal(15, image.HeightMillimeters);
            output.Position = 0; using var raster = new MemoryStream();
            await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 4 }).ConvertAsync(output, raster, page);
            using var pixels = Image.Load<Rgb24>(raster.ToArray());
            Assert.Equal(new Rgb24(255, 255, 255), pixels[60, 10]);
            Assert.True(pixels[20, 40].R > 220); Assert.True(pixels[100, 40].B > 220);
            Assert.True(pixels[20, 80].R > 220); Assert.True(pixels[100, 80].G > 220);
        }
        using var natural = new MemoryStream(); using var original = new MemoryStream(png);
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = 2 }).ConvertAsync(original, natural);
        natural.Position = 0; var naturalPage = Assert.Single((await new OfdReader().ReadAsync(natural)).Pages);
        Assert.Equal(40, naturalPage.WidthMillimeters); Assert.Equal(20, naturalPage.HeightMillimeters);
        using var largePage = new MemoryStream(); original.Position = 0;
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = 2, PageSize = new OfdPageSize { WidthMillimeters = 100, HeightMillimeters = 100 } }).ConvertAsync(original, largePage);
        largePage.Position = 0; var centered = Assert.IsType<OfdImageElement>(Assert.Single(Assert.Single((await new OfdReader().ReadAsync(largePage)).Pages).Elements));
        Assert.Equal(40, centered.WidthMillimeters); Assert.Equal(30, centered.XMillimeters); Assert.Equal(40, centered.YMillimeters);
    }

    [Fact]
    public async Task Import_IdenticalImagesAreDeduplicated()
    {
        using var a = new MemoryStream(ImageBytes()); using var b = new MemoryStream(ImageBytes()); using var output = new MemoryStream();
        await new ImageToOfdConverter().ConvertAsync([a, b], output); output.Position = 0;
        using var zip = new ZipArchive(output, ZipArchiveMode.Read);
        Assert.Single(zip.Entries, entry => entry.Name.EndsWith(".png"));
    }

    [Theory]
    [InlineData(-1)] [InlineData(2)] [InlineData(int.MaxValue)]
    public async Task InvalidPage_LeavesExistingOutputUntouched(int page)
    {
        using var input = new MemoryStream(await Source()); using var output = Sentinel();
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => new OfdToImageConverter().ConvertAsync(input, output, page)); AssertSentinel(output);
    }

    [Theory]
    [InlineData("pixels")] [InlineData("working")] [InlineData("pdf")] [InlineData("output")] [InlineData("entries")] [InlineData("input")]
    public async Task Export_BudgetsFailBeforePublication(string budget)
    {
        var options = new OfdToImageOptions();
        switch (budget)
        {
            case "pixels": options.MaxPixels = 10; break;
            case "working": options.MaxRasterWorkingBytes = 16; break;
            case "pdf": options.MaxIntermediatePdfBytes = 10; break;
            case "output": options.MaxOutputBytes = 10; break;
            case "entries": options.PackageLoadOptions.MaxEntryCount = 1; break;
            case "input": options.PackageLoadOptions.MaxInputBytes = 1; break;
        }
        using var input = new MemoryStream(await Source()); using var output = Sentinel();
        await Assert.ThrowsAnyAsync<InvalidDataException>(() => new OfdToImageConverter(options).ConvertAsync(input, output)); AssertSentinel(output);
    }

    [Theory]
    [InlineData("pixels")] [InlineData("working")] [InlineData("input")] [InlineData("total")] [InlineData("output")]
    public async Task Import_BudgetsFailBeforePublication(string budget)
    {
        var options = new ImageToOfdOptions();
        switch (budget)
        {
            case "pixels": options.MaxPixelsPerImage = 10; break;
            case "working": options.MaxRasterWorkingBytes = 16; break;
            case "input": options.MaxInputBytesPerImage = 1; break;
            case "total": options.MaxTotalInputBytes = 1; break;
            case "output": options.MaxOutputBytes = 10; break;
        }
        using var input = new MemoryStream(ImageBytes()); using var output = Sentinel();
        await Assert.ThrowsAnyAsync<InvalidDataException>(() => new ImageToOfdConverter(options).ConvertAsync(input, output)); AssertSentinel(output);
    }

    [Theory]
    [InlineData(0)] [InlineData(1)] [InlineData(2)]
    public async Task Import_BadLaterImageLeavesNoPartialPackage(int invalidKind)
    {
        byte[] invalid;
        if (invalidKind == 0) invalid = [1, 2, 3];
        else if (invalidKind == 1) invalid = ImageBytes()[..40]; // valid header, missing encoded pixel data
        else { using var gif = new MemoryStream(); using var image = new Image<Rgb24>(2, 2); image.SaveAsGif(gif); invalid = gif.ToArray(); }
        using var valid = new MemoryStream(ImageBytes()); using var bad = new MemoryStream(invalid); using var output = Sentinel();
        await Assert.ThrowsAnyAsync<Exception>(() => new ImageToOfdConverter().ConvertAsync([valid, bad], output)); AssertSentinel(output);
    }

    [Fact]
    public async Task CountsAndCancellation_AreValidatedBeforeOutput()
    {
        using var a = new MemoryStream(ImageBytes()); using var b = new MemoryStream(ImageBytes()); using var output = Sentinel();
        await Assert.ThrowsAsync<ArgumentException>(() => new ImageToOfdConverter().ConvertAsync(Array.Empty<Stream>(), output));
        await Assert.ThrowsAsync<ArgumentException>(() => new ImageToOfdConverter(new ImageToOfdOptions { MaxPageCount = 1 }).ConvertAsync([a, b], output));
        await Assert.ThrowsAsync<ArgumentException>(() => new ImageToOfdConverter(new ImageToOfdOptions { MaxEntryCount = 1 }).ConvertAsync(a, output));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new ImageToOfdConverter().ConvertAsync(a, output, cancellation.Token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new OfdToImageConverter().ConvertAsync(a, output, cancellationToken: cancellation.Token)); AssertSentinel(output);
    }

    [Theory]
    [InlineData(double.NaN)] [InlineData(double.PositiveInfinity)] [InlineData(0)] [InlineData(-1)]
    public void ResolutionAndPageOptionsMustBeFiniteAndPositive(double value)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = value }));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = value }));
        Assert.Throws<InvalidDataException>(() => new ImageToOfdConverter(new ImageToOfdOptions { PageSize = new OfdPageSize { WidthMillimeters = value, HeightMillimeters = 30 } }));
    }

    [Fact]
    public void EncodingAndQualityAreValidated()
    {
        Assert.Throws<ArgumentException>(() => new OfdToImageConverter(new OfdToImageOptions { Format = (OfdImageFormat)9 }));
        Assert.Throws<ArgumentException>(() => new OfdToImageConverter(new OfdToImageOptions { JpegQuality = 101 }));
    }

    [Theory]
    [InlineData(0.5)] [InlineData(0.01)]
    public async Task TinyPageAtLowResolutionUsesOnePixelCeilGrid(double ppm)
    {
        using var pixel = new Image<Rgb24>(1, 1, new Rgb24(220, 40, 20));
        using var png = new MemoryStream(); pixel.SaveAsPng(png); png.Position = 0;
        using var ofd = new MemoryStream();
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = 1 }).ConvertAsync(png, ofd);
        ofd.Position = 0; using var output = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = ppm }).ConvertAsync(ofd, output);
        using var image = Image.Load<Rgb24>(output.ToArray()); Assert.Equal(1, image.Width); Assert.Equal(1, image.Height);
        Assert.True(image[0, 0].R > 180);
    }

    [Fact]
    public async Task ApngAnimationChunkIsRejectedBeforeDecode()
    {
        var png = ImageBytes();
        // Insert acTL after IHDR. ImageSharp 2.x would ignore this ancillary chunk and report one frame.
        var control = new byte[] { 0,0,0,8, 97,99,84,76, 0,0,0,2, 0,0,0,0, 0,0,0,0 };
        var encoded = png[..33].Concat(control).Concat(png[33..]).ToArray();
        using var input = new MemoryStream(encoded); using var output = Sentinel();
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => new ImageToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("Animated PNG", exception.Message); AssertSentinel(output);
    }

    [Fact]
    public async Task SelectedSignatureNestedOfdInheritsEntryBudgetButOtherPagesAreNotDecoded()
    {
        var nested = new OfdDocumentPackage(); nested.Pages.Add(new OfdPage { WidthMillimeters = 10, HeightMillimeters = 10 });
        for (var i = 0; i < 30; i++) nested.PreservedEntries[$"extensions/{i}.bin"] = [1, 2, 3];
        using var appearance = new MemoryStream(); await new OfdPackageWriter().WriteAsync(nested, appearance);
        using var outer = new MemoryStream(); outer.Write(await Source());
        int count;
        using (var zip = new ZipArchive(outer, ZipArchiveMode.Update, true))
        {
            var rootEntry = zip.GetEntry("OFD.xml")!;
            System.Xml.Linq.XDocument root;
            using (var stream = rootEntry.Open()) root = System.Xml.Linq.XDocument.Load(stream);
            var body = root.Descendants().Single(element => element.Name.LocalName == "DocBody");
            var documentPath = body.Elements().Single(element => element.Name.LocalName == "DocRoot").Value;
            var prefix = documentPath[..documentPath.LastIndexOf('/')];
            System.Xml.Linq.XDocument document;
            using (var stream = zip.GetEntry(documentPath)!.Open()) document = System.Xml.Linq.XDocument.Load(stream);
            var pageId = document.Descendants().First(element => element.Name.LocalName == "Page").Attribute("ID")!.Value;
            var signaturePath = prefix + "/Signs/Signatures.xml";
            body.Add(new System.Xml.Linq.XElement(body.Name.Namespace + "Signatures", signaturePath));
            rootEntry.Delete(); using (var stream = zip.CreateEntry("OFD.xml").Open()) root.Save(stream);
            void Add(string name, byte[] data) { using var stream = zip.CreateEntry(name).Open(); stream.Write(data); }
            Add(signaturePath, System.Text.Encoding.UTF8.GetBytes("<Signatures><Signature BaseLoc=\"S/Signature.xml\"/></Signatures>"));
            Add(prefix + "/Signs/S/Signature.xml", System.Text.Encoding.UTF8.GetBytes($"<Signature><SignedInfo><Seal BaseLoc=\"Seal.ofd\"/><StampAnnot PageRef=\"{pageId}\" Boundary=\"0 0 10 10\"/></SignedInfo></Signature>"));
            Add(prefix + "/Signs/S/Seal.ofd", appearance.ToArray());
            count = zip.Entries.Count;
        }
        var options = new OfdToImageOptions { PixelsPerMillimeter = 1, PackageLoadOptions = new() { MaxEntryCount = count } };
        outer.Position = 0; using var output = Sentinel();
        await Assert.ThrowsAnyAsync<InvalidDataException>(() => new OfdToImageConverter(options).ConvertAsync(outer, output, 0)); AssertSentinel(output);
        outer.Position = 0; using var other = new MemoryStream();
        await new OfdToImageConverter(options).ConvertAsync(outer, other, 1); Assert.True(other.Length > 0);
    }

    [Fact]
    public async Task SaveReviewEvidence_WhenRequested()
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_IMAGE_EVIDENCE");
        if (string.IsNullOrWhiteSpace(directory)) return;
        Directory.CreateDirectory(directory);
        var source = await Source(); await File.WriteAllBytesAsync(Path.Combine(directory, "source.ofd"), source);
        using (var input = new MemoryStream(source))
        using (var pdf = File.Create(Path.Combine(directory, "source.pdf"))) await new OfdToPdfConverter().ConvertAsync(input, pdf);
        for (var page = 0; page < 2; page++)
        {
            using var input = new MemoryStream(source);
            using var output = File.Create(Path.Combine(directory, $"export-{page + 1}.png"));
            await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 4 }).ConvertAsync(input, output, page);
        }
        var png = ImageBytes(); var jpeg = ImageBytes(true);
        await File.WriteAllBytesAsync(Path.Combine(directory, "input.png"), png); await File.WriteAllBytesAsync(Path.Combine(directory, "input.jpg"), jpeg);
        using var a = new MemoryStream(png); using var b = new MemoryStream(jpeg); using var imported = new MemoryStream();
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = 2,
            PageSize = new OfdPageSize { WidthMillimeters = 60, HeightMillimeters = 50 } }).ConvertAsync([a, b], imported);
        await File.WriteAllBytesAsync(Path.Combine(directory, "imported.ofd"), imported.ToArray());
        imported.Position = 0; using (var pdf = File.Create(Path.Combine(directory, "imported.pdf"))) await new OfdToPdfConverter().ConvertAsync(imported, pdf);
        for (var page = 0; page < 2; page++)
        {
            imported.Position = 0; using var image = File.Create(Path.Combine(directory, $"imported-{page + 1}.png"));
            await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 4 }).ConvertAsync(imported, image, page);
        }
    }

    private static MemoryStream Sentinel() { var output = new MemoryStream(); output.Write([4, 5, 6]); return output; }
    private static void AssertSentinel(MemoryStream output) { Assert.Equal(new byte[] { 4, 5, 6 }, output.ToArray()); Assert.Equal(3, output.Position); }
}
