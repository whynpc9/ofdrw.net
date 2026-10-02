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

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task SelectedSignatureNestedOfdFailsClosedButOtherPagesAreNotDecoded(bool malformed)
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
            Add(prefix + "/Signs/S/Seal.ofd", malformed ? new byte[] { 0x50, 0x4b, 3, 4, 1 } : appearance.ToArray());
            count = zip.Entries.Count;
        }
        var options = new OfdToImageOptions { PixelsPerMillimeter = 1, PackageLoadOptions = new() { MaxEntryCount = count } };
        outer.Position = 0; using var output = Sentinel();
        await Assert.ThrowsAnyAsync<InvalidDataException>(() => new OfdToImageConverter(options).ConvertAsync(outer, output, 0)); AssertSentinel(output);
        outer.Position = 0; using var other = new MemoryStream();
        await new OfdToImageConverter(options).ConvertAsync(outer, other, 1); Assert.True(other.Length > 0);
    }

    [Fact]
    public async Task IntermediatePdfWriteHandleIsClosedBeforeExternalReaderOpens()
    {
        string path;
        using (var staged = new ImageIoStagingStream(100))
        {
            path = staged.PathOnDisk;
            staged.Write([1, 2, 3]);
            await staged.CloseWriterAsync(CancellationToken.None);
            // FileShare.None proves the original handle was released; this also runs in Windows CI.
            using var external = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None);
            Assert.Equal(3, external.Length);
        }
        Assert.False(File.Exists(path));
    }

    [Theory]
    [InlineData(0.001)] [InlineData(double.Epsilon)]
    public async Task FixedPageShrinksOversizeNaturalImageWithoutOverflow(double ppm)
    {
        using var input = new MemoryStream(ImageBytes()); using var output = new MemoryStream();
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = ppm,
            PageSize = new OfdPageSize { WidthMillimeters = 30, HeightMillimeters = 30 } }).ConvertAsync(input, output);
        output.Position = 0;
        var page = Assert.Single((await new OfdReader().ReadAsync(output)).Pages);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(page.Elements));
        Assert.Equal(30, image.WidthMillimeters); Assert.Equal(15, image.HeightMillimeters);
        Assert.Equal(0, image.XMillimeters); Assert.Equal(7.5, image.YMillimeters);
        input.Position = 0; using var natural = Sentinel();
        await Assert.ThrowsAsync<InvalidDataException>(() => new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = ppm }).ConvertAsync(input, natural));
        AssertSentinel(natural);
    }

    [Theory]
    [InlineData(145, 10000, 0.01)] [InlineData(39, 210, 0.001)]
    public async Task NonDyadicFixedPageScalingNeverOvershootsOrProducesNegativeOrigins(int pixels, double millimeters, double ppm)
    {
        using var image = new Image<Rgb24>(pixels, pixels, new Rgb24(220, 40, 20));
        using var input = new MemoryStream(); image.SaveAsPng(input); input.Position = 0; using var output = new MemoryStream();
        await new ImageToOfdConverter(new ImageToOfdOptions { PixelsPerMillimeter = ppm,
            PageSize = new OfdPageSize { WidthMillimeters = millimeters, HeightMillimeters = millimeters } }).ConvertAsync(input, output);
        output.Position = 0; var page = Assert.Single((await new OfdReader().ReadAsync(output)).Pages);
        var placed = Assert.IsType<OfdImageElement>(Assert.Single(page.Elements));
        Assert.Equal(millimeters, placed.WidthMillimeters); Assert.Equal(millimeters, placed.HeightMillimeters);
        Assert.Equal(0, placed.XMillimeters); Assert.Equal(0, placed.YMillimeters);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task SubMicrometerImportedGeometryFailsBeforePublication(bool fixedPage)
    {
        using var image = new Image<Rgb24>(1, 1); using var input = new MemoryStream(); image.SaveAsPng(input); input.Position = 0;
        using var output = Sentinel();
        await Assert.ThrowsAsync<InvalidDataException>(() => new ImageToOfdConverter(new ImageToOfdOptions
        { PixelsPerMillimeter = 3000, PageSize = fixedPage ? new OfdPageSize { WidthMillimeters = 30, HeightMillimeters = 30 } : null }).ConvertAsync(input, output));
        AssertSentinel(output);
        Assert.Throws<InvalidDataException>(() => new ImageToOfdConverter(new ImageToOfdOptions
        { PageSize = new OfdPageSize { WidthMillimeters = 0.0001, HeightMillimeters = 10 } }));
        // Fitting a very thin input must also validate its final, shrunken short axis.
        using var thin = new Image<Rgb24>(2000, 1); using var thinInput = new MemoryStream(); thin.SaveAsPng(thinInput); thinInput.Position = 0;
        await Assert.ThrowsAsync<InvalidDataException>(() => new ImageToOfdConverter(new ImageToOfdOptions
        { PixelsPerMillimeter = 1, PageSize = new OfdPageSize { WidthMillimeters = 1, HeightMillimeters = 1 } }).ConvertAsync(thinInput, output));
        AssertSentinel(output);
    }

    internal static async Task<byte[]> NestedSeal(bool blue = false)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 10, HeightMillimeters = 10 };
        page.Elements.Add(new OfdPathElement { WidthMillimeters = 10, HeightMillimeters = 10,
            AbbreviatedData = "M 0 0 L 10 0 L 10 10 L 0 10 C", Fill = true, Stroke = false,
            FillColor = blue ? new OfdColor(30, 80, 220) : new OfdColor(220, 40, 20) });
        package.Pages.Add(page);
        package.PreservedEntries["extensions/budget.bin"] = Enumerable.Range(0, 20000).Select(i => (byte)i).ToArray();
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }

    internal static async Task<byte[]> WithSeals(byte[][] payloads, int stampsPerRecord = 1, bool invalidBoundary = false, int pageCount = 2)
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 120, HeightMillimeters = 80 });
        if (pageCount == 2) package.Pages.Add(new OfdPage { Index = 1, WidthMillimeters = 120, HeightMillimeters = 80 });
        using var outer = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, outer);
        using (var zip = new ZipArchive(outer, ZipArchiveMode.Update, true))
        {
            var rootEntry = zip.GetEntry("OFD.xml")!; System.Xml.Linq.XDocument root;
            using (var stream = rootEntry.Open()) root = System.Xml.Linq.XDocument.Load(stream);
            var body = root.Descendants().Single(element => element.Name.LocalName == "DocBody");
            body.Add(new System.Xml.Linq.XElement(body.Name.Namespace + "Signatures", "Doc_0/Signs/Signatures.xml"));
            rootEntry.Delete(); using (var stream = zip.CreateEntry("OFD.xml").Open()) root.Save(stream);
            System.Xml.Linq.XDocument document;
            using (var stream = zip.GetEntry("Doc_0/Document.xml")!.Open()) document = System.Xml.Linq.XDocument.Load(stream);
            var pageId = document.Descendants().First(element => element.Name.LocalName == "Page").Attribute("ID")!.Value;
            void Add(string name, string text) { using var stream = zip.CreateEntry(name).Open(); stream.Write(System.Text.Encoding.UTF8.GetBytes(text)); }
            var records = new List<string>(); var stored = new Dictionary<string, string>();
            for (var index = 0; index < payloads.Length; index++)
            {
                var hash = Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(payloads[index]));
                if (!stored.TryGetValue(hash, out var file))
                {
                    file = $"Seal-{stored.Count}.dat"; stored.Add(hash, file);
                    using var stream = zip.CreateEntry("Doc_0/Signs/" + file).Open(); stream.Write(payloads[index]);
                }
                records.Add($"<Signature BaseLoc=\"S{index}/Signature.xml\"/>");
                var stamps = string.Concat(Enumerable.Range(0, stampsPerRecord).Select(stamp =>
                    $"<StampAnnot PageRef=\"{pageId}\" Boundary=\"{5 + (stamp % 6) * 18} {10 + (stamp / 6) * 18} {(invalidBoundary ? 0 : 10)} 10\"/>"));
                Add($"Doc_0/Signs/S{index}/Signature.xml", $"<Signature><SignedInfo><Seal BaseLoc=\"../{file}\"/>{stamps}</SignedInfo></Signature>");
            }
            Add("Doc_0/Signs/Signatures.xml", "<Signatures>" + string.Concat(records) + "</Signatures>");
        }
        return outer.ToArray();
    }

    [Fact]
    public async Task SharedNestedSealUsesOneExpansionBudgetAndOnePdfForm()
    {
        var seal = await NestedSeal(); var bytes = await WithSeals([seal, seal], stampsPerRecord: 12);
        using var outer = new MemoryStream(bytes); var package = await new OfdReader().ReadAsync(outer);
        using var nested = new MemoryStream(seal); var nestedPackage = await new OfdReader().ReadAsync(nested);
        var budget = package.PreservedEntries.Values.Sum(data => (long)data.Length) + nestedPackage.PreservedEntries.Values.Sum(data => (long)data.Length);
        outer.Position = 0; using var output = new MemoryStream();
        await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2,
            PackageLoadOptions = new() { MaxTotalUncompressedBytes = budget } }).ConvertAsync(outer, output);
        using var image = Image.Load<Rgb24>(output.ToArray());
        Assert.True(image[20, 30].R > 180); Assert.True(image[56, 66].R > 180);
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter(new OfdToPdfOptions { PackageLoadOptions = new() { MaxTotalUncompressedBytes = budget } })
            .ConvertPackageAsync(package, pdf, [0], CancellationToken.None, strictAppearanceBudgets: true, maximumSignatureAppearances: 1000);
        var forms = System.Text.Encoding.ASCII.GetString(pdf.ToArray()).Split("/Subtype /Form", StringSplitOptions.None).Length - 1;
        Assert.Equal(1, forms); // Repeated drawing also reuses the PDF form, rather than multiplying PDF state.
    }

    [Theory]
    [InlineData("bytes")] [InlineData("entries")] [InlineData("pages")]
    public async Task DistinctNestedSealsShareCumulativeExpansionBudgets(string kind)
    {
        var seal = await NestedSeal(); var bytes = await WithSeals([seal, await NestedSeal(true)]);
        using var input = new MemoryStream(bytes); var package = await new OfdReader().ReadAsync(input);
        using var nested = new MemoryStream(seal); var nestedPackage = await new OfdReader().ReadAsync(nested);
        var load = new Ofdrw.Net.Packaging.Archive.OfdPackageLoadOptions();
        if (kind == "bytes") load.MaxTotalUncompressedBytes = package.PreservedEntries.Values.Sum(data => (long)data.Length) + nestedPackage.PreservedEntries.Values.Sum(data => (long)data.Length);
        if (kind == "entries") load.MaxEntryCount = package.PreservedEntries.Count + nestedPackage.PreservedEntries.Count;
        if (kind == "pages") load.MaxPageCount = package.Pages.Count + nestedPackage.Pages.Count;
        input.Position = 0; using var output = Sentinel();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2, PackageLoadOptions = load }).ConvertAsync(input, output));
        AssertSentinel(output);
    }

    [Fact]
    public async Task SharedAsn1PayloadIsExtractedOnceAndInvalidCandidatesStillConsumeGlobalStampBudget()
    {
        var seal = await NestedSeal();
        var wrapped = new byte[] { 4, 0x82, (byte)(seal.Length >> 8), (byte)seal.Length }.Concat(seal).ToArray();
        var bytes = await WithSeals([wrapped, wrapped]); using var input = new MemoryStream(bytes);
        var package = await new OfdReader().ReadAsync(input);
        var appearances = OfdSignatureAppearanceReader.Read(package, new HashSet<string> { package.Pages[0].Id! }, 1000);
        Assert.Equal(2, appearances.Count); Assert.Same(appearances[0].Data, appearances[1].Data); Assert.Equal(seal, appearances[0].Data);
        var bad = await WithSeals([ImageBytes()], stampsPerRecord: 2, invalidBoundary: true);
        using var invalid = new MemoryStream(bad); using var output = Sentinel();
        var rejection = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions
        { PixelsPerMillimeter = 2, MaxSignatureAppearanceCount = 1 }).ConvertAsync(invalid, output));
        Assert.Contains("Selected signature appearance count", rejection.Message); AssertSentinel(output);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task ForgedAsn1LengthFailsBeforeAllocationAndPublication(bool supportedPrefix)
    {
        var bad = new byte[] { 4, 0x84, 0x7f, 0xff, 0xff, 0xff };
        if (supportedPrefix) bad = bad.Concat(new byte[] { 0x50, 0x4b, 3, 4 }).ToArray();
        using var input = new MemoryStream(await WithSeals([bad])); using var output = Sentinel();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions
        { PixelsPerMillimeter = 2 }).ConvertAsync(input, output));
        Assert.Contains("ASN.1", error.Message); AssertSentinel(output);
        // The existing public PDF path still ignores unreadable vendor appearance data.
        input.Position = 0; using var legacy = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, legacy);
        Assert.True(legacy.Length > 0);
    }

    [Fact]
    public void BitmapResourceCacheEvictsBeforeAllocatingAndRetainsOnlyOneDecodedResource()
    {
        var live = 0; var peak = 0; var made = 0;
        using var cache = new SinglePayloadResource<TrackedResource>(_ =>
        {
            live++; made++; peak = Math.Max(peak, live);
            return new TrackedResource(() => live--);
        });
        var repeated = new byte[] { 1 };
        var first = cache.Get(repeated); Assert.Same(first, cache.Get(repeated)); Assert.Equal(1, made);
        for (var index = 0; index < 500; index++) cache.Get(new byte[] { (byte)index });
        Assert.Equal(1, peak); Assert.Equal(1, live); Assert.Equal(501, made);
        cache.Dispose(); Assert.Equal(0, live);
    }

    private sealed class TrackedResource(Action dispose) : IDisposable
    {
        private bool _disposed;
        public void Dispose() { if (!_disposed) { _disposed = true; dispose(); } }
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task LegacyPdfKeepsValidSealBeforeMalformedAsn1(bool sameRecord)
    {
        var jpeg = ImageBytes(true); var bad = new byte[] { 4, 0x84, 0x7f, 0xff, 0xff, 0xff };
        byte[] Wrap(byte tag, byte[] content) => new byte[] { tag, 0x82, (byte)(content.Length >> 8), (byte)content.Length }.Concat(content).ToArray();
        var encoded = sameRecord ? await WithSeals([Wrap(0x30, Wrap(4, jpeg).Concat(bad).ToArray())]) : await WithSeals([jpeg, bad]);
        using var input = new MemoryStream(encoded); var package = await new OfdReader().ReadAsync(input);
        Assert.Equal(jpeg, Assert.Single(OfdSignatureAppearanceReader.Read(package)).Data);
        input.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = reader.GetPageReader(0); var pixels = page.GetImage();
        var offset = ((int)(12 * 72d / 25.4d * 2) * page.GetPageWidth() + (int)(7 * 72d / 25.4d * 2)) * 4;
        Assert.True(pixels[offset + 2] > 180); Assert.True(pixels[offset] < 60); // The valid JPEG's red quadrant still paints.
        input.Position = 0; using var strict = Sentinel();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2 }).ConvertAsync(input, strict));
        AssertSentinel(strict);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void RealPdfImageTableRetainsNoDecodedPixelsAfterRealization(bool jpeg)
    {
        using var document = new PdfSharpCore.Pdf.PdfDocument(); var page = document.AddPage();
        using (var graphics = PdfSharpCore.Drawing.XGraphics.FromPdfPage(page))
        {
            for (var index = 0; index < 12; index++)
            {
                using var image = new Image<Rgba32>(32, 32, new Rgba32((byte)(index * 20), 30, 220));
                using var encoded = new MemoryStream();
                if (jpeg) image.SaveAsJpeg(encoded); else image.SaveAsPng(encoded);
                Image<Rgba32>? actualPixels = null;
                var source = new EncodedBitmapImageSource(encoded.ToArray(), 1024, CancellationToken.None, data =>
                { actualPixels = Image.Load<Rgba32>(data); return actualPixels; });
                using var ximage = PdfSharpCore.Drawing.XImage.FromImageSource(source);
                graphics.DrawImage(ximage, index * 5, 0, 5, 5);
                Assert.NotNull(actualPixels);
                Assert.Throws<ObjectDisposedException>(() => { _ = actualPixels![0, 0]; });
                // PdfImageTable still owns ximage; cached dimensions remain usable without decoded buffers.
                Assert.Equal(32, ximage.PixelWidth); Assert.Equal(32, ximage.PixelHeight);
            }
        }
        using var output = new MemoryStream(); document.Save(output, false);
        Assert.True(output.Length > 0);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(output.ToArray());
        Assert.Single(pdf.GetPages());
    }

    [Theory]
    [InlineData("xml")] [InlineData("missing")] [InlineData("bitmap")]
    public async Task StrictSelectedSealParseAndDecodeFailuresNeverPublish(string kind)
    {
        byte[] bad;
        if (kind == "bitmap") bad = ImageBytes()[..40];
        else
        {
            using var buffer = new MemoryStream();
            using (var zip = new ZipArchive(buffer, ZipArchiveMode.Create, true))
            {
                using var entry = zip.CreateEntry(kind == "xml" ? "OFD.xml" : "missing-root.txt").Open();
                entry.Write(System.Text.Encoding.UTF8.GetBytes(kind == "xml" ? "<OFD" : "no OFD root"));
            }
            bad = buffer.ToArray();
        }
        using var input = new MemoryStream(await WithSeals([ImageBytes(true), bad])); using var output = Sentinel();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2 }).ConvertAsync(input, output));
        AssertSentinel(output);
        // The public PDF converter remains tolerant of the bad second vendor appearance.
        input.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = reader.GetPageReader(0); var pixels = page.GetImage();
        var offset = ((int)(12 * 72d / 25.4d * 2) * page.GetPageWidth() + (int)(7 * 72d / 25.4d * 2)) * 4;
        Assert.True(pixels[offset + 2] > 180); Assert.True(pixels[offset] < 60);
    }

    [Theory]
    [InlineData("missing")] [InlineData("empty")] [InlineData("unsupported")]
    [InlineData("Infinity")] [InlineData("NaN")]
    [InlineData("1e308")] [InlineData("1e-5")]
    [InlineData("origin-NaN")] [InlineData("origin-Infinity")] [InlineData("origin-1e308")]
    public async Task StrictSelectedStampRequiresPayloadAndFiniteNestedPage(string kind)
    {
        byte[] bad;
        if (kind is "Infinity" or "NaN" or "1e308" or "1e-5" || kind.StartsWith("origin-"))
        {
            using var rewritten = new MemoryStream(); rewritten.Write(await NestedSeal());
            using (var zip = new ZipArchive(rewritten, ZipArchiveMode.Update, true))
            {
                foreach (var path in new[] { "Doc_0/Document.xml", "Doc_0/Pages/Page_0/Content.xml" })
                {
                    var entry = zip.GetEntry(path)!; System.Xml.Linq.XDocument xml;
                    using (var stream = entry.Open()) xml = System.Xml.Linq.XDocument.Load(stream);
                    foreach (var box in xml.Descendants().Where(e => e.Name.LocalName == "PhysicalBox"))
                        box.Value = kind.StartsWith("origin-") ? $"{kind.Substring(7)} 0 10 10" : $"0 0 {kind} 210";
                    entry.Delete(); using var output = zip.CreateEntry(path).Open(); xml.Save(output);
                }
            }
            bad = rewritten.ToArray();
        }
        else bad = kind == "unsupported" ? new byte[] { 4, 1, 0 } : Array.Empty<byte>();
        var source = await WithSeals([ImageBytes(true), bad]);
        if (kind == "missing")
        {
            using var rewritten = new MemoryStream(); rewritten.Write(source);
            using (var zip = new ZipArchive(rewritten, ZipArchiveMode.Update, true)) zip.GetEntry("Doc_0/Signs/Seal-1.dat")!.Delete();
            source = rewritten.ToArray();
        }
        using var input = new MemoryStream(source); using var sentinel = Sentinel();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2 }).ConvertAsync(input, sentinel));
        AssertSentinel(sentinel);
        input.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = reader.GetPageReader(0); var pixels = page.GetImage();
        var offset = ((int)(12 * 72d / 25.4d * 2) * page.GetPageWidth() + (int)(7 * 72d / 25.4d * 2)) * 4;
        Assert.True(pixels[offset + 2] > 180); Assert.True(pixels[offset] < 60);
    }

    [Theory]
    [InlineData(null)] [InlineData("")] [InlineData("invalid")] [InlineData("0 0 10")]
    [InlineData("0 0 0 10")] [InlineData("0 0 10 0")]
    [InlineData("0 0 -10 10")] [InlineData("0 0 10 -10")]
    [InlineData("NaN 0 10 10")] [InlineData("0 NaN 10 10")]
    [InlineData("0 0 NaN 10")] [InlineData("0 0 10 NaN")]
    [InlineData("Infinity 0 10 10")] [InlineData("0 -Infinity 10 10")]
    [InlineData("0 0 Infinity 10")] [InlineData("0 0 10 -Infinity")]
    [InlineData("1e309 0 10 10")] [InlineData("0 1e309 10 10")]
    [InlineData("0 0 1e309 10")] [InlineData("0 0 10 1e309")]
    [InlineData("1e308 0 10 10")] [InlineData("0 -1e308 10 10")]
    [InlineData("0 0 1e308 10")] [InlineData("0 0 10 1e308")]
    [InlineData("0 6.3e307 10 6.3e307")]
    [InlineData("0 0 1e-5 10")] [InlineData("0 0 10 1e-5")]
    public async Task StrictSelectedStampRejectsInvalidBoundaryAndLegacyPdfKeepsValidSeal(string? boundary)
    {
        using var rewritten = new MemoryStream(); rewritten.Write(await WithSeals([ImageBytes(true), ImageBytes(true)]));
        using (var zip = new ZipArchive(rewritten, ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("Doc_0/Signs/S1/Signature.xml")!; System.Xml.Linq.XDocument xml;
            using (var stream = entry.Open()) xml = System.Xml.Linq.XDocument.Load(stream);
            xml.Descendants().Single(e => e.Name.LocalName == "StampAnnot").SetAttributeValue("Boundary", boundary);
            entry.Delete(); using var output = zip.CreateEntry("Doc_0/Signs/S1/Signature.xml").Open(); xml.Save(output);
        }
        using var input = new MemoryStream(rewritten.ToArray()); using var sentinel = Sentinel();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions
        { PixelsPerMillimeter = 2 }).ConvertAsync(input, sentinel));
        Assert.Contains("Boundary", error.Message); AssertSentinel(sentinel);
        input.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = reader.GetPageReader(0); var pixels = page.GetImage();
        var offset = ((int)(12 * 72d / 25.4d * 2) * page.GetPageWidth() + (int)(7 * 72d / 25.4d * 2)) * 4;
        Assert.True(pixels[offset + 2] > 180); Assert.True(pixels[offset] < 60);
        pdf.Position = 0;
        using var parsed = PdfSharpCore.Pdf.IO.PdfReader.Open(pdf, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var content = System.Text.Encoding.ASCII.GetString(parsed.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        Assert.DoesNotContain("Infinity", content); Assert.DoesNotContain("NaN", content);
    }

    [Theory]
    [InlineData("NaN", false)] [InlineData("NaN", true)]
    [InlineData("Infinity", false)] [InlineData("-Infinity", true)]
    [InlineData("1e308", false)] [InlineData("-1e308", true)]
    public async Task StrictSelectedPageRejectsNonRepresentableOrigin(string origin, bool yAxis)
    {
        using var rewritten = new MemoryStream(); rewritten.Write(await WithSeals([ImageBytes(true)]));
        using (var zip = new ZipArchive(rewritten, ZipArchiveMode.Update, true))
        {
            foreach (var path in new[] { "Doc_0/Document.xml", "Doc_0/Pages/Page_0/Content.xml" })
            {
                var entry = zip.GetEntry(path)!; System.Xml.Linq.XDocument xml;
                using (var stream = entry.Open()) xml = System.Xml.Linq.XDocument.Load(stream);
                foreach (var box in xml.Descendants().Where(e => e.Name.LocalName == "PhysicalBox"))
                    box.Value = yAxis ? $"0 {origin} 120 80" : $"{origin} 0 120 80";
                entry.Delete(); using var output = zip.CreateEntry(path).Open(); xml.Save(output);
            }
        }
        using var input = new MemoryStream(rewritten.ToArray()); using var sentinel = Sentinel();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions
        { PixelsPerMillimeter = 2 }).ConvertAsync(input, sentinel));
        Assert.Contains("origin", error.Message); AssertSentinel(sentinel);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task SelectedStampUsesPageOriginAndRejectsRelativeCoordinateOverflow(bool overflow)
    {
        using var rewritten = new MemoryStream(); rewritten.Write(await WithSeals([ImageBytes(true)]));
        using (var zip = new ZipArchive(rewritten, ZipArchiveMode.Update, true))
        {
            foreach (var path in new[] { "Doc_0/Document.xml", "Doc_0/Pages/Page_0/Content.xml", "Doc_0/Signs/S0/Signature.xml" })
            {
                var entry = zip.GetEntry(path)!; System.Xml.Linq.XDocument xml;
                using (var stream = entry.Open()) xml = System.Xml.Linq.XDocument.Load(stream);
                foreach (var box in xml.Descendants().Where(e => e.Name.LocalName == "PhysicalBox"))
                    box.Value = overflow ? "-2e306 0 120 80" : "2 3 120 80";
                if (overflow)
                    foreach (var stamp in xml.Descendants().Where(e => e.Name.LocalName == "StampAnnot"))
                        stamp.SetAttributeValue("Boundary", "2e306 0 10 10");
                entry.Delete(); using var output = zip.CreateEntry(path).Open(); xml.Save(output);
            }
        }
        using var input = new MemoryStream(rewritten.ToArray()); using var outputImage = overflow ? Sentinel() : new MemoryStream();
        var converter = new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 2 });
        if (overflow)
        {
            var error = await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, outputImage));
            Assert.Contains("placement", error.Message); AssertSentinel(outputImage);
            input.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, pdf);
            pdf.Position = 0;
            using var parsed = PdfSharpCore.Pdf.IO.PdfReader.Open(pdf, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
            var content = System.Text.Encoding.ASCII.GetString(parsed.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
            Assert.DoesNotContain("Infinity", content); Assert.DoesNotContain("NaN", content);
        }
        else
        {
            await converter.ConvertAsync(input, outputImage);
            using var raster = Image.Load<Rgb24>(outputImage.ToArray());
            Assert.True(raster[10, 18].R > 180); Assert.True(raster[10, 18].B < 60);
        }
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task StrictNestedStampRejectsOverflowOrZeroSerializedFormScale(bool overflow)
    {
        using var seal = new MemoryStream(); seal.Write(await NestedSeal());
        if (overflow)
        {
            using var zip = new ZipArchive(seal, ZipArchiveMode.Update, true);
            foreach (var path in new[] { "Doc_0/Document.xml", "Doc_0/Pages/Page_0/Content.xml" })
            {
                var entry = zip.GetEntry(path)!; System.Xml.Linq.XDocument xml;
                using (var stream = entry.Open()) xml = System.Xml.Linq.XDocument.Load(stream);
                foreach (var box in xml.Descendants().Where(e => e.Name.LocalName == "PhysicalBox")) box.Value = "0 0 0.001 10";
                entry.Delete(); using var output = zip.CreateEntry(path).Open(); xml.Save(output);
            }
        }
        using var rewritten = new MemoryStream(); rewritten.Write(await WithSeals([ImageBytes(true), seal.ToArray()]));
        using (var zip = new ZipArchive(rewritten, ZipArchiveMode.Update, true))
        {
            var entry = zip.GetEntry("Doc_0/Signs/S1/Signature.xml")!; System.Xml.Linq.XDocument xml;
            using (var stream = entry.Open()) xml = System.Xml.Linq.XDocument.Load(stream);
            xml.Descendants().Single(e => e.Name.LocalName == "StampAnnot").SetAttributeValue("Boundary",
                overflow ? "23 10 1e306 10" : "23 10 0.0001 10");
            entry.Delete(); using var output = zip.CreateEntry("Doc_0/Signs/S1/Signature.xml").Open(); xml.Save(output);
        }
        using var input = new MemoryStream(rewritten.ToArray()); using var sentinel = Sentinel();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToImageConverter(new OfdToImageOptions
        { PixelsPerMillimeter = 2 }).ConvertAsync(input, sentinel));
        Assert.Contains("form scale", error.Message); AssertSentinel(sentinel);
        input.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf.ToArray(), new Docnet.Core.Models.PageDimensions(2d));
        using var page = reader.GetPageReader(0); var pixels = page.GetImage();
        var offset = ((int)(12 * 72d / 25.4d * 2) * page.GetPageWidth() + (int)(7 * 72d / 25.4d * 2)) * 4;
        Assert.True(pixels[offset + 2] > 180); Assert.True(pixels[offset] < 60);
        pdf.Position = 0;
        using var parsed = PdfSharpCore.Pdf.IO.PdfReader.Open(pdf, PdfSharpCore.Pdf.IO.PdfDocumentOpenMode.Import);
        var content = System.Text.Encoding.ASCII.GetString(parsed.Pages[0].Contents.CreateSingleContent().Stream.UnfilteredValue);
        Assert.DoesNotContain("Infinity", content); Assert.DoesNotContain("NaN", content);
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
        // Appearance-only fixture: repeated nested OFD, no cryptographic-validity claim.
        var repeated = await WithSeals([await NestedSeal()], stampsPerRecord: 6, pageCount: 1);
        await File.WriteAllBytesAsync(Path.Combine(directory, "shared-seal.ofd"), repeated);
        using (var input = new MemoryStream(repeated))
        using (var pdf = File.Create(Path.Combine(directory, "shared-seal.pdf"))) await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using (var input = new MemoryStream(repeated))
        using (var pngOutput = File.Create(Path.Combine(directory, "shared-seal.png")))
            await new OfdToImageConverter(new OfdToImageOptions { PixelsPerMillimeter = 4 }).ConvertAsync(input, pngOutput);
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
