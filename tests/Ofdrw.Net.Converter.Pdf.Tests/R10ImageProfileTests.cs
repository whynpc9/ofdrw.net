using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Packaging;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class R10ImageProfileTests
{
    private static readonly byte[] Red = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR42mP4z8AAAAMBAQD3A0FDAAAAAElFTkSuQmCC");
    private static readonly byte[] Black = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQAAAAA3bvkkAAAACklEQVR42mNgAAAAAgAB5Sfe/AAAAABJRU5ErkJggg==");
    private static readonly byte[] White = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABAQAAAAA3bvkkAAAACklEQVR42mNoAAAAggCB2kUIOwAAAABJRU5ErkJggg==");
    private static readonly byte[] NonBinary = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAAAAAA6fptVAAAACklEQVR4nGOoBwAAgQCA05RTSgAAAABJRU5ErkJggg==");
    private static readonly byte[] HalfAlpha = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4z8DQAAAEgQGALFXOsAAAAABJRU5ErkJggg==");
    private static byte[] Jbig2 => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "R10", "owned-generic-region.jb2"));
    private static byte[] Pbm => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "R10", "owned-pattern.pbm"));

    [Fact]
    public async Task OwnedJbig2RoundTripsThroughReaderAndThreeExportersWithPngPixelsMatchingIndependentPbm()
    {
        var encoded = Jbig2;
        var ofd = Package(encoded, "JBIG2");
        using var source = new MemoryStream(ofd);
        var package = await new OfdReader().ReadAsync(source);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[0].Elements));
        Assert.Equal("image/jbig2", image.MediaType);
        Assert.Equal(encoded, image.Data);
        var projected = OfdImageRenderProjection.Prepare(image, 4096);
        Assert.Equal("image/png", projected.MediaType);
        using (var png = Image.Load<Rgb24>(projected.Data))
        {
            Assert.Equal(64, png.Width); Assert.Equal(56, png.Height);
            var pbm = Pbm; Assert.Equal("P4\n64 56\n", Encoding.ASCII.GetString(pbm, 0, 9));
            for (var y = 0; y < 56; y++)
                for (var x = 0; x < 64; x++)
                {
                    var black = (pbm[9 + y * 8 + x / 8] & (0x80 >> (x % 8))) != 0;
                    var pixel = png[x, y];
                    Assert.Equal(black ? (byte)0 : (byte)255, pixel.R);
                    Assert.Equal(pixel.R, pixel.G); Assert.Equal(pixel.R, pixel.B);
                }
        }
        using var pdf = new MemoryStream(); using var svg = new MemoryStream(); using var html = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        Assert.True(pdf.Length > 100);
        using (var emittedSvg = EmbeddedPng(svg)) Assert.Equal(64, emittedSvg.Width);
        using (var emittedHtml = EmbeddedPng(html)) Assert.Equal(56, emittedHtml.Height);
        Assert.DoesNotContain("data:image/jbig2", Encoding.UTF8.GetString(svg.ToArray()));
    }

    [Theory]
    [InlineData(false, 0)]
    [InlineData(true, 255)]
    public async Task ReaderResolvesBinaryMaskAndProjectionAppliesPerPixelAlpha(bool white, int expectedAlpha)
    {
        var ofd = Package(Red, "PNG", white ? White : Black, "PNG");
        using var source = new MemoryStream(ofd);
        var package = await new OfdReader().ReadAsync(source);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[0].Elements));
        Assert.Equal("21", image.MaskResourceId);
        Assert.Equal(Red, image.Data);
        var projected = OfdImageRenderProjection.Prepare(image, 100);
        using var pixels = Image.Load<Rgba32>(projected.Data);
        Assert.Equal(expectedAlpha, pixels[0, 0].A);
        Assert.Equal(255, pixels[0, 0].R);
        var sourceXml = image.SourceXml;
        using var roundtrip = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, roundtrip);
        roundtrip.Position = 0;
        var reopened = await new OfdReader().ReadAsync(roundtrip);
        var retained = Assert.IsType<OfdImageElement>(Assert.Single(reopened.Pages[0].Elements));
        var originalElement = XElement.Parse(sourceXml!);
        var retainedElement = XElement.Parse(retained.SourceXml!);
        Assert.Equal(originalElement.Name, retainedElement.Name);
        Assert.Equal(originalElement.Attribute("ID")?.Value, retainedElement.Attribute("ID")?.Value);
        Assert.Equal(originalElement.Attribute("ResourceID")?.Value, retainedElement.Attribute("ResourceID")?.Value);
        Assert.Equal(originalElement.Attribute("ImageMask")?.Value, retainedElement.Attribute("ImageMask")?.Value);
        Assert.Equal(sourceXml, image.SourceXml);
        Assert.Equal(image.Data, retained.Data);
        Assert.Equal(image.MaskData, retained.MaskData);
        using var svg = new MemoryStream(); using var pdf = new MemoryStream(); using var html = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(ofd), svg);
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(ofd), pdf);
        await new OfdToHtmlConverter().ConvertAsync(new MemoryStream(ofd), html);
        using (var emittedSvg = EmbeddedPng(svg)) Assert.Equal(expectedAlpha, emittedSvg[0, 0].A);
        using (var emittedHtml = EmbeddedPng(html)) Assert.Equal(expectedAlpha, emittedHtml[0, 0].A);
        Assert.True(pdf.Length > 100); Assert.True(html.Length > 100);
    }

    [Fact]
    public async Task MissingMaskAndUnadmittedJbig2FailClosed()
    {
        var missing = Package(Red, "PNG", null, "PNG", maskReference: true);
        using var source = new MemoryStream(missing);
        var package = await new OfdReader().ReadAsync(source);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[0].Elements));
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare(image, 100));
        using var output = new MemoryStream();
        await Assert.ThrowsAnyAsync<Exception>(async () => await new OfdToSvgConverter().ConvertAsync(new MemoryStream(missing), output));
        Assert.Equal(0, output.Length);
        var trailing = Jbig2.Concat(new byte[] { 0 }).ToArray();
        var invalid = Package(trailing, "JBIG2");
        using var input = new MemoryStream(invalid);
        var malformed = await new OfdReader().ReadAsync(input);
        Assert.ThrowsAny<Exception>(() => OfdImageRenderProjection.Prepare((OfdImageElement)malformed.Pages[0].Elements[0], 4096));
    }

    [Fact]
    public async Task Jbig2BoundsAndLiteralMediaIdentityRejectBeforeExport()
    {
        var huge = Jbig2;
        huge[24] = 0x7f; huge[25] = 0xff; huge[26] = 0xff; huge[27] = 0xff;
        var malformed = await new OfdReader().ReadAsync(new MemoryStream(Package(huge, "JBIG2")));
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare((OfdImageElement)malformed.Pages[0].Elements[0], 4096));
        var mislabeled = await new OfdReader().ReadAsync(new MemoryStream(Package(Jbig2, "PNG")));
        var image = (OfdImageElement)mislabeled.Pages[0].Elements[0];
        Assert.True(image.SourceMediaAmbiguous);
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare(image, 4096));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdImageRenderProjection.Prepare(image, 4096, cancelled.Token));
    }

    [Fact]
    public void FreshTypedImageChecksDeclarationAgainstJbig2SignatureWithoutChangingOrdinaryImagePassThrough()
    {
        // SourceXml and Reader ambiguity flags are absent in a caller-created image.
        // The signature still enters the restricted branch and must never emerge as a PNG label.
        var mislabeled = new OfdImageElement { Data = Jbig2, MediaType = "image/png" };
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare(mislabeled, 4096));
        var reversed = new OfdImageElement { Data = Red, MediaType = "image/jbig2" };
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare(reversed, 4096));
        var admitted = new OfdImageElement { Data = Jbig2, MediaType = "image/jbig2" };
        var projected = OfdImageRenderProjection.Prepare(admitted, 4096);
        Assert.Equal("image/png", projected.MediaType);
        using (var pixels = Image.Load<Rgb24>(projected.Data))
        { Assert.Equal(64, pixels.Width); Assert.Equal(56, pixels.Height); }
        var ordinary = new OfdImageElement { Data = Red, MediaType = "image/png" };
        var unchanged = OfdImageRenderProjection.Prepare(ordinary, 1);
        Assert.Same(Red, unchanged.Data);
        Assert.Equal("image/png", unchanged.MediaType);
    }

    [Fact]
    public async Task ImageMaskRejectsSelfReferenceAndDifferentDimensions()
    {
        var cycle = await new OfdReader().ReadAsync(new MemoryStream(Package(Red, "PNG", maskReference: true, maskId: "20")));
        var cyclicImage = (OfdImageElement)cycle.Pages[0].Elements[0];
        Assert.True(cyclicImage.MaskSourceAmbiguous);
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare(cyclicImage, 100));
        var mismatch = await new OfdReader().ReadAsync(new MemoryStream(Package(Jbig2, "JBIG2", White, "PNG")));
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare((OfdImageElement)mismatch.Pages[0].Elements[0], 4096));
        var nonBinary = await new OfdReader().ReadAsync(new MemoryStream(Package(Red, "PNG", NonBinary, "PNG")));
        Assert.Throws<NotSupportedException>(() => OfdImageRenderProjection.Prepare((OfdImageElement)nonBinary.Pages[0].Elements[0], 100));
    }

    [Fact]
    public async Task WhiteMaskPreservesSourceAlphaRatherThanReplacingIt()
    {
        var ofd = Package(HalfAlpha, "PNG", White, "PNG");
        var package = await new OfdReader().ReadAsync(new MemoryStream(ofd));
        var projected = OfdImageRenderProjection.Prepare((OfdImageElement)package.Pages[0].Elements[0], 100);
        using var pixels = Image.Load<Rgba32>(projected.Data);
        Assert.Equal(128, pixels[0, 0].A);
    }

    private static byte[] Package(byte[] image, string format, byte[]? mask = null, string? maskFormat = null,
        bool maskReference = false, string maskId = "21")
    {
        const string ns = "http://www.ofdspec.org/2016";
        var reference = mask is not null || maskReference ? $" ImageMask=\"{maskId}\"" : "";
        var resources = $"<o:MultiMedia ID=\"20\" Type=\"Image\" Format=\"{format}\"><o:MediaFile>image.bin</o:MediaFile></o:MultiMedia>" +
            (mask is null ? "" : $"<o:MultiMedia ID=\"21\" Type=\"Image\" Format=\"{maskFormat}\"><o:MediaFile>mask.bin</o:MediaFile></o:MultiMedia>");
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Encoding.UTF8.GetBytes($"<o:OFD xmlns:o=\"{ns}\" Version=\"1.0\" DocType=\"OFD-H\"><o:DocBody><o:DocInfo/><o:DocRoot>Doc_0/Document.xml</o:DocRoot></o:DocBody></o:OFD>"),
            ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes($"<o:Document xmlns:o=\"{ns}\"><o:CommonData><o:MaxUnitID>100</o:MaxUnitID><o:PageArea><o:PhysicalBox>0 0 100 100</o:PhysicalBox></o:PageArea><o:DocumentRes>DocumentRes.xml</o:DocumentRes></o:CommonData><o:Pages><o:Page ID=\"1\" BaseLoc=\"Pages/Page_0/Content.xml\"/></o:Pages></o:Document>"),
            ["Doc_0/DocumentRes.xml"] = Encoding.UTF8.GetBytes($"<o:Res xmlns:o=\"{ns}\"><o:MultiMedias>{resources}</o:MultiMedias></o:Res>"),
            ["Doc_0/Pages/Page_0/Content.xml"] = Encoding.UTF8.GetBytes($"<o:Page xmlns:o=\"{ns}\"><o:Area><o:PhysicalBox>0 0 100 100</o:PhysicalBox></o:Area><o:Content><o:Layer ID=\"2\"><o:ImageObject ID=\"3\" Boundary=\"10 10 40 35\" ResourceID=\"20\"{reference}/></o:Layer></o:Content></o:Page>"),
            ["Doc_0/image.bin"] = image
        };
        if (mask is not null) entries["Doc_0/mask.bin"] = mask;
        using var output = new MemoryStream();
        using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var (name, bytes) in entries)
            { using var entry = archive.CreateEntry(name, CompressionLevel.NoCompression).Open(); entry.Write(bytes); }
        return output.ToArray();
    }

    private static Image<Rgba32> EmbeddedPng(MemoryStream output)
    {
        var text = Encoding.UTF8.GetString(output.ToArray());
        var match = Regex.Match(text, "data:image/png;base64,([A-Za-z0-9+/=]+)");
        Assert.True(match.Success, "Exporter must embed a PNG, not relabel encoded JBIG2 or drop the mask.");
        return Image.Load<Rgba32>(Convert.FromBase64String(match.Groups[1].Value));
    }
}
