using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using PdfSharpCore.Drawing;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class Audit016RegressionTests
{
    [Fact]
    public async Task Audit016_NormalImageControl()
    {
        var package = Package(1); package.Pages.RemoveAt(1);
        var pdf = await Convert(package);
        using var parsed = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
        Assert.Single(Images(parsed));
        AssertColor(pdf, 5, 5, 250, 30, 20);
    }

    [Fact]
    public async Task Audit016_RepeatedImageResources()
    {
        var pdf = await Convert(Package(8));
        using var parsed = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
        Assert.Single(Images(parsed));
        AssertColor(pdf, 5, 5, 250, 30, 20);
        AssertColor(pdf, 19, 19, 250, 30, 20);
    }

    [Fact]
    public async Task Audit016_InterpolationVariants()
    {
        var package = Package(2); package.Pages.RemoveAt(1);
        ((OfdImageElement)package.Pages[0].Elements[0]).SourceXml = Hint(false);
        ((OfdImageElement)package.Pages[0].Elements[1]).SourceXml = Hint(true);
        var pdf = await Convert(package);
        using var parsed = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
        var images = Images(parsed);
        Assert.Equal(2, images.Length);
        Assert.Contains(images, image => image.Elements.GetBoolean("/Interpolate"));
        Assert.Contains(images, image => !image.Elements.GetBoolean("/Interpolate"));
    }

    [Theory]
    [InlineData(false, 255)] [InlineData(true, 255)] [InlineData(false, 128)]
    public void Audit016_NoLiveDecodedPixels(bool jpeg, int alpha)
    {
        var decoded = new List<Image<Rgba32>>();
        var bytes = Bytes();
        if (jpeg) { using var image = Image.Load<Rgba32>(bytes); using var output = new MemoryStream(); image.SaveAsJpeg(output); bytes = output.ToArray(); }
        using var document = new PdfDocument(); var page = document.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
        for (var index = 0; index < 4; index++)
        {
            var element = new OfdImageElement { Data = bytes, Alpha = alpha, WidthMillimeters = 8, HeightMillimeters = 8 };
            OfdToPdfConverter.DrawImage(graphics, new OfdPage(), element, 16, new OfdModelMaterialization(new OfdModelLimits()), default,
                data => { var pixels = Image.Load<Rgba32>(data); decoded.Add(pixels); return pixels; });
            Assert.NotEmpty(decoded);
            foreach (var pixels in decoded) Assert.Throws<ObjectDisposedException>(() => { _ = pixels[0, 0]; });
        }
        // The real PDF image table retains encoded sources only, and realizes a repeated resource once.
        Assert.Equal(alpha == 255 ? 1 : 5, decoded.Count);
        using var pdf = new MemoryStream(); document.Save(pdf, false);
        using var parsed = PdfReader.Open(new MemoryStream(pdf.ToArray()), PdfDocumentOpenMode.Import);
        Assert.Single(Images(parsed));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Audit016_CancelAndDecodeFailure_EncoderReleasesPixels(bool bitmap)
    {
        Image<Rgba32>? decoded = null;
        var source = new EncodedBitmapImageSource(Bytes(), 16, default, data => decoded = Image.Load<Rgba32>(data));
        using var failure = new FailingStream();
        Assert.Throws<IOException>(() => { if (bitmap) source.SaveAsPdfBitmap(failure); else source.SaveAsJpeg(failure); });
        Assert.NotNull(decoded);
        Assert.Throws<ObjectDisposedException>(() => { _ = decoded![0, 0]; });
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void Audit016_CancelAndDecodeFailure_CancelAfterDecodeReleasesPixels(bool bitmap)
    {
        using var cancellation = new CancellationTokenSource();
        Image<Rgba32>? decoded = null;
        var source = new EncodedBitmapImageSource(Bytes(), 16, cancellation.Token, data =>
        { decoded = Image.Load<Rgba32>(data); cancellation.Cancel(); return decoded; });
        using var output = new MemoryStream();
        Assert.ThrowsAny<OperationCanceledException>(() => { if (bitmap) source.SaveAsPdfBitmap(output); else source.SaveAsJpeg(output); });
        Assert.NotNull(decoded);
        Assert.Throws<ObjectDisposedException>(() => { _ = decoded![0, 0]; });
        Assert.Empty(output.ToArray());
    }

    [Fact]
    public void Audit016_CancelAndDecodeFailure_DrawRestoresGraphicsAndReleasesAlphaPixels()
    {
        using var document = new PdfDocument(); var page = document.AddPage();
        using var graphics = XGraphics.FromPdfPage(page);
        var element = new OfdImageElement { Data = Bytes(), Alpha = 128, WidthMillimeters = 8, HeightMillimeters = 8 };
        var calls = 0; Image<Rgba32>? pixels = null;
        Assert.Throws<InvalidDataException>(() => OfdToPdfConverter.DrawImage(graphics, new OfdPage(), element, 16,
            new OfdModelMaterialization(new OfdModelLimits()), default, data =>
            { if (++calls == 2) throw new InvalidDataException("decode failure"); return pixels = Image.Load<Rgba32>(data); }));
        Assert.Equal(2, calls);
        Assert.Throws<ObjectDisposedException>(() => { _ = pixels![0, 0]; });
        graphics.DrawRectangle(XBrushes.Red, 30, 30, 5, 5); // Restored graphics remain usable after failure.
    }

    [Fact]
    public async Task Audit016_AlphaAndOccurrenceTransforms()
    {
        var pdf = await Convert(AlphaPackage());
        using var parsed = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
        Assert.Equal(2, Images(parsed).Length);
        AssertColor(pdf, 4, 4, 255, 0, 0);
        AssertColor(pdf, 16, 4, 255, 255, 0);
        AssertColor(pdf, 28, 4, 255, 127, 127);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit016_AlphaAndNestedAppearance(bool strict)
    {
        var (ofd, pdf) = await Nested(strict);
        using var parsed = PdfReader.Open(new MemoryStream(pdf), PdfDocumentOpenMode.Import);
        Assert.Single(Images(parsed));
        AssertColor(pdf, 6, 11, 250, 30, 20);
        AssertColor(pdf, 24, 11, 250, 30, 20);
    }

    [Fact]
    public async Task Audit016_SaveReviewEvidenceWhenRequested()
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_PLAN016_EVIDENCE");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        foreach (var (name, package) in new[] { ("repeated-images", Package(8)), ("alpha-transforms", AlphaPackage()) })
        {
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd);
            await File.WriteAllBytesAsync(Path.Combine(directory, name + ".ofd"), ofd.ToArray());
            ofd.Position = 0; using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            await File.WriteAllBytesAsync(Path.Combine(directory, name + ".pdf"), pdf.ToArray());
        }
        var nested = await Nested(false);
        await File.WriteAllBytesAsync(Path.Combine(directory, "nested-images.ofd"), nested.Ofd);
        await File.WriteAllBytesAsync(Path.Combine(directory, "nested-images.pdf"), nested.Pdf);
    }

    private static OfdDocumentPackage AlphaPackage()
    {
        var package = Package(0); package.Pages.RemoveAt(1);
        var bytes = ImageIoTests.ImageBytes();
        for (var index = 0; index < 3; index++) package.Pages[0].Elements.Add(new OfdImageElement
        { Data = bytes.ToArray(), FileName = "quadrants.png", XMillimeters = 2 + index * 12, YMillimeters = 2,
          WidthMillimeters = 8, HeightMillimeters = 8, Alpha = index == 2 ? 128 : 255,
          Transform = index == 1 ? new double[] { 0, 8, -8, 0, 8, 0 } : null });
        return package;
    }

    private static async Task<(byte[] Ofd, byte[] Pdf)> Nested(bool strict)
    {
        var package = Package(3); package.Pages.RemoveAt(1);
        using var nested = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, nested);
        var ofd = await ImageIoTests.WithSeals([nested.ToArray()], stampsPerRecord: 2, pageCount: 1);
        using var input = new MemoryStream(ofd); using var pdf = new MemoryStream();
        var model = await new OfdReader().ReadAsync(input);
        await new OfdToPdfConverter().ConvertPackageAsync(model, pdf, null, default, strictAppearanceBudgets: strict);
        return (ofd, pdf.ToArray());
    }

    private sealed class FailingStream : MemoryStream
    {
        public override void Write(byte[] buffer, int offset, int count) => throw new IOException("write failure");
        public override void Write(ReadOnlySpan<byte> buffer) => throw new IOException("write failure");
        public override void WriteByte(byte value) => throw new IOException("write failure");
    }

    private static string Hint(bool value) => $"<ofd:ImageObject xmlns:ofd=\"http://www.ofdspec.org/2016\" xmlns:h=\"https://ofdrw.net/image-hints\" h:PdfInterpolateV1=\"{value.ToString().ToLowerInvariant()}\" />";

    internal static OfdDocumentPackage Package(int count)
    {
        var bytes = Bytes();
        var package = new OfdDocumentPackage();
        for (var pageIndex = 0; pageIndex < 2; pageIndex++)
        {
            var page = new OfdPage { Id = $"{pageIndex + 1}", Index = pageIndex, WidthMillimeters = 40, HeightMillimeters = 40 };
            for (var index = 0; index < count; index++) page.Elements.Add(new OfdImageElement
            {
                Data = bytes.ToArray(), FileName = "image.png", XMillimeters = 2 + index * 2,
                YMillimeters = 2 + index * 2, WidthMillimeters = 8, HeightMillimeters = 8
            });
            package.Pages.Add(page);
        }
        return package;
    }

    internal static byte[] Bytes()
    {
        using var pixels = new Image<Rgba32>(4, 4, new Rgba32(250, 30, 20));
        using var encoded = new MemoryStream(); pixels.SaveAsPng(encoded); return encoded.ToArray();
    }

    private static async Task<byte[]> Convert(OfdDocumentPackage package)
    {
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertPackageAsync(package, pdf, null, default);
        return pdf.ToArray();
    }

    private static PdfDictionary[] Images(PdfDocument document) => document.Internals.GetAllObjects()
        .OfType<PdfDictionary>().Where(item => item.Elements.GetName("/Subtype") == "/Image" && !item.Elements.GetBoolean("/ImageMask") && item.Elements.GetName("/ColorSpace") != "/DeviceGray").ToArray();

    private static void AssertColor(byte[] pdf, double x, double y, byte red, byte green, byte blue)
    {
        using var doc = Docnet.Core.DocLib.Instance.GetDocReader(pdf, new Docnet.Core.Models.PageDimensions(2d));
        using var page = doc.GetPageReader(0); var pixels = page.GetImage();
        var offset = ((int)(y * 72 / 25.4 * 2) * page.GetPageWidth() + (int)(x * 72 / 25.4 * 2)) * 4;
        // PDFium returns BGRA with transparent page background; compare the visible white-page composite.
        int Composite(byte channel) => (channel * pixels[offset + 3] + 255 * (255 - pixels[offset + 3])) / 255;
        Assert.InRange(Composite(pixels[offset + 2]), red - 5, red + 5);
        Assert.InRange(Composite(pixels[offset + 1]), green - 5, green + 5);
        Assert.InRange(Composite(pixels[offset]), blue - 5, blue + 5);
    }
}
