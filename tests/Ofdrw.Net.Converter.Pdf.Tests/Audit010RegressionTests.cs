using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class Audit010RegressionTests
{
    // Only the four-byte magic: no IFD, count, or executable non-progressing payload.
    public static TheoryData<byte[]> BigTiffHeaders => new()
    {
        new byte[] { 0x49, 0x49, 0x2b, 0 },
        new byte[] { 0x4d, 0x4d, 0, 0x2b }
    };

    [Theory]
    [MemberData(nameof(BigTiffHeaders))]
    public async Task Audit010_BigTiffAllEntrypoints_ImportRejectsBeforeDecoder(byte[] header)
    {
        using var input = new MemoryStream(header);
        using var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => new ImageToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("Only PNG and JPEG", exception.Message);
        Assert.Empty(output.ToArray());
    }

    [Theory]
    [MemberData(nameof(BigTiffHeaders))]
    public void Audit010_BigTiffAllEntrypoints_EncodedBitmapRejectsBeforeDecoder(byte[] header)
    {
        var exception = Assert.Throws<InvalidDataException>(() => new EncodedBitmapImageSource(header, 100, default));
        Assert.Contains("BigTIFF", exception.Message);
    }

    [Theory]
    [MemberData(nameof(BigTiffHeaders))]
    public async Task Audit010_BigTiffAllEntrypoints_OfdRejectsBeforeDecoder(byte[] header)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 20, HeightMillimeters = 20 };
        page.Elements.Add(new OfdImageElement { Data = header, FileName = "image.tif", MediaType = "image/tiff",
            WidthMillimeters = 10, HeightMillimeters = 10 });
        package.Pages.Add(page);
        using var input = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, input);
        input.Position = 0;
        using var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdToPdfConverter().ConvertAsync(input, output));
        Assert.Contains("BigTIFF", exception.Message);
        Assert.Empty(output.ToArray());
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Audit010_BigTiffAllEntrypoints_DeferredDecodeRechecksPayload(bool pdfBitmap)
    {
        var bytes = Encoded("png");
        var decoded = false;
        var source = new EncodedBitmapImageSource(bytes, 100, default, _ =>
        {
            decoded = true;
            throw new InvalidOperationException("Decoder must not run for BigTIFF.");
        });
        new byte[] { 0x49, 0x49, 0x2b, 0 }.CopyTo(bytes, 0);
        using var output = new MemoryStream();
        var exception = Assert.Throws<InvalidDataException>(() =>
        {
            if (pdfBitmap) source.SaveAsPdfBitmap(output); else source.SaveAsJpeg(output);
        });
        Assert.Contains("BigTIFF", exception.Message);
        Assert.False(decoded);
        Assert.Empty(output.ToArray());
    }

    [Theory]
    [InlineData("gif")]
    [InlineData("bmp")]
    [InlineData("tiff")]
    public async Task Audit010_ImportRejectsBeforeDecoder_UnsupportedFormatStillWorksInOfd(string format)
    {
        var bytes = Encoded(format);
        using var input = new MemoryStream(bytes);
        using var output = new MemoryStream();
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => new ImageToOfdConverter().ConvertAsync(input, output));
        Assert.Contains("Only PNG and JPEG", exception.Message);
        var source = new EncodedBitmapImageSource(bytes, 100, default);
        Assert.Equal(3, source.Width);
        Assert.Equal(2, source.Height);
        using var bitmap = new MemoryStream();
        source.SaveAsPdfBitmap(bitmap);
        using var pixels = Image.Load<Rgba32>(bitmap.ToArray());
        Assert.Equal(3, pixels.Width);
        Assert.True(pixels[0, 0].R > 200);
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 20, HeightMillimeters = 20 };
        page.Elements.Add(new OfdImageElement { Data = bytes, FileName = "image." + format,
            WidthMillimeters = 10, HeightMillimeters = 10 });
        package.Pages.Add(page);
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        Assert.True(pdf.Length > 0);
    }

    [Theory]
    [InlineData("png", false)]
    [InlineData("jpeg", false)]
    [InlineData("png", true)]
    [InlineData("jpeg", true)]
    public async Task Audit010_ValidImages_StreamPositionAndShortNonSeekableReads(string format, bool nonSeekable)
    {
        var bytes = Encoded(format);
        using var backing = new MemoryStream(new byte[] { 1, 2, 3, 4, 5 }.Concat(bytes).ToArray());
        backing.Position = 5;
        using Stream input = nonSeekable ? new ShortNonSeekableStream(backing) : backing;
        using var output = new MemoryStream();
        await new ImageToOfdConverter().ConvertAsync(input, output);
        Assert.Equal(backing.Length, backing.Position);
        output.Position = 0;
        var image = Assert.IsType<OfdImageElement>(Assert.Single(Assert.Single((await new OfdReader().ReadAsync(output)).Pages).Elements));
        Assert.Equal(bytes, image.Data);
        using var decoded = Image.Load<Rgba32>(image.Data);
        Assert.Equal(3, decoded.Width);
        Assert.Equal(2, decoded.Height);
        if (format == "png") Assert.Equal((byte)128, decoded[0, 0].A);
    }

    [Fact]
    public async Task Audit010_StreamPositionAndCancel_PrecancelDoesNotConsumeOrPublish()
    {
        using var input = new MemoryStream(Encoded("png"));
        using var output = new MemoryStream();
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new ImageToOfdConverter().ConvertAsync(input, output, cancellation.Token));
        Assert.Equal(0, input.Position);
        Assert.Empty(output.ToArray());
        Assert.ThrowsAny<OperationCanceledException>(() => new EncodedBitmapImageSource(new byte[] { 0x49, 0x49, 0x2b, 0 }, 100, cancellation.Token));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Audit010_MalformedAdmittedHeader_CannotPublish(bool png)
    {
        var bytes = png ? new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 } : new byte[] { 255, 216, 255, 224 };
        using var input = new MemoryStream(bytes);
        using var output = new MemoryStream();
        var exception = await Record.ExceptionAsync(() => new ImageToOfdConverter().ConvertAsync(input, output));
        Assert.True(exception is InvalidDataException or InvalidImageContentException or UnknownImageFormatException,
            $"Expected image validation failure, got {exception}");
        Assert.Empty(output.ToArray());
    }

    private static byte[] Encoded(string format)
    {
        using var image = new Image<Rgba32>(3, 2, new Rgba32(250, 30, 20, format == "png" ? (byte)128 : (byte)255));
        using var output = new MemoryStream();
        switch (format)
        {
            case "png": image.SaveAsPng(output); break;
            case "jpeg": image.SaveAsJpeg(output); break;
            case "gif": image.SaveAsGif(output); break;
            case "bmp": image.SaveAsBmp(output); break;
            case "tiff": image.SaveAsTiff(output); break;
            default: throw new ArgumentException(nameof(format));
        }
        return output.ToArray();
    }

    private sealed class ShortNonSeekableStream(Stream source) : Stream
    {
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override int Read(byte[] buffer, int offset, int count) => source.Read(buffer, offset, Math.Min(count, 2));
        public override void Flush() => throw new NotSupportedException();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }
}
