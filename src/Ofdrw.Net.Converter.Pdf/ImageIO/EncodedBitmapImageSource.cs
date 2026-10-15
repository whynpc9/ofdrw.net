using System;
using System.IO;
using System.Threading;
using MigraDocCore.DocumentObjectModel.MigraDoc.DocumentObjectModel.Shapes;
using Ofdrw.Net.Core.IO;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.Formats.Bmp;
using SixLabors.ImageSharp.Formats.Jpeg;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// PdfSharpCore retains IImageSource through document.ImageTable. Keep only encoded bytes and dimensions there.
// Its PdfImage constructor realizes the stream synchronously; each encoder call owns and releases decoded pixels.
internal sealed class EncodedBitmapImageSource : ImageSource.IImageSource
{
    private readonly byte[] _encoded;
    private readonly Func<byte[], Image<Rgba32>> _decode;
    private readonly CancellationToken _token;
    internal EncodedBitmapImageSource(byte[] encoded, long maximumPixels, CancellationToken token,
        Func<byte[], Image<Rgba32>>? decode = null)
    {
        _encoded = encoded;
        _token = token;
        token.ThrowIfCancellationRequested();
        var info = Image.Identify(encoded, out var format) ?? throw new InvalidDataException("Cannot identify signature bitmap.");
        ImageIoBudget.Pixels(info.Width, info.Height, maximumPixels, long.MaxValue);
        Width = info.Width; Height = info.Height;
        Transparent = format.Name == "PNG"; // Match the locked PdfSharpCore ImageSharp source's encoding choice.
        Name = "*ofdrw-seal-" + BinaryIdentity.Hash(encoded);
        _decode = decode ?? (bytes => Image.Load<Rgba32>(bytes));
    }
    public int Width { get; }
    public int Height { get; }
    public string Name { get; }
    public bool Transparent { get; }
    public void SaveAsJpeg(MemoryStream stream)
    {
        _token.ThrowIfCancellationRequested();
        using var image = _decode(_encoded);
        image.Save(stream, new JpegEncoder { Quality = 75 });
        _token.ThrowIfCancellationRequested();
    }
    public void SaveAsPdfBitmap(MemoryStream stream)
    {
        _token.ThrowIfCancellationRequested();
        using var image = _decode(_encoded);
        image.Save(stream, new BmpEncoder { BitsPerPixel = BmpBitsPerPixel.Pixel32 });
        _token.ThrowIfCancellationRequested();
    }
}
