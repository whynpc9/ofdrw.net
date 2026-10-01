using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class ImageIoBudget
{
    internal static void Positive(double value, string name)
    {
        if (!(value > 0) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(name, "Value must be finite and positive.");
    }

    internal static void Pixels(int width, int height, long maximumPixels, long workingBytes)
    {
        // PDFium uses signed 32-bit lengths for BGRA; bound before native allocation.
        if (width <= 0 || height <= 0 || width > 32768 || height > 32768 ||
            (long)width * height > Math.Min(maximumPixels, int.MaxValue / 4) ||
            (long)width * height > workingBytes / 16)
            throw new InvalidDataException("Image exceeds the configured pixel/dimension/raster working buffer limit.");
    }

    internal static int Dimension(double millimeters, double ppm)
    {
        var pixels = Math.Ceiling(millimeters * ppm);
        if (!(pixels >= 1) || pixels > 32768)
            throw new InvalidDataException("Image dimension must be between 1 and 32768 pixels.");
        return (int)pixels;
    }

    internal static void Page(double width, double height)
    {
        if (!(width > 0) || !(height > 0) || width > 10000 || height > 10000)
            throw new InvalidDataException("Page dimensions must be finite, positive and at most 10000 millimeters.");
    }

    internal static void ImportExtent(double width, double height)
    {
        Page(width, height);
        // The package writer serializes millimeters to three decimal places. Reject sub-micrometer
        // imported extents instead of silently writing a zero-sized page/image or enlarging the input.
        if (width < 0.001d || height < 0.001d)
            throw new InvalidDataException("Imported page/image extents must be at least 0.001 millimeters for OFD serialization.");
    }

    internal static void Streams(Stream input, Stream output)
    {
        if (input is null) throw new ArgumentNullException(nameof(input));
        if (output is null) throw new ArgumentNullException(nameof(output));
        if (!input.CanRead || !output.CanWrite || ReferenceEquals(input, output))
            throw new ArgumentException("Use distinct readable input and writable output streams.");
    }
}

// Disk staging bounds encoded output during encoding, before the caller's output is touched.
// Also supports PdfSharp's seeks and SetLength; seeking cannot bypass the size limit.
internal sealed class ImageIoStagingStream : Stream
{
    private readonly FileStream _file;
    private readonly string _path;
    private readonly long _limit;
    internal ImageIoStagingStream(long limit)
    {
        _limit = limit;
        _path = Path.Combine(Path.GetTempPath(), "ofdrw-image-" + Guid.NewGuid().ToString("N") + ".tmp");
        _file = new FileStream(_path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 81920, FileOptions.Asynchronous);
    }
    internal string PathOnDisk => _path;
    internal async Task CloseWriterAsync(CancellationToken token)
    {
        await _file.FlushAsync(token).ConfigureAwait(false);
        token.ThrowIfCancellationRequested();
        // PDFium reopens by filename. Close the write handle first so Windows sharing rules cannot reject that read.
        _file.Dispose();
    }
    public override bool CanRead => true;
    public override bool CanSeek => true;
    public override bool CanWrite => true;
    public override long Length => _file.Length;
    public override long Position { get => _file.Position; set { Check(value); _file.Position = value; } }
    private void Check(long end) { if (end < 0 || end > _limit) throw new InvalidDataException("Encoded output exceeds the configured byte limit."); }
    public override void Flush() => _file.Flush();
    public override Task FlushAsync(CancellationToken token) => _file.FlushAsync(token);
    public override int Read(byte[] buffer, int offset, int count) => _file.Read(buffer, offset, count);
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) => _file.ReadAsync(buffer, offset, count, token);
    public override long Seek(long offset, SeekOrigin origin)
    {
        var end = checked((origin == SeekOrigin.Begin ? 0 : origin == SeekOrigin.Current ? Position : Length) + offset);
        Check(end); return _file.Seek(offset, origin);
    }
    public override void SetLength(long value) { Check(value); _file.SetLength(value); }
    public override void Write(byte[] buffer, int offset, int count) { Check(checked(Position + count)); _file.Write(buffer, offset, count); }
    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
    { Check(checked(Position + count)); return _file.WriteAsync(buffer, offset, count, token); }
    internal async Task PublishAsync(Stream output, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        Position = 0;
        await CopyToAsync(output, 81920, token).ConfigureAwait(false);
        await output.FlushAsync(token).ConfigureAwait(false);
    }
    protected override void Dispose(bool disposing)
    {
        if (disposing) { _file.Dispose(); File.Delete(_path); }
        base.Dispose(disposing);
    }
}
