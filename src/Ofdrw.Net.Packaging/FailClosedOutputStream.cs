using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Packaging;

// Transparent pass-through to a caller-owned destination. Once poisoned, writes and flushes are
// discarded, so a failed ZIP write never gets its central directory: the partial bytes are not a
// readable archive. Seeks still forward so ZipArchive bookkeeping does not throw. Leaves the
// destination open.
internal sealed class FailClosedOutputStream : Stream
{
    private readonly Stream _inner;
    private bool _poisoned;

    internal FailClosedOutputStream(Stream inner) => _inner = inner ?? throw new ArgumentNullException(nameof(inner));

    internal void Poison() => _poisoned = true;

    // Writes one ZIP through a guard; success disposes normally, failure poisons first.
    internal static async Task WriteZipAsync(Stream destination, Func<ZipArchive, Task> write)
    {
        var guarded = new FailClosedOutputStream(destination);
        var zip = new ZipArchive(guarded, ZipArchiveMode.Create, leaveOpen: true);
        try
        {
            await write(zip).ConfigureAwait(false);
        }
        catch
        {
            guarded.Poison();
            try { zip.Dispose(); } catch { /* the primary exception wins */ }
            throw;
        }
        zip.Dispose();
    }

    public override bool CanRead => _inner.CanRead;
    public override bool CanSeek => _inner.CanSeek;
    public override bool CanWrite => _inner.CanWrite;
    public override long Length => _inner.Length;
    public override long Position { get => _inner.Position; set => _inner.Position = value; }
    public override long Seek(long offset, SeekOrigin origin) => _inner.Seek(offset, origin);
    public override void SetLength(long value) { if (!_poisoned) _inner.SetLength(value); }
    public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);

    public override void Write(byte[] buffer, int offset, int count)
    {
        if (!_poisoned) _inner.Write(buffer, offset, count);
    }

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) =>
        _poisoned ? Task.CompletedTask : _inner.WriteAsync(buffer, offset, count, cancellationToken);

    public override void WriteByte(byte value)
    {
        if (!_poisoned) _inner.WriteByte(value);
    }

    public override void Flush()
    {
        if (!_poisoned) _inner.Flush();
    }

    public override Task FlushAsync(CancellationToken cancellationToken) =>
        _poisoned ? Task.CompletedTask : _inner.FlushAsync(cancellationToken);

    // Never disposes the caller's stream.
    protected override void Dispose(bool disposing) => base.Dispose(disposing);
}
