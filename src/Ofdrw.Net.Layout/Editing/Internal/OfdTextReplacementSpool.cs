using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Layout.Editing.Internal;

// Legacy private callsites retain their Stream shape. All allocation, I/O,
// hash/copy, position accounting and cleanup use the one shared concrete engine.
internal sealed class OfdTextReplacementSpool : Stream
{
    private readonly OfdAccountedTemporaryFile actual;
    internal OfdTextReplacementSpool(OfdControlledEditContext context, long capacity, CancellationToken token = default)
    {
        try { context.Material.ChargeCopy(128, token); actual = new OfdAccountedTemporaryFile(context, capacity, token); }
        catch { context.Fail(); throw; }
    }
    internal long ReservedPosition => actual.ReservedPosition;
    internal long ActualReadBytes => actual.ActualReadBytes;
    internal long ActualWrittenBytes => actual.ActualWrittenBytes;
    internal long ReservedReadBytes => actual.ReservedReadBytes;
    internal long ReservedWriteBytes => actual.ReservedWriteBytes;
    internal long OriginalReservedReadBytes => actual.OriginalReservedReadBytes;
    internal long OriginalActualReadBytes => actual.OriginalActualReadBytes;
    internal long SeekCount => actual.SeekCount;
    internal long StageWrittenBytes => actual.StageWrittenBytes;
    internal long StageReservedPosition => actual.StageReservedPosition;
    public override bool CanRead => actual.CanRead;
    public override bool CanWrite => actual.CanWrite;
    public override bool CanSeek => actual.CanSeek;
    public override long Length => actual.Length;
    public override long Position { get => actual.Position; set => actual.Position = value; }
    public override int Read(byte[] buffer, int offset, int count) => actual.Read(buffer, offset, count);
    public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token) => actual.ReadAsync(buffer, offset, count, token);
    public override int ReadByte() => actual.ReadByte();
    public override void Write(byte[] buffer, int offset, int count) => actual.Write(buffer, offset, count);
    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token) => actual.WriteAsync(buffer, offset, count, token);
    public override void WriteByte(byte value) => actual.WriteByte(value);
#if NETSTANDARD2_1
    public override int Read(Span<byte> buffer) => actual.Read(buffer);
    public override ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default) => actual.ReadAsync(buffer, token);
    public override void Write(ReadOnlySpan<byte> buffer) => actual.Write(buffer);
    public override ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default) => actual.WriteAsync(buffer, token);
#endif
    public override long Seek(long offset, SeekOrigin origin) => actual.Seek(offset, origin);
    public override void SetLength(long value) => actual.SetLength(value);
    public override void Flush() => actual.Flush();
    public override Task FlushAsync(CancellationToken token) => actual.FlushAsync(token);
    internal Task<(long Length, byte[] Sha256)> CaptureOriginalAsync(string originalPath, long maximumInputBytes, CancellationToken token)
        => actual.CaptureOriginalAsync(originalPath, maximumInputBytes, token);
    internal Task<(long Length, byte[] Sha256)> ObserveRawZipAsync(CancellationToken token) => actual.ObserveRawZipAsync(token);
    internal Task CopyToStageAsync(FileStream destination, long expectedLength, byte[] expectedSha256, CancellationToken token = default)
        => actual.CopyToStageAsync(destination, expectedLength, expectedSha256, token);
    protected override void Dispose(bool disposing) { if (disposing) actual.Dispose(); base.Dispose(disposing); }
}
