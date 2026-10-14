using System.IO;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing.Internal;

// A concrete private file with precharged allocation, I/O and maximum file
// positions. Logical Writer output reservations remain on the same ledger.
internal sealed class OfdTextReplacementSpool : Stream
{
    private readonly OfdControlledEditContext context;
    private readonly FileStream file;
    private readonly string path;
    private readonly long capacity;
    private readonly object sync = new();
    private readonly CancellationToken operationToken;
    private int busy, disposed;
    private bool failed;
    internal long ReservedPosition { get; private set; }
    internal long ActualReadBytes { get; private set; }
    internal long ActualWrittenBytes { get; private set; }
    internal long ReservedReadBytes { get; private set; }
    internal long ReservedWriteBytes { get; private set; }
    internal long SeekCount { get; private set; }

    internal OfdTextReplacementSpool(OfdControlledEditContext context, long capacity, CancellationToken token = default)
    {
        if (capacity <= 0 || capacity > 512L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(capacity));
        this.context = context; this.capacity = capacity; operationToken = token;
        try
        {
            context.AssertActive(); context.Material.ChargeCopy(16_384, token);
            var directory = Path.GetTempPath();
            context.Material.ChargeCopy(8L * directory.Length, token);
            path = Path.Combine(directory, "ofdrw-replacement-" + Guid.NewGuid().ToString("N") + ".tmp");
            file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous);
        }
        catch { Fail(); throw; }
    }
    public override bool CanRead => disposed == 0;
    public override bool CanSeek => disposed == 0;
    public override bool CanWrite => disposed == 0;
    public override long Length { get { Check(); return file.Length; } }
    public override long Position { get { Check(); return file.Position; } set { Seek(value, SeekOrigin.Begin); } }

    private void Check()
    {
        if (Volatile.Read(ref disposed) != 0) throw new ObjectDisposedException(nameof(OfdTextReplacementSpool));
        context.AssertActive();
    }
    private void Enter()
    {
        lock (sync)
        {
            Check();
            if (busy != 0) throw new InvalidOperationException("Private spool already has an active operation.");
            try { operationToken.ThrowIfCancellationRequested(); } catch { Fail(); throw; }
            busy = 1;
        }
    }
    private void Exit() { lock (sync) busy = 0; }
    private void Fail() { failed = true; context.Fail(); }
    private void CheckAfterIo() { context.AssertActive(); operationToken.ThrowIfCancellationRequested(); }
    private void ReservePosition(long position, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (position < 0 || position > capacity) throw new InvalidDataException("ReplacementSpoolPositionLimit");
        if (position > ReservedPosition) { context.Material.ChargeOutput(position - ReservedPosition); ReservedPosition = position; }
    }
    private void ReserveRead(int count, CancellationToken token)
    { context.Material.ChargeWork(count, token); ReservedReadBytes += count; }
    private void ReserveWrite(int count, CancellationToken token)
    {
        context.Material.ChargeWork(count, token);
        ReservePosition(checked(file.Position + count), token); ReservedWriteBytes += count;
    }
    private static void CheckBuffer(byte[] buffer, int offset, int count)
    {
        if (buffer is null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
    }
    public override int Read(byte[] buffer, int offset, int count)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveRead(count, default); var read = file.Read(buffer, offset, count); ActualReadBytes += read; return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveRead(count, token); var read = await file.ReadAsync(buffer, offset, count, token).ConfigureAwait(false); ActualReadBytes += read; CheckAfterIo(); return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void Write(byte[] buffer, int offset, int count)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveWrite(count, default); file.Write(buffer, offset, count); ActualWrittenBytes += count; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override int ReadByte()
    {
        Enter();
        try { ReserveRead(1, default); var value = file.ReadByte(); if (value >= 0) ActualReadBytes++; return value; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void WriteByte(byte value)
    {
        Enter();
        try { ReserveWrite(1, default); file.WriteByte(value); ActualWrittenBytes++; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveWrite(count, token); await file.WriteAsync(buffer, offset, count, token).ConfigureAwait(false); ActualWrittenBytes += count; CheckAfterIo(); }
        catch { Fail(); throw; } finally { Exit(); }
    }
#if NETSTANDARD2_1
    public override int Read(Span<byte> buffer)
    {
        Enter();
        try { ReserveRead(buffer.Length, default); var read = file.Read(buffer); ActualReadBytes += read; return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
    {
        Enter();
        try { ReserveRead(buffer.Length, token); var read = await file.ReadAsync(buffer, token).ConfigureAwait(false); ActualReadBytes += read; CheckAfterIo(); return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        Enter();
        try { ReserveWrite(buffer.Length, default); file.Write(buffer); ActualWrittenBytes += buffer.Length; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
    {
        Enter();
        try { ReserveWrite(buffer.Length, token); await file.WriteAsync(buffer, token).ConfigureAwait(false); ActualWrittenBytes += buffer.Length; CheckAfterIo(); }
        catch { Fail(); throw; } finally { Exit(); }
    }
#endif
    public override long Seek(long offset, SeekOrigin origin)
    {
        Enter();
        try
        {
            var target = origin switch
            {
                SeekOrigin.Begin => offset,
                SeekOrigin.Current => checked(file.Position + offset),
                SeekOrigin.End => checked(file.Length + offset),
                _ => throw new ArgumentOutOfRangeException(nameof(origin))
            };
            context.Material.ChargeWork(16, default); ReservePosition(target, default);
            var position = file.Seek(target, SeekOrigin.Begin); SeekCount++; return position;
        }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void SetLength(long value)
    {
        Enter();
        try
        {
            context.Material.ChargeWork(Math.Max(16, value - file.Length), default);
            ReservePosition(value, default); file.SetLength(value);
        }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void Flush()
    {
        Enter();
        try { context.Material.ChargeWork(16, default); file.Flush(); }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async Task FlushAsync(CancellationToken token)
    {
        Enter();
        try { context.Material.ChargeWork(16, token); await file.FlushAsync(token).ConfigureAwait(false); CheckAfterIo(); }
        catch { Fail(); throw; } finally { Exit(); }
    }
    protected override void Dispose(bool disposing)
    {
        if (!disposing) return;
        lock (sync)
        {
            if (busy != 0) throw new InvalidOperationException("Cannot dispose a private spool during I/O.");
            if (disposed != 0) return;
            disposed = 1;
        }
        var primaryFailure = failed;
        try { context.AssertActive(); }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { primaryFailure = true; }
        Exception? cleanup = null;
        try { file.Dispose(); } catch (Exception error) { cleanup = error; }
        try { File.Delete(path); } catch (Exception error) { if (cleanup is null) cleanup = error; }
        base.Dispose(disposing);
        if (cleanup is not null && !primaryFailure) ExceptionDispatchInfo.Capture(cleanup).Throw();
    }
}
