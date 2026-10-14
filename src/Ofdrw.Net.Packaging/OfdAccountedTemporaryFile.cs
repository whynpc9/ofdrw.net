using System.IO;
using System.Runtime.ExceptionServices;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

// Shared concrete temporary-file storage, extracted from the private Layout
// spool. Layout's legacy adapter and the owned save capabilities use this one
// implementation; construction alone never grants owned-model authority.
// A concrete private file with precharged allocation, I/O and maximum file
// positions. Logical Writer output reservations remain on the same ledger.
internal sealed class OfdAccountedTemporaryFile : Stream
{
    private readonly OfdControlledEditContext context;
    private readonly OfdOwnedIoObservation? observation;
    private readonly FileStream file;
    private readonly string path;
    private readonly long capacity;
    private readonly object sync = new();
    private readonly CancellationToken operationToken;
    private int busy, disposed;
    private bool failed;
    private bool copiedToStage;
    internal long ReservedPosition { get; private set; }
    internal long ActualReadBytes { get; private set; }
    internal long ActualWrittenBytes { get; private set; }
    internal long ReservedReadBytes { get; private set; }
    internal long OriginalReservedReadBytes { get; private set; }
    internal long OriginalActualReadBytes { get; private set; }
    internal long ReservedWriteBytes { get; private set; }
    internal long SeekCount { get; private set; }
    internal long StageWrittenBytes { get; private set; }
    internal long StageReservedPosition { get; private set; }

    internal OfdAccountedTemporaryFile(OfdControlledEditContext context, long capacity, CancellationToken token = default, OfdOwnedIoObservation? observation = null)
    {
        if (capacity <= 0 || capacity > 512L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(capacity));
        this.context = context; this.capacity = capacity; operationToken = token; this.observation = observation;
        try
        {
            context.AssertActive(); context.Material.ChargeCopy(16_384, token);
            var directory = Path.GetTempPath();
            context.Material.ChargeCopy(8L * directory.Length, token);
            path = Path.Combine(directory, "ofdrw-replacement-" + Guid.NewGuid().ToString("N") + ".tmp");
            if (observation is not null) context.Material.ChargeWork(16, token);
            file = new FileStream(path, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous);
            observation?.RecordOpened();
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
        if (Volatile.Read(ref disposed) != 0) throw new ObjectDisposedException(nameof(OfdAccountedTemporaryFile));
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
    private void Fail() { failed = true; observation?.RecordIoFailure(); context.Fail(); }
    private void CheckAfterIo() { context.AssertActive(); operationToken.ThrowIfCancellationRequested(); }
    private void ReservePosition(long position, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (position < 0 || position > capacity) throw new InvalidDataException("ReplacementSpoolPositionLimit");
        if (position > ReservedPosition) { context.Material.ChargeOutput(position - ReservedPosition); ReservedPosition = position; }
        observation?.ReservePosition(position);
    }
    private void ReserveRead(int count, CancellationToken token)
    { context.Material.ChargeWork(count, token); if (observation is not null) context.Material.ChargeWork(16, token); ReservedReadBytes += count; observation?.ReserveRead(count); }
    private void ReserveWrite(int count, CancellationToken token)
    {
        context.Material.ChargeWork(count, token);
        ReservePosition(checked(file.Position + count), token); ReservedWriteBytes += count;
        if (observation is not null) context.Material.ChargeWork(16, token);
        observation?.ReserveWrite(count);
    }
    private static void CheckBuffer(byte[] buffer, int offset, int count)
    {
        if (buffer is null) throw new ArgumentNullException(nameof(buffer));
        if (offset < 0 || count < 0 || offset > buffer.Length - count) throw new ArgumentOutOfRangeException(nameof(count));
    }
    public override int Read(byte[] buffer, int offset, int count)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveRead(count, default); var read = file.Read(buffer, offset, count); ActualReadBytes += read; observation?.CompleteRead(read); return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken token)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveRead(count, token); var read = await file.ReadAsync(buffer, offset, count, token).ConfigureAwait(false); ActualReadBytes += read; observation?.CompleteRead(read); CheckAfterIo(); return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void Write(byte[] buffer, int offset, int count)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveWrite(count, default); observation?.BeginActualWrite(); file.Write(buffer, offset, count); ActualWrittenBytes += count; observation?.CompleteWrite(count); }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override int ReadByte()
    {
        Enter();
        try { ReserveRead(1, default); var value = file.ReadByte(); if (value >= 0) ActualReadBytes++; observation?.CompleteRead(value >= 0 ? 1 : 0); return value; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void WriteByte(byte value)
    {
        Enter();
        try { ReserveWrite(1, default); observation?.BeginActualWrite(); file.WriteByte(value); ActualWrittenBytes++; observation?.CompleteWrite(1); }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken token)
    {
        CheckBuffer(buffer, offset, count); Enter();
        try { ReserveWrite(count, token); observation?.BeginActualWrite(); await file.WriteAsync(buffer, offset, count, token).ConfigureAwait(false); ActualWrittenBytes += count; observation?.CompleteWrite(count); CheckAfterIo(); }
        catch { Fail(); throw; } finally { Exit(); }
    }
#if NETSTANDARD2_1
    public override int Read(Span<byte> buffer)
    {
        Enter();
        try { ReserveRead(buffer.Length, default); var read = file.Read(buffer); ActualReadBytes += read; observation?.CompleteRead(read); return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async ValueTask<int> ReadAsync(Memory<byte> buffer, CancellationToken token = default)
    {
        Enter();
        try { ReserveRead(buffer.Length, token); var read = await file.ReadAsync(buffer, token).ConfigureAwait(false); ActualReadBytes += read; observation?.CompleteRead(read); CheckAfterIo(); return read; }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void Write(ReadOnlySpan<byte> buffer)
    {
        Enter();
        try { ReserveWrite(buffer.Length, default); observation?.BeginActualWrite(); file.Write(buffer); ActualWrittenBytes += buffer.Length; observation?.CompleteWrite(buffer.Length); }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async ValueTask WriteAsync(ReadOnlyMemory<byte> buffer, CancellationToken token = default)
    {
        Enter();
        try { ReserveWrite(buffer.Length, token); observation?.BeginActualWrite(); await file.WriteAsync(buffer, token).ConfigureAwait(false); ActualWrittenBytes += buffer.Length; observation?.CompleteWrite(buffer.Length); CheckAfterIo(); }
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
            observation?.ReserveSeek(); var position = file.Seek(target, SeekOrigin.Begin); SeekCount++; observation?.CompleteSeek(position); return position;
        }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void SetLength(long value)
    {
        Enter();
        try
        {
            context.Material.ChargeWork(Math.Max(16, value - file.Length), default);
            ReservePosition(value, default); file.SetLength(value); observation?.CompleteLength(value);
        }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override void Flush()
    {
        Enter();
        try { context.Material.ChargeWork(16, default); observation?.ReserveFlush(); file.Flush(); observation?.CompleteFlush(); }
        catch { Fail(); throw; } finally { Exit(); }
    }
    public override async Task FlushAsync(CancellationToken token)
    {
        Enter();
        try { context.Material.ChargeWork(16, token); observation?.ReserveFlush(); await file.FlushAsync(token).ConfigureAwait(false); observation?.CompleteFlush(); CheckAfterIo(); }
        catch { Fail(); throw; } finally { Exit(); }
    }

    // Private initial snapshot primitive. Source read, ordered SHA and snapshot
    // write are charged together before each read; no Stream override doubles it.
    internal async Task<(long Length, byte[] Sha256)> CaptureOriginalAsync(string originalPath,
        long maximumInputBytes, CancellationToken token)
    {
        Enter();
        FileStream? source = null; IncrementalHash? hash = null; ExceptionDispatchInfo? primary = null;
        (long Length, byte[] Sha256) result = default;
        try
        {
          try
          {
            if (file.Length != 0 || file.Position != 0) throw new InvalidOperationException("Snapshot is not fresh.");
            context.Material.ChargeCopy(81_920 + 16_384 + 64, token);
            source = new FileStream(originalPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
            var length = source.Length;
            if (length < 22 || length > maximumInputBytes) throw new InvalidDataException("Original input length is outside its limit.");
            ReservePosition(length, token);
            var buffer = new byte[81_920];
            hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            long copied = 0;
            while (copied < length)
            {
                CheckAfterIo(); token.ThrowIfCancellationRequested();
                var count = (int)Math.Min(buffer.Length, length - copied);
                context.Material.ChargeWork(3L * count + (observation is null ? 0 : 16), token); ReservedWriteBytes += count; OriginalReservedReadBytes += count;
                observation?.ReserveOriginalRead(count); observation?.ReserveWrite(count);
                var read = await source.ReadAsync(buffer, 0, count, token).ConfigureAwait(false);
                OriginalActualReadBytes += read; observation?.CompleteOriginalRead(read);
                CheckAfterIo(); token.ThrowIfCancellationRequested();
                if (read == 0) throw new EndOfStreamException("Original input was truncated.");
                hash.AppendData(buffer, 0, read);
                observation?.BeginActualWrite(); await file.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
                ActualWrittenBytes += read; observation?.CompleteWrite(read); copied += read;
            }
            context.Material.ChargeWork(17, token);
            observation?.ReserveOriginalRead(1);
            if (source.Length != length || source.ReadByte() != -1 || file.Length != length)
                throw new InvalidDataException("Original input changed during snapshot capture.");
            context.Material.ChargeWork(32 + 16, token);
            var digest = hash.GetHashAndReset(); observation?.ReserveFlush();
            await file.FlushAsync(token).ConfigureAwait(false); observation?.CompleteFlush();
            CheckAfterIo(); token.ThrowIfCancellationRequested();
            context.Material.ChargeWork(16, token); observation?.ReserveSeek(); file.Position = 0; SeekCount++; observation?.CompleteSeek(0);
            result = (length, digest);
          }
          catch (Exception error) { Fail(); primary = ExceptionDispatchInfo.Capture(error); }
          DisposeHashOnce(ref hash, context, ref primary);
          DisposeFileOnce(ref source, context, ref primary);
          primary?.Throw(); return result;
        }
        finally { Exit(); }
    }

    // A complete ordered current-byte pass. Hashing seeked Writer writes cannot
    // substitute for this final physical ZIP observation.
    internal async Task<(long Length, byte[] Sha256)> ObserveRawZipAsync(CancellationToken token)
    {
        Enter();
        IncrementalHash? hash = null; ExceptionDispatchInfo? primary = null;
        (long Length, byte[] Sha256) result = default;
        try
        {
          try
          {
            context.Material.ChargeCopy(81_920 + 8192 + 32, token);
            var length = file.Length; var buffer = new byte[81_920];
            hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            context.Material.ChargeWork(16, token); observation?.ReserveSeek(); file.Position = 0; SeekCount++; observation?.CompleteSeek(0);
            long total = 0;
            while (total < length)
            {
                CheckAfterIo(); token.ThrowIfCancellationRequested();
                var count = (int)Math.Min(buffer.Length, length - total);
                context.Material.ChargeWork(2L * count + (observation is null ? 0 : 16), token); ReservedReadBytes += count; observation?.ReserveRead(count);
                var read = await file.ReadAsync(buffer, 0, count, token).ConfigureAwait(false); ActualReadBytes += read; observation?.CompleteRead(read);
                CheckAfterIo(); token.ThrowIfCancellationRequested();
                if (read == 0) throw new EndOfStreamException("Private ZIP was truncated during observation.");
                hash.AppendData(buffer, 0, read); total += read;
            }
            context.Material.ChargeWork(32 + 16, token);
            if (file.Length != length) throw new InvalidDataException("Private ZIP length changed.");
            var digest = hash.GetHashAndReset();
            observation?.ReserveSeek(); file.Position = 0; SeekCount++; observation?.CompleteSeek(0); result = (length, digest);
          }
          catch (Exception error) { Fail(); primary = ExceptionDispatchInfo.Capture(error); }
          DisposeHashOnce(ref hash, context, ref primary);
          primary?.Throw(); return result;
        }
        finally { Exit(); }
    }

    internal static async Task ValidateOriginalPathAsync(OfdControlledEditContext context,
        string path, long expectedLength, byte[] expectedSha256, CancellationToken token)
    {
        FileStream? source = null; IncrementalHash? hash = null; ExceptionDispatchInfo? primary = null;
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            context.Material.ChargeCopy(81_920 + 16_384 + 32, token);
            source = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
            if (source.Length != expectedLength) throw new InvalidDataException("Original input length changed before commit.");
            var buffer = new byte[81_920];
            hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            long total = 0;
            while (total < expectedLength)
            {
                context.AssertActive(); token.ThrowIfCancellationRequested();
                var count = (int)Math.Min(buffer.Length, expectedLength - total);
                context.Material.ChargeWork(2L * count, token);
                var read = await source.ReadAsync(buffer, 0, count, token).ConfigureAwait(false);
                context.AssertActive(); token.ThrowIfCancellationRequested();
                if (read == 0) throw new EndOfStreamException("Original input was truncated before commit.");
                hash.AppendData(buffer, 0, read); total += read;
            }
            context.Material.ChargeWork(32 + 17, token);
            if (source.Length != expectedLength || source.ReadByte() != -1 || !expectedSha256.SequenceEqual(hash.GetHashAndReset()))
                throw new InvalidDataException("Original input bytes changed before commit.");
            context.AssertActive(); token.ThrowIfCancellationRequested();
        }
        catch (Exception error) { context.Fail(); primary = ExceptionDispatchInfo.Capture(error); }
        DisposeHashOnce(ref hash, context, ref primary);
        DisposeFileOnce(ref source, context, ref primary);
        primary?.Throw();
    }

    private static void DisposeHashOnce(ref IncrementalHash? handle, OfdControlledEditContext context,
        ref ExceptionDispatchInfo? primary)
    {
        var actual = handle; handle = null;
        if (actual is null) return;
        try { actual.Dispose(); }
        catch (Exception error) { context.Fail(); if (primary is null) primary = ExceptionDispatchInfo.Capture(error); }
    }
    private static void DisposeFileOnce(ref FileStream? handle, OfdControlledEditContext context,
        ref ExceptionDispatchInfo? primary)
    {
        var actual = handle; handle = null;
        if (actual is null) return;
        try { actual.Dispose(); }
        catch (Exception error) { context.Fail(); if (primary is null) primary = ExceptionDispatchInfo.Capture(error); }
    }

    // Byte-copy primitive only. Its caller must first verify the actual Reader
    // result and later run final source/candidate/readback guards before commit.
    // This expected digest grants no geometry or Writer receipt authority.
    internal async Task CopyToStageAsync(FileStream destination, long expectedLength, byte[] expectedSha256,
        CancellationToken token = default, OfdOwnedIoObservation? stageObservation = null)
    {
        Enter();
        try
        {
            if (copiedToStage) throw new InvalidOperationException("Private spool stage copy has already been consumed.");
            copiedToStage = true;
            token.ThrowIfCancellationRequested();
            context.Material.ChargeWork(16, token);
            if (destination is null || destination.GetType() != typeof(FileStream) || !destination.CanWrite || !destination.CanSeek ||
                destination.Position != 0 || destination.Length != 0 || expectedLength < 0 || expectedLength > capacity ||
                file.Length != expectedLength || expectedSha256 is null || expectedSha256.Length != 32)
                throw new InvalidDataException("ReplacementStageBindingMismatch");
            context.Material.ChargeCopy(81_920 + 8192 + 32, token);
            var baseline = expectedSha256.ToArray(); var buffer = new byte[81_920];
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            context.Material.ChargeOutput(expectedLength); StageReservedPosition = expectedLength; stageObservation?.ReservePosition(expectedLength);
            context.Material.ChargeWork(16, token); observation?.ReserveSeek(); file.Position = 0; SeekCount++; observation?.CompleteSeek(0);
            long copied = 0;
            while (copied < expectedLength)
            {
                CheckAfterIo(); token.ThrowIfCancellationRequested();
                var count = (int)Math.Min(buffer.Length, expectedLength - copied);
                // Read, SHA and destination write all admitted before the read.
                context.Material.ChargeWork(3L * count + (observation is null ? 0 : 16), token); ReservedReadBytes += count;
                observation?.ReserveRead(count); stageObservation?.ReserveWrite(count);
                var read = await file.ReadAsync(buffer, 0, count, token).ConfigureAwait(false); ActualReadBytes += read; observation?.CompleteRead(read);
                CheckAfterIo(); token.ThrowIfCancellationRequested();
                if (read == 0) throw new EndOfStreamException("Private spool was truncated during stage copy.");
                hash.AppendData(buffer, 0, read);
                stageObservation?.BeginActualWrite();
                await destination.WriteAsync(buffer, 0, read, token).ConfigureAwait(false); StageWrittenBytes += read; stageObservation?.CompleteWrite(read);
                CheckAfterIo(); token.ThrowIfCancellationRequested(); copied += read;
            }
            context.Material.ChargeWork(32 + 16, token);
            if (file.Length != expectedLength || destination.Length != expectedLength || !baseline.SequenceEqual(hash.GetHashAndReset()))
                throw new InvalidDataException("ReplacementStageBindingMismatch");
            context.Material.ChargeWork(16, token); stageObservation?.ReserveFlush(); await destination.FlushAsync(token).ConfigureAwait(false); stageObservation?.CompleteFlush();
            CheckAfterIo(); token.ThrowIfCancellationRequested();
        }
        catch { stageObservation?.RecordIoFailure(); Fail(); throw; } finally { Exit(); }
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
        try { File.Delete(path); if (cleanup is null) observation?.RecordRemoved(); } catch (Exception error) { if (cleanup is null) cleanup = error; }
        base.Dispose(disposing);
        if (cleanup is not null && !primaryFailure) ExceptionDispatchInfo.Capture(cleanup).Throw();
    }
}
