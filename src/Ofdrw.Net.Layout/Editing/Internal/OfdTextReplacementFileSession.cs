using System.IO;
using System.Security.Cryptography;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Editing.Internal;

// The file entry owns the context before reading any source bytes. There is no
// public overload accepting a context, source model, stream or certificate.
internal sealed class OfdTextReplacementFileSession : IDisposable
{
    private const int BufferSize = 81_920;
    private readonly object sync = new();
    private readonly CancellationToken preparationToken;
    private readonly string inputPath;
    private readonly string snapshotPath;
    private readonly FileStream input;
    private readonly byte[] buffer;
    private readonly byte[] inputDigest;
    private readonly long inputLength;
    private OfdTextReplacementResult? result;
    private OfdTextReplacementOperation? consuming;
    private FileSaveLease? saveLease;
    private int state; // ready, preparing, prepared, consuming, disposed, failed, consumed
    private bool busy;
    private bool validating;
    private bool usingOperation;
    private bool faulted;

    private readonly OfdControlledEditContext context;
    private OfdDocumentPackage? source;
    private bool ownedMode;
    private OfdTextReplacementOperation? ownedOperation;
    private OfdPackageWriter.WrittenTextEdit? ownedWriter;
    private OfdTextReplacementSpool? ownedSpool;
    private OfdDocumentPackage? ownedReadback;
    private OfdSerializedGuard? ownedReadbackGuard;
    private OfdTextReplacementOwnedReport? ownedReport;
    private int ownedSnapshots;
    private bool atomicMetadataReserved, ownedStageVerified, ownedCommitClaimed;
    private bool inputSnapshotReleased;
    private long ownedAtomicPosition;
    // Retain bounded scalar observations after failed owned cleanup so deleted
    // partial bytes remain observable without retaining file/Writer authority.
    private long failedSpoolPosition, failedSpoolRead, failedSpoolWritten, failedSpoolReservedRead, failedSpoolReservedWrite;
    internal OfdControlledEditContext BoundContext { get { lock (sync) { AssertLegacyMode(); return context; } } }
    internal OfdDocumentPackage BoundSource { get { lock (sync) { AssertLegacyMode(); return source ?? throw new ObjectDisposedException(nameof(OfdTextReplacementFileSession)); } } }
    internal OfdTextReplacementOptions Options { get; }
    internal long RawReadBytes { get; private set; }
    internal long RawHashBytes { get; private set; }
    internal long RawReservedReadBytes { get; private set; }
    internal long RawReservedHashBytes { get; private set; }
    internal long InputSnapshotBytes => inputLength;
    internal long NoOpSnapshotCopyBytes { get; private set; }
    internal OfdTextReplacementPhase InitialReadStatistics { get; }

    private OfdTextReplacementFileSession(string path, string snapshotPath, FileStream stream, byte[] buffer, byte[] digest, long length,
        OfdDocumentPackage source, OfdControlledEditContext context, OfdTextReplacementOptions options,
        CancellationToken token, long readBytes, long reservedBytes)
    {
        inputPath = path; this.snapshotPath = snapshotPath; input = stream; this.buffer = buffer; inputDigest = digest; inputLength = length;
        this.source = source; this.context = context; Options = options; preparationToken = token;
        RawReadBytes = RawHashBytes = readBytes; RawReservedReadBytes = RawReservedHashBytes = reservedBytes;
        InitialReadStatistics = Snapshot("InitialRawBindingAndControlledRead", context);
    }

    internal static async Task<OfdTextReplacementFileSession> OpenAsync(string inputPath,
        OfdTextReplacementOptions? options = null, CancellationToken token = default)
    {
        if (inputPath is null) throw new ArgumentNullException(nameof(inputPath));
        var policies = options ?? new OfdTextReplacementOptions();
        var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = policies.Limits.MaxMaterializedBytes });
        FileStream? stream = null;
        string? snapshot = null;
        var created = false;
        var diagnosticsReserved = false;
        var phase = "InitialInputReservation";
        try
        {
            // Session state, path normalization, owned buffer, FileStream buffer
            // and bounded hash state are all reserved before allocation/open.
            context.Material.ChargeCopy(checked(BufferSize + 32_768L + 8L * inputPath.Length + 8L * Path.GetTempPath().Length), token);
            diagnosticsReserved = true;
            var path = Path.GetFullPath(inputPath);
            var buffer = new byte[BufferSize];
            using var sourceInput = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
            var length = sourceInput.Length;
            if (length > new OfdPackageLoadOptions().MaxInputBytes) throw new InvalidDataException("ReplacementInputLimit");
            // One stable owned snapshot is the actual Reader input. Its digest
            // describes exactly the bytes copied, even if the external file
            // changes A->B->A during acquisition. No caller gets this handle.
            context.Material.ChargeOutput(length);
            snapshot = Path.Combine(Path.GetTempPath(), "ofdrw-input-" + Guid.NewGuid().ToString("N") + ".tmp");
            stream = new FileStream(snapshot, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous);
            created = true; phase = "CopyAndHashInputSnapshot";
            var first = await HashAsync(sourceInput, buffer, length, context, token, stream).ConfigureAwait(false);
            context.Material.ChargeWork(16, token); await stream.FlushAsync(token).ConfigureAwait(false);
            SeekStart(stream, context, token);
            phase = "InitialControlledReader";
            var source = await new OfdReader().ReadControlledAsync(stream, new OfdPackageLoadOptions(), context, token).ConfigureAwait(false);
            phase = "VerifyOriginalPathBinding";
            context.Material.ChargeCopy(8192, token);
            using var current = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
            var second = await HashAsync(current, buffer, length, context, token).ConfigureAwait(false);
            if (!first.Digest.SequenceEqual(second.Digest)) throw new InvalidDataException("Replacement input changed while reading.");
            token.ThrowIfCancellationRequested(); context.AssertActive();
            return new OfdTextReplacementFileSession(path, snapshot, stream, buffer, first.Digest, length, source, context, policies,
                token, first.Read + second.Read, first.Reserved + second.Reserved);
        }
        catch (Exception error)
        {
            context.Fail();
            try
            {
                if (diagnosticsReserved)
                { error.Data["Ofdrw.FileReplacement.Phase"] = phase; error.Data["Ofdrw.FileReplacement.Counters"] = Snapshot(phase, context); }
            }
            catch { /* Preserve the original exception, including untrusted Data. */ }
            // A cleanup failure must not replace the actual quota/cancellation/
            // parse error. These are exclusively owned exact temporary paths.
            try { stream?.Dispose(); } catch { }
            try { if (created) File.Delete(snapshot!); } catch { }
            context.Dispose();
            throw;
        }
    }

    internal OfdTextReplacementStatistics Prepare(IOfdGlyphGeometryProvider provider, string keyword, string replacementText,
        IReadOnlyList<int>? pages = null, CancellationToken token = default)
    {
        lock (sync) { AssertLegacyMode(); AssertReady(0, token); AssertIdle(); busy = true; }
        try
        {
            result = new OfdTextReplacer(provider).ReplaceControlled(this, pages, keyword, replacementText, token);
            lock (sync) { context.AssertActive(); state = 2; }
            return result.Statistics;
        }
        catch { Fail(); throw; }
        finally { lock (sync) busy = false; }
    }

    // Sole production caller: CLI's private job, which constructs the concrete
    // trusted provider and freezes ordinary values itself. No public result or
    // model/Writer authority is returned by this private composition path.
    internal Task<OfdTextReplacementOwnedReport> PrepareAndWriteOwnedAsync(IOfdGlyphGeometryProvider provider, string keyword,
        string replacementText, int[]? pageIndexes = null, CancellationToken token = default)
    {
        lock (sync)
        {
            AssertReady(0, token); AssertIdle();
            if (ownedMode) throw new InvalidOperationException("Private preparation is already claimed.");
            ownedMode = true; busy = true;
        }
        try
        {
            if (pageIndexes is not null && pageIndexes.Length > 1000) throw new InvalidDataException("ReplacementPageLimit");
            // Prepay the four private async continuations, Tasks and owner
            // reference shells before invoking any async method. Diagnostic
            // reports retain their separate original operation reservation.
            context.Material.ChargeCopy(4096L + 256L + 8L * (pageIndexes?.Length ?? 0), token);
            var pages = pageIndexes?.ToArray();
            return PrepareAndWriteOwnedCoreAsync(provider, keyword, replacementText, pages, token);
        }
        catch { Fail(); lock (sync) busy = false; throw; }
    }

    private async Task<OfdTextReplacementOwnedReport> PrepareAndWriteOwnedCoreAsync(IOfdGlyphGeometryProvider provider, string keyword,
        string replacementText, int[]? pages, CancellationToken token)
    {
        try
        {
            await new OfdTextReplacer(provider).PrepareAndWriteKeywordForOwnedAsync(this, source!, context, pages, keyword, replacementText, token).ConfigureAwait(false);
            lock (sync)
            {
                context.AssertActive(); preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested();
                state = ownedWriter is null ? 2 : 3;
                return ownedReport!;
            }
        }
        catch
        {
            Fail(); CleanupOwnedWriteFailure(); throw;
        }
        finally { lock (sync) busy = false; }
    }

    internal void ReserveOwnedConfiguration(int identityCount, int pageCount, CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(0, token); AssertIdle();
            try
            {
                if (identityCount <= 0 || identityCount > 64 || pageCount < 0 || pageCount > 1000)
                    throw new ArgumentOutOfRangeException(nameof(identityCount));
                context.Material.ChargeCopy(8192L + 512L * identityCount + 16L * pageCount, token);
            }
            catch { Fail(); throw; }
        }
    }

    internal void ClaimOwnedPreparation(OfdDocumentPackage exactSource, OfdControlledEditContext exactContext, CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(0, token);
            if (!ownedMode || !busy || token != preparationToken || !ReferenceEquals(exactSource, source) || !ReferenceEquals(exactContext, context))
            { Fail(); throw new InvalidOperationException("Private preparation has a foreign operation binding."); }
            state = 1;
        }
    }
    internal void FailOwnedPreparation() => Fail();
    private void AssertOwnedOperation(OfdTextReplacementOperation operation)
    {
        context.AssertActive(); preparationToken.ThrowIfCancellationRequested();
        if (!ownedMode || !busy || state != 1 || !ReferenceEquals(operation.Source, source) || !ReferenceEquals(operation.Context, context) ||
            (ownedOperation is not null && !ReferenceEquals(ownedOperation, operation)))
            throw new InvalidOperationException("Private preparation cannot publish a foreign owner.");
    }

    internal void CreateOwnedOutputSpool(OfdTextReplacementOperation operation, CancellationToken token)
    {
        lock (sync)
        {
            AssertOwnedOperation(operation); token.ThrowIfCancellationRequested();
            if (ownedSpool is not null) throw new InvalidOperationException("Private output spool was already created.");
            ownedOperation = operation;
            operation.Begin("CreateOwnedOutputSpool");
            ownedSpool = new OfdTextReplacementSpool(context, Options.Limits.MaxMaterializedBytes, preparationToken);
        }
    }

    internal void AcceptOwnedNoOp(OfdTextReplacementOperation operation, OfdTextReplacementObservation[] observations)
    {
        lock (sync)
        {
            AssertOwnedOperation(operation);
            if (ownedSpool is not null || ownedWriter is not null || ownedReport is not null)
                throw new InvalidOperationException("No-op cannot accept changed output ownership.");
            var report = new OfdTextReplacementOwnedReport(observations, operation.Snapshot());
            ownedOperation = operation; ownedReport = report;
        }
    }

    internal async Task WriteContinuousOwnedAsync(OfdTextReplacementOperation operation, OfdRetainedTextEditBatch batch,
        OfdTextReplacement[] requests, CancellationToken token)
    {
        lock (sync)
        {
            AssertOwnedOperation(operation);
            if (ownedSpool is null || ownedWriter is not null || ownedReport is not null)
                throw new InvalidOperationException("Private changed output has no unconsumed spool.");
        }
        OfdPackageWriter.WrittenTextEdit? written = null;
        try
        {
            // Sole real shared entry: its synchronous prefix runs before this
            // first await, with no PreparedWrite escape/caller-return boundary.
            written = await OfdPackageWriter.WriteContinuousTextEditAsync(operation.Source!, operation.SourceGuard!, context, batch, ownedSpool, token).ConfigureAwait(false);
            preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested(); context.AssertActive();
            operation.Mark("WrittenOwnedZip");
            await ownedSpool.FlushAsync(token).ConfigureAwait(false);
            preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested(); context.AssertActive();
            operation.Mark("WroteOwnedOutput");
            context.Material.ChargeCopy(512L * requests.Length + 256, token);
            var observations = requests.Select(value => new OfdTextReplacementObservation(value)).ToArray();
            var report = new OfdTextReplacementOwnedReport(observations, operation.Snapshot());
            lock (sync)
            {
                AssertOwnedOperation(operation); token.ThrowIfCancellationRequested();
                ownedWriter = written; written = null; ownedReport = report;
            }
        }
        catch (Exception error)
        {
            Fail(); operation.Fail(error);
            try { written?.Dispose(); } catch { }
            CleanupOwnedWriteFailure(); throw;
        }
    }

    private void CleanupOwnedWriteFailure()
    {
        // Called only after terminal transition; secondary owned disposal errors
        // never mask the actual write/quota/cancellation failure.
        try { ownedWriter?.Dispose(); } catch { }
        ownedWriter = null;
        if (ownedSpool is not null)
        {
            failedSpoolPosition = ownedSpool.ReservedPosition;
            failedSpoolRead = ownedSpool.ActualReadBytes; failedSpoolWritten = ownedSpool.ActualWrittenBytes;
            failedSpoolReservedRead = ownedSpool.ReservedReadBytes; failedSpoolReservedWrite = ownedSpool.ReservedWriteBytes;
            try { ownedSpool.Dispose(); } catch { }
            ownedSpool = null;
        }
    }

    internal OfdTextReplacementPhase SnapshotOwned()
    {
        lock (sync)
        {
            if (!ownedMode || ++ownedSnapshots > 16) throw new InvalidOperationException("Private snapshot capacity exceeded.");
            // Fixed bounded reporting is covered by initial reservation, even
            // when a terminal quota failure prevents further ledger charges.
            return Snapshot(ownedOperation?.Phase ?? "OwnedPreparation", context);
        }
    }
    internal OfdTextReplacementOwnedFiles SnapshotOwnedFiles()
    {
        lock (sync)
        {
            if (!ownedMode || ++ownedSnapshots > 16) throw new InvalidOperationException("Private snapshot capacity exceeded.");
            return new OfdTextReplacementOwnedFiles(inputLength, ownedSpool?.ReservedPosition ?? failedSpoolPosition, ownedAtomicPosition,
                ownedSpool?.ActualReadBytes ?? failedSpoolRead, ownedSpool?.ActualWrittenBytes ?? failedSpoolWritten,
                ownedSpool?.ReservedReadBytes ?? failedSpoolReservedRead, ownedSpool?.ReservedWriteBytes ?? failedSpoolReservedWrite);
        }
    }

    internal async Task ReadOwnedOutputAsync(CancellationToken token = default)
    {
        lock (sync)
        {
            AssertReady(3, token); AssertIdle();
            if (!ownedMode || ownedSpool is null || ownedReadback is not null)
                throw new InvalidOperationException("Owned output is unavailable or was already read back.");
            busy = true;
        }
        try
        {
            ownedOperation!.Begin("ReadbackSourceCandidateGuard");
            ownedWriter!.AssertSourceAndCandidateUnchanged(token);
            ownedOperation!.Begin("ActualOwnedReadback"); ownedSpool.Position = 0;
            ownedReadback = await new OfdReader().ReadControlledAsync(ownedSpool, new OfdPackageLoadOptions(), context, token).ConfigureAwait(false);
            preparationToken.ThrowIfCancellationRequested(); context.AssertActive();
            ownedOperation.Mark("ReadOwnedOutput");
            ownedOperation.Begin("CaptureOwnedReadbackGuard"); ownedReadbackGuard = context.CaptureSource(ownedReadback, token);
            ownedWriter.AssertSourceAndCandidateUnchanged(token); ownedOperation.Mark("CapturedOwnedReadbackGuard");
            // Runtime full geometry/SCR verification is a distinct mandatory
            // stage. Reading bytes alone never qualifies AtomicOutput commit.
        }
        catch (Exception error) { Fail(); ownedOperation!.Fail(error); CleanupOwnedWriteFailure(); throw; }
        finally { lock (sync) busy = false; }
    }

    internal void ReserveOwnedAtomicMetadata(string destination, CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(2, token); AssertIdle();
            if (!ownedMode || ownedReport is null || ownedReport.Changed || atomicMetadataReserved)
                throw new InvalidOperationException("Atomic stage admission requires an unconsumed no-op job.");
            try
            {
                context.Material.ChargeCopy(81_920 + 32_768L + 8L * (inputPath.Length + (long)destination.Length), token);
                // Directory/file creation, Commit's flush/close/rename and the
                // known AtomicOutput cleanup are metadata work, not payload I/O.
                context.Material.ChargeWork(256, token);
                atomicMetadataReserved = true;
            }
            catch { Fail(); throw; }
        }
    }

    internal void ReserveOwnedPathLookup(long existingCharacters, CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(2, token); AssertIdle();
            if (!ownedMode || !atomicMetadataReserved) throw new InvalidOperationException("Output path admission is unavailable.");
            try
            {
                // Native path/link metadata is bounded by the platform path
                // APIs; reserve a 32K-character target before each lookup.
                context.Material.ChargeCopy(1024 + 4L * (existingCharacters + 32_768), token);
                context.Material.ChargeWork(16 + 2L * existingCharacters, token);
            }
            catch { Fail(); throw; }
        }
    }

    internal void ReserveOwnedPathDecomposition(int characters, CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(2, token); AssertIdle();
            if (!ownedMode || !atomicMetadataReserved) throw new InvalidOperationException("Output path admission is unavailable.");
            try
            {
                // FullPath, substring, components array and one string shell
                // per smallest possible component, including recursive targets.
                context.Material.ChargeCopy(4096L + 32L * (characters + 1L), token);
                context.Material.ChargeWork(2L * characters, token);
            }
            catch { Fail(); throw; }
        }
    }

    internal async Task CopyOwnedNoOpToStageAsync(FileStream destination, CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(2, token); AssertIdle();
            if (!ownedMode || ownedReport is null || ownedReport.Changed || !atomicMetadataReserved || ownedOperation is null)
                throw new InvalidOperationException("No-op stage copy is not available.");
            state = 3; busy = true;
        }
        try
        {
            ownedOperation.Begin("CopyOwnedNoOpStage");
            if (destination.GetType() != typeof(FileStream) || !destination.CanWrite || !destination.CanSeek || destination.Position != 0 || destination.Length != 0)
                throw new InvalidOperationException("No-op copy requires the new concrete AtomicOutput file.");
            ownedOperation.Validate(token); context.Material.ChargeOutput(inputLength); ownedAtomicPosition = inputLength; SeekStart(input, context, token);
            var copy = await HashAsync(input, buffer, inputLength, context, token, destination, preparationToken).ConfigureAwait(false);
            if (!inputDigest.SequenceEqual(copy.Digest) || destination.Length != inputLength) throw new InvalidDataException("Private input snapshot changed.");
            NoOpSnapshotCopyBytes += copy.Read;
            context.Material.ChargeWork(16, token); await destination.FlushAsync(token).ConfigureAwait(false);
            preparationToken.ThrowIfCancellationRequested(); context.AssertActive();
            ownedOperation.Begin("FinalNoOpOriginalBinding"); await ValidateInputAsync(token).ConfigureAwait(false);
            ownedOperation.Validate(token); ownedOperation.Mark("VerifiedOwnedNoOpStage"); ownedStageVerified = true;
        }
        catch (Exception error) { Fail(); ownedOperation.Fail(error); throw; }
        finally { lock (sync) busy = false; }
    }

    internal async Task ClaimOwnedCommitAsync(CancellationToken token)
    {
        lock (sync)
        {
            AssertReady(3, token); AssertIdle();
            if (!ownedMode || !ownedStageVerified || ownedCommitClaimed) throw new InvalidOperationException("Owned stage is not ready for publication.");
            busy = true;
        }
        try
        {
            ownedOperation!.Begin("FinalizeOwnedNoOpCommit");
            // Cleanup that can still fail precedes the final rename. Keep the
            // context/source guard alive through final raw validation/commit.
            context.Material.ChargeWork(32, token); input.Dispose(); File.Delete(snapshotPath); inputSnapshotReleased = true;
            await ValidateInputAsync(token).ConfigureAwait(false);
            ownedOperation.Validate(token); preparationToken.ThrowIfCancellationRequested(); context.AssertActive();
            ownedCommitClaimed = true;
        }
        catch (Exception error) { Fail(); ownedOperation!.Fail(error); throw; }
        finally { lock (sync) busy = false; }
    }

    internal void ClaimPreparation(OfdDocumentPackage source, OfdControlledEditContext context, CancellationToken token)
    {
        lock (sync)
        {
            AssertLegacyMode(); AssertReady(0, token);
            if (!ReferenceEquals(source, this.source) || !ReferenceEquals(context, this.context) || token != preparationToken)
            { Fail(); throw new InvalidOperationException("File replacement preparation belongs to another source, context or token."); }
            if (!busy) { Fail(); throw new InvalidOperationException("File preparation must be owned by its session."); }
            state = 1;
        }
    }

    internal FileSaveLease BeginSave(CancellationToken token = default)
    {
        lock (sync)
        {
            AssertLegacyMode(); AssertReady(2, token);
            AssertIdle();
            try
            {
                context.Material.ChargeCopy(256, token);
                var lease = new FileSaveLease(this);
                consuming = result!.Claim(); saveLease = lease; state = 3; busy = true;
                return lease;
            }
            catch { Fail(); throw; }
        }
    }

    // Reopen the pathname as well as retaining the original handle: a renamed
    // replacement at that path must not silently pass an old-handle hash.
    internal async Task AssertInputUnchangedAsync(CancellationToken token = default)
    {
        lock (sync)
        {
            AssertLegacyMode(); AssertIdle();
            if (state is not (0 or 2)) throw new InvalidOperationException("File replacement session cannot validate input in this state.");
            busy = true;
        }
        try { await ValidateInputAsync(token).ConfigureAwait(false); }
        finally { lock (sync) busy = false; }
    }

    private async Task ValidateInputAsync(CancellationToken token)
    {
        try
        {
            preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested(); context.AssertActive();
            context.Material.ChargeCopy(8192, token);
            using var current = new FileStream(inputPath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.Asynchronous);
            var actual = await HashAsync(current, buffer, inputLength, context, token, operationToken: preparationToken).ConfigureAwait(false);
            RawReadBytes += actual.Read; RawHashBytes += actual.Read;
            RawReservedReadBytes += actual.Reserved; RawReservedHashBytes += actual.Reserved;
            preparationToken.ThrowIfCancellationRequested();
            if (!inputDigest.SequenceEqual(actual.Digest)) throw new InvalidDataException("Replacement input changed during the operation.");
        }
        catch { Fail(); throw; }
    }

    internal sealed class FileSaveLease : IDisposable
    {
        private readonly OfdTextReplacementFileSession owner;
        private int disposed;
        internal FileSaveLease(OfdTextReplacementFileSession owner) { this.owner = owner; }
        internal OperationUse UseOperation()
        {
            lock (owner.sync)
            {
                AssertOwned();
                if (owner.validating || owner.usingOperation) throw new InvalidOperationException("Save operation I/O is already active.");
                try
                {
                    owner.context.Material.ChargeCopy(128, owner.preparationToken);
                    var use = new OperationUse(this, owner.consuming!); owner.usingOperation = true; return use;
                }
                catch { owner.Fail(); throw; }
            }
        }
        internal sealed class OperationUse : IDisposable
        {
            private readonly FileSaveLease lease;
            private OfdTextReplacementOperation? operation;
            internal OperationUse(FileSaveLease lease, OfdTextReplacementOperation operation) { this.lease = lease; this.operation = operation; }
            internal OfdTextReplacementOperation Operation
            { get { lock (lease.owner.sync) { lease.AssertOwned(); return operation ?? throw new ObjectDisposedException(nameof(OperationUse)); } } }
            public void Dispose()
            {
                lock (lease.owner.sync)
                {
                    if (operation is null) return;
                    lease.owner.ObserveContextFailure(); operation = null; lease.owner.usingOperation = false;
                }
            }
        }
        internal CancellationToken CancellationToken => owner.preparationToken;
        private void AssertOwned()
        {
            if (disposed != 0 || !ReferenceEquals(owner.saveLease, this) || owner.state != 3)
                throw new InvalidOperationException("File save lease is no longer active.");
        }
        internal async Task AssertInputUnchangedAsync(CancellationToken token = default)
        {
            lock (owner.sync)
            {
                AssertOwned();
                if (owner.validating || owner.usingOperation) throw new InvalidOperationException("Save operation I/O is already active.");
                owner.validating = true;
            }
            try { await owner.ValidateInputAsync(token).ConfigureAwait(false); }
            finally { lock (owner.sync) owner.validating = false; }
        }
        // Only an owned concrete file destination is accepted. No caller Stream
        // receives the private buffer; the source is always the bound snapshot.
        internal async Task CopyNoOpSnapshotAsync(FileStream destination, CancellationToken token = default)
        {
            lock (owner.sync)
            {
                AssertOwned();
                if (owner.validating || owner.usingOperation) throw new InvalidOperationException("Save operation I/O is already active.");
                owner.validating = true;
            }
            try
            {
                owner.preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested();
                if (destination.GetType() != typeof(FileStream) || !destination.CanWrite || !destination.CanSeek || destination.Position != 0 || destination.Length != 0)
                    throw new InvalidOperationException("Snapshot copy requires a new owned concrete output file.");
                if (owner.result!.Changed || owner.consuming!.Candidate is not null)
                    throw new InvalidOperationException("Changed replacement cannot consume the no-op snapshot path.");
                owner.context.Material.ChargeOutput(owner.inputLength);
                SeekStart(owner.input, owner.context, token);
                var copy = await HashAsync(owner.input, owner.buffer, owner.inputLength, owner.context, token, destination, owner.preparationToken).ConfigureAwait(false);
                if (!owner.inputDigest.SequenceEqual(copy.Digest)) throw new InvalidDataException("Private input snapshot changed.");
                owner.preparationToken.ThrowIfCancellationRequested(); owner.NoOpSnapshotCopyBytes += copy.Read;
            }
            catch { owner.Fail(); throw; }
            finally { lock (owner.sync) owner.validating = false; }
        }
        public void Dispose()
        {
            lock (owner.sync)
            {
                if (disposed != 0) return;
                if (owner.validating || owner.usingOperation) throw new InvalidOperationException("Cannot release a save lease during active I/O.");
                if (!ReferenceEquals(owner.saveLease, this)) throw new InvalidOperationException("Foreign file save lease.");
                disposed = 1;
                owner.ObserveContextFailure();
                try { owner.consuming!.Dispose(); owner.result!.CompleteConsumption(); }
                finally { owner.consuming = null; owner.saveLease = null; owner.busy = false; owner.state = 6; }
            }
        }
    }

    private void AssertReady(int expected, CancellationToken token)
    {
        lock (sync)
        {
            if (state != expected) throw new InvalidOperationException("File replacement session is not ready for this operation.");
            try { preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested(); context.AssertActive(); }
            catch { Fail(); throw; }
        }
    }
    private void Fail() { lock (sync) { faulted = true; if (state != 4) state = 5; context.Fail(); } }
    private void ObserveContextFailure()
    {
        try { context.AssertActive(); }
        catch (ObjectDisposedException) { }
        catch (InvalidOperationException) { faulted = true; }
    }
    private void AssertLegacyMode()
    { if (ownedMode) throw new InvalidOperationException("The private owned job does not expose the legacy model/operation surfaces."); }
    private void AssertIdle() { if (busy) throw new InvalidOperationException("File replacement session already has an active operation."); }

    private static void SeekStart(FileStream stream, OfdControlledEditContext context, CancellationToken token)
    { context.Material.ChargeWork(16, token); stream.Position = 0; }

    private static async Task<(byte[] Digest, long Read, long Reserved)> HashAsync(FileStream stream, byte[] buffer, long length,
        OfdControlledEditContext context, CancellationToken token, FileStream? ownedSnapshot = null, CancellationToken operationToken = default)
    {
        void Check() { context.AssertActive(); token.ThrowIfCancellationRequested(); operationToken.ThrowIfCancellationRequested(); }
        Check();
        if (stream.Length != length) throw new InvalidDataException("Replacement input length changed.");
        context.Material.ChargeCopy(4096 + 32, token);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        long total = 0, reserved = 0;
        while (total < length)
        {
            Check();
            var count = (int)Math.Min(buffer.Length, length - total);
            context.Material.ChargeWork(checked((ownedSnapshot is null ? 2L : 3L) * count), token); reserved += count;
            var read = await stream.ReadAsync(buffer, 0, count, token).ConfigureAwait(false);
            Check();
            if (read == 0) throw new EndOfStreamException("Replacement input was truncated.");
            hash.AppendData(buffer, 0, read);
            if (ownedSnapshot is not null) await ownedSnapshot.WriteAsync(buffer, 0, read, token).ConfigureAwait(false);
            Check(); total += read;
        }
        Check();
        if (stream.Length != length) throw new InvalidDataException("Replacement input length changed.");
        return (hash.GetHashAndReset(), total, reserved);
    }

    internal static OfdTextReplacementPhase Snapshot(string name, OfdControlledEditContext context)
    {
        var m = context.Material;
        return new(name, m.StorageBytes, m.CloneCopyBytes, m.OutputUpperBoundBytes, m.ValidationWorkBytes,
            m.Structure.References, m.Structure.Nodes);
    }

    public void Dispose()
    {
        lock (sync)
        {
            if (state == 4) return;
            AssertIdle();
            if (state == 1) throw new InvalidOperationException("Cannot dispose a file session during preparation.");
            try
            {
                if (consuming is not null) { consuming.Dispose(); consuming = null; result!.CompleteConsumption(); }
                result?.Dispose(); result = null;
                ObserveContextFailure();
                ownedWriter?.Dispose(); ownedWriter = null;
                ownedSpool?.Dispose(); ownedSpool = null;
                ownedReadback = null; ownedReadbackGuard = null;
                ownedOperation?.Dispose(); ownedOperation = null;
                source = null;
            }
            finally
            {
                state = 4;
                Exception? cleanup = null;
                if (!inputSnapshotReleased)
                {
                    try { input.Dispose(); } catch (Exception error) { cleanup = error; }
                    try { File.Delete(snapshotPath); } catch (Exception error) { if (cleanup is null) cleanup = error; }
                }
                context.Dispose();
                if (cleanup is not null && !faulted) ExceptionDispatchInfo.Capture(cleanup).Throw();
            }
        }
    }
}
