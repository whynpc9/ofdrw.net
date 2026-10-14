using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

// Actual original-byte owner. No stream, path or hash is exposed as a getter.
internal sealed class OfdOwnedInputSnapshot : IDisposable
{
    private readonly OfdOwnedReadClaim readClaim;
    private readonly OfdOwnedSnapshotCapture capture;
    private readonly OfdControlledEditContext context;
    private readonly string originalPath;
    private readonly OfdAccountedTemporaryFile file;
    private byte[] digest = Array.Empty<byte>();
    private long length;
    private readonly object sync = new();
    private int readerStarted, disposed, busy;
    private OfdOwnedInputSnapshot(OfdOwnedReadClaim claim, OfdOwnedSnapshotCapture capture,
        OfdControlledEditContext context, string originalPath, long capacity, CancellationToken token)
    {
        readClaim = claim; this.capture = capture; this.context = context; this.originalPath = originalPath;
        file = new OfdAccountedTemporaryFile(context, capacity, token, claim.Owner.BindActualInputObservation(claim, capture, token));
    }
    internal static async Task<OfdOwnedInputSnapshot> CaptureActualAsync(OfdOwnedReadClaim claim,
        OfdControlledEditContext context, string originalPath, long capacity, long maximumInputBytes, CancellationToken token)
    {
        OfdOwnedInputSnapshot? result = null;
        try
        {
            claim.Owner.RequireInitialAcquisition(claim, context, originalPath, token);
            context.AssertActive();
            var capture = OfdOwnedSaveAuthority.BeginActualSnapshotCapture(claim);
            context.Material.ChargeCopy(1024, token);
            result = new OfdOwnedInputSnapshot(claim, capture, context, originalPath, capacity, token);
            var actual = await result.file.CaptureOriginalAsync(originalPath, maximumInputBytes, token).ConfigureAwait(false);
            result.length = actual.Length; result.digest = actual.Sha256;
            context.Material.ChargeCopy(512, token);
            var hex = BitConverter.ToString(actual.Sha256).Replace("-", string.Empty).ToLowerInvariant();
            OfdOwnedSaveAuthority.CompleteActualSnapshotCapture(capture, hex, actual.Length, token);
            return result;
        }
        catch { claim.Owner.Fail(); result?.Dispose(); throw; }
    }
    // One-use actual Reader handoff. The stream local stays inside that bridge;
    // this cannot be used as a repeatable capability/property on the snapshot.
    internal void EnterActualInitialReader(OfdOwnedReadClaim exact, out Stream actualInput,
        out OfdOwnedSnapshotCapture actualCapture, CancellationToken token)
    {
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            lock (sync)
            {
                if (!ReferenceEquals(exact, readClaim) || disposed != 0 || busy != 0 ||
                    Interlocked.CompareExchange(ref readerStarted, 1, 0) != 0)
                    throw new InvalidOperationException("Foreign or repeated initial Reader handoff.");
                busy = 1;
            }
            capture.RequireCompleted(readClaim.Owner);
            file.Position = 0; actualInput = file; actualCapture = capture;
        }
        catch { readClaim.Owner.Fail(); throw; }
    }
    internal void CompleteActualInitialReader(OfdOwnedReadClaim exact)
    {
        lock (sync)
        {
            if (!ReferenceEquals(readClaim, exact) || busy != 1)
            { readClaim.Owner.Fail(); throw new InvalidOperationException("Foreign initial Reader completion."); }
            busy = 0;
        }
    }
    internal void AbortActualInitialReader()
    {
        readClaim.Owner.Fail(); lock (sync) busy = 0;
    }
    private void EnterOwnedOperation(CancellationToken token)
    {
        lock (sync)
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (disposed != 0 || busy != 0) throw new InvalidOperationException("Original snapshot is disposed or in use.");
            busy = 2;
        }
    }
    private void CloseAfterOwnedOperation()
    {
        lock (sync) { if (busy != 2 || disposed != 0) throw new InvalidOperationException("Invalid original cleanup phase."); disposed = 1; busy = 0; }
        try { file.Dispose(); } catch { readClaim.Owner.Fail(); throw; }
    }
    internal OfdOwnedAtomicStage CreateNoopStage(OfdOwnedNoopScope exact, string destination, CancellationToken token)
    {
        try
        {
            if (!ReferenceEquals(exact.Owner, readClaim.Owner) || Volatile.Read(ref disposed) != 0)
                throw new InvalidOperationException("No-change stage has a foreign or disposed original snapshot.");
            return OfdOwnedAtomicStage.CreateNoop(exact, capture, destination, token);
        }
        catch { readClaim.Owner.Fail(); throw; }
    }
    internal async Task<OfdOwnedNoopCopyClaim> CopyNoopAsync(OfdOwnedNoopScope exact,
        OfdOwnedAtomicStage stage, CancellationToken token)
    {
        var entered = false;
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (!ReferenceEquals(exact.Owner, readClaim.Owner) || Volatile.Read(ref disposed) != 0)
                throw new InvalidOperationException("No-change copy has a foreign or disposed snapshot.");
            EnterOwnedOperation(token); entered = true;
            await stage.CopyActualNoopAsync(exact, capture, file, length, digest, token).ConfigureAwait(false);
            // Potentially throwing close/delete happens before current input/S
            // and rename, with private immutable original-byte evidence retained.
            CloseAfterOwnedOperation();
            context.Material.ChargeCopy(512, token);
            return OfdOwnedNoopCopyClaim.Issue(exact, this, stage);
        }
        catch
        {
            readClaim.Owner.Fail();
            if (entered) { lock (sync) busy = 0; try { Dispose(); } catch { } }
            throw;
        }
    }
    internal async Task<OfdOwnedCommitClaim> FinalizeNoopAsync(OfdOwnedNoopScope exact,
        OfdOwnedNoopCopyClaim copy, CancellationToken token)
    {
        try
        {
            var stage = copy.ConsumeActual(exact, this);
            readClaim.Owner.RequireInputIntent(originalPath, token);
            await OfdAccountedTemporaryFile.ValidateOriginalPathAsync(context, originalPath, length, digest, token).ConfigureAwait(false);
            readClaim.Owner.AssertFinalNoopSource(exact, capture, stage, token);
            return stage.ClaimCommit(token);
        }
        catch { readClaim.Owner.Fail(); throw; }
    }
    internal async Task<OfdOwnedCommitClaim> FinalizeChangedAsync(OfdVerifiedStage actualCopy, CancellationToken token)
    {
        try
        {
            var stage = actualCopy.ConsumeForFinalInput(readClaim.Owner);
            EnterOwnedOperation(token); CloseAfterOwnedOperation(); // before final checks; immutable input evidence remains private
            readClaim.Owner.RequireInputIntent(originalPath, token);
            await OfdAccountedTemporaryFile.ValidateOriginalPathAsync(context, originalPath, length, digest, token).ConfigureAwait(false);
            return stage.ClaimCommit(token);
        }
        catch { readClaim.Owner.Fail(); throw; }
    }
    public void Dispose()
    {
        lock (sync)
        {
            if (busy != 0) { readClaim.Owner.Fail(); throw new InvalidOperationException("Cannot dispose an original snapshot during its Reader/copy operation."); }
            if (disposed != 0) return;
            disposed = 1;
        }
        readClaim.Owner.Fail(); // abandonment cannot retain private save eligibility
        file.Dispose();
    }
}

internal sealed class OfdOwnedNoopCopyClaim
{
    private readonly OfdOwnedNoopScope scope;
    private readonly OfdOwnedInputSnapshot snapshot;
    private readonly OfdOwnedAtomicStage stage;
    private int consumed;
    private OfdOwnedNoopCopyClaim(OfdOwnedNoopScope scope, OfdOwnedInputSnapshot snapshot, OfdOwnedAtomicStage stage)
    { this.scope = scope; this.snapshot = snapshot; this.stage = stage; }
    internal static OfdOwnedNoopCopyClaim Issue(OfdOwnedNoopScope scope, OfdOwnedInputSnapshot snapshot, OfdOwnedAtomicStage stage)
        => new(scope, snapshot, stage);
    internal OfdOwnedAtomicStage ConsumeActual(OfdOwnedNoopScope exact, OfdOwnedInputSnapshot actual)
    {
        if (!ReferenceEquals(scope, exact) || !ReferenceEquals(snapshot, actual) || Interlocked.CompareExchange(ref consumed, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or replayed no-change copy proof.");
        return stage;
    }
}
