using System;
using System.IO;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

// Both branches use this exact CreateNew/copy/close/rename owner. A caller
// FileStream or arbitrary commit delegate cannot create or replace its file.
internal sealed class OfdOwnedAtomicStage : IDisposable
{
    private readonly OfdOwnedModelsLease owner;
    private readonly OfdControlledEditContext context;
    private readonly string destination, path;
    private readonly FileStream file;
    private readonly OfdOwnedIoObservation observation;
    private readonly object sync = new();
    private int busy, copied, closed, committed, disposed;
    private OfdOwnedAtomicStage(OfdOwnedNoopScope scope, OfdOwnedSnapshotCapture capture,
        string destination, CancellationToken token)
    {
        owner = scope.Owner; this.destination = destination;
        context = owner.BeginNoopStage(scope, capture, this, destination, token);
        context.Material.ChargeCopy(16_384L + 16L * destination.Length, token);
        var directory = Path.GetDirectoryName(destination);
        if (string.IsNullOrEmpty(directory)) throw new InvalidDataException("Atomic destination requires a canonical parent.");
        path = Path.Combine(directory, ".ofdrw-save-" + Guid.NewGuid().ToString("N") + ".tmp");
        observation = owner.BindActualStageObservation(this, token);
        context.Material.ChargeWork(16, token);
        file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous);
        observation.RecordOpened();
    }
    private OfdOwnedAtomicStage(OfdOwnedStageScope scope, string destination, CancellationToken token)
    {
        owner = scope.Owner; this.destination = destination;
        context = owner.BeginStage(scope, this, destination, token);
        context.Material.ChargeCopy(16_384L + 16L * destination.Length, token);
        var directory = Path.GetDirectoryName(destination);
        if (string.IsNullOrEmpty(directory)) throw new InvalidDataException("Atomic destination requires a canonical parent.");
        path = Path.Combine(directory, ".ofdrw-save-" + Guid.NewGuid().ToString("N") + ".tmp");
        observation = owner.BindActualStageObservation(this, token);
        context.Material.ChargeWork(16, token);
        file = new FileStream(path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096, FileOptions.Asynchronous);
        observation.RecordOpened();
    }
    internal static OfdOwnedAtomicStage CreateChanged(OfdOwnedStageScope scope, string destination, CancellationToken token)
    {
        try { return new(scope, destination, token); } catch { scope.Owner.Fail(); throw; }
    }
    internal async Task CopyActualOutputAsync(OfdOwnedStageScope scope, OfdOwnedWrittenSpool written,
        OfdAccountedTemporaryFile output, long length, byte[] digest, CancellationToken token)
    {
        Enter(token);
        try
        {
            if (Interlocked.CompareExchange(ref copied, 1, 0) != 0) throw new InvalidOperationException("Atomic stage has already been copied.");
            owner.BeginStageCopy(scope, written, this, token);
            await output.CopyToStageAsync(file, length, digest, token, observation).ConfigureAwait(false);
            context.Material.ChargeWork(16, token); file.Dispose(); Volatile.Write(ref closed, 1);
            context.AssertActive(); token.ThrowIfCancellationRequested(); Volatile.Write(ref copied, 2);
        }
        catch { owner.Fail(); throw; }
        finally { Volatile.Write(ref busy, 0); }
    }
    internal static OfdOwnedAtomicStage CreateNoop(OfdOwnedNoopScope scope, OfdOwnedSnapshotCapture capture,
        string destination, CancellationToken token)
    {
        try { return new(scope, capture, destination, token); }
        catch { scope.Owner.Fail(); throw; }
    }
    private void Enter(CancellationToken token)
    {
        lock (sync)
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (Volatile.Read(ref disposed) != 0 || Interlocked.CompareExchange(ref busy, 1, 0) != 0)
                throw new InvalidOperationException("Atomic stage is disposed or already active.");
        }
    }
    internal async Task CopyActualNoopAsync(OfdOwnedNoopScope scope, OfdOwnedSnapshotCapture capture,
        OfdAccountedTemporaryFile original, long length, byte[] digest, CancellationToken token)
    {
        Enter(token);
        try
        {
            if (Interlocked.CompareExchange(ref copied, 1, 0) != 0) throw new InvalidOperationException("Atomic stage has already been copied.");
            owner.BeginNoopCopy(scope, capture, this, token);
            context.Material.ChargeCopy(512, token);
            capture.RequireBytes(BitConverter.ToString(digest).Replace("-", string.Empty).ToLowerInvariant(), length, owner);
            await original.CopyToStageAsync(file, length, digest, token, observation).ConfigureAwait(false);
            context.Material.ChargeWork(16, token); file.Dispose(); Volatile.Write(ref closed, 1);
            context.AssertActive(); token.ThrowIfCancellationRequested(); Volatile.Write(ref copied, 2);
        }
        catch { owner.Fail(); throw; }
        finally { Volatile.Write(ref busy, 0); }
    }
    internal OfdOwnedCommitClaim ClaimCommit(CancellationToken token)
    {
        Enter(token);
        try
        {
            if (Volatile.Read(ref copied) != 2 || Volatile.Read(ref closed) != 1)
                throw new InvalidOperationException("Atomic commit requires the actual copied and closed stage.");
            return owner.ClaimCommit(this, token);
        }
        catch { owner.Fail(); throw; }
        finally { Volatile.Write(ref busy, 0); }
    }
    internal Task CommitActualAsync(OfdOwnedCommitClaim claim, CancellationToken token)
    {
      lock (sync)
      {
        Enter(token);
        try
        {
            if (Volatile.Read(ref copied) != 2 || Volatile.Read(ref closed) != 1 || Volatile.Read(ref committed) != 0)
                throw new InvalidOperationException("Atomic stage is not commit-ready.");
            context.Material.ChargeWork(32, token);
            owner.RequireCurrentDestination(token);
            using var critical = owner.EnterCommitCriticalSection(claim, this, token);
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (File.Exists(destination)) File.Replace(path, destination, null);
            else File.Move(path, destination);
            Volatile.Write(ref committed, 1); observation.RecordRemoved();
            // No cleanup or cancellation check can reinterpret successful rename.
            critical.RecordActualSuccessfulRename();
            return Task.CompletedTask;
        }
        catch { owner.Fail(); throw; }
        finally { Volatile.Write(ref busy, 0); }
      }
    }
    public void Dispose()
    {
        lock (sync)
        {
            if (Volatile.Read(ref busy) != 0) { owner.Fail(); throw new InvalidOperationException("Cannot dispose an active atomic stage."); }
            if (Interlocked.Exchange(ref disposed, 1) != 0) return;
        }
        if (Volatile.Read(ref committed) != 0) return;
        Exception? error = null;
        try { file.Dispose(); } catch (Exception failure) { owner.Fail(); error = failure; }
        try { File.Delete(path); if (error is null) observation.RecordRemoved(); } catch (Exception failure) { owner.Fail(); if (error is null) error = failure; }
        if (error is not null) ExceptionDispatchInfo.Capture(error).Throw();
    }
}
