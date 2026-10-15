using System;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Layout.Editing.Internal;

// Only the qualified Layout kernel can name the issuing parameters. CLI cannot
// inject a stage/file/stream/delegate or obtain any owner, path, claim or model.
internal sealed class OfdPrivateAtomicBridge : IDisposable
{
    private readonly OfdOwnedStageScope? scope;
    private readonly OfdOwnedNoopScope? noop;
    private readonly OfdOwnedWrittenSpool? written;
    private readonly OfdOwnedInputSnapshot original;
    private readonly OfdOwnedAtomicStage stage;
    private OfdVerifiedStage? verified;
    private OfdOwnedNoopCopyClaim? noopCopy;
    private OfdOwnedCommitClaim? commit;
    private readonly object sync = new();
    private enum Phase { Created, Copying, Copied, Claiming, Claimed, Committing, Committed, Failed, Disposed }
    private Phase phase;
    private OfdPrivateAtomicBridge(OfdOwnedStageScope scope, OfdOwnedWrittenSpool written,
        OfdOwnedInputSnapshot original, OfdOwnedAtomicStage stage)
    { this.scope = scope; this.written = written; this.original = original; this.stage = stage; }

    private OfdPrivateAtomicBridge(OfdOwnedNoopScope noop, OfdOwnedInputSnapshot original, OfdOwnedAtomicStage stage)
    { this.noop = noop; this.original = original; this.stage = stage; }
    internal static OfdPrivateAtomicBridge CreateNoop(OfdOwnedNoopScope scope, OfdOwnedInputSnapshot actualOriginal,
        string boundOutputPath, CancellationToken token)
    {
        var stage = OfdOwnedPackageSave.CreateStageNoop(scope, actualOriginal, boundOutputPath, token);
        try { return new OfdPrivateAtomicBridge(scope, actualOriginal, stage); }
        catch { try { stage.Dispose(); } catch { } throw; }
    }
    internal static OfdPrivateAtomicBridge Create(OfdOwnedStageScope scope, OfdOwnedWrittenSpool written,
        OfdOwnedInputSnapshot original, string boundOutputPath, CancellationToken token)
    {
        var stage = OfdOwnedPackageSave.CreateAtomicStage(scope, boundOutputPath, token);
        try { return new OfdPrivateAtomicBridge(scope, written, original, stage); }
        catch { try { stage.Dispose(); } catch { } throw; }
    }
    internal Task CopyAsync(CancellationToken token)
    { lock (sync) { Require(Phase.Created); phase = Phase.Copying; } return CopyCoreAsync(token); }
    private async Task CopyCoreAsync(CancellationToken token)
    {
        try
        {
            if (noop is not null)
            {
                var proof = await OfdOwnedPackageSave.CopyActualSnapshotNoopAsync(noop, original, stage, token).ConfigureAwait(false);
                lock (sync) { Require(Phase.Copying); noopCopy = proof; phase = Phase.Copied; }
            }
            else
            {
                var proof = await OfdOwnedPackageSave.CopyVerifiedOutputToStageAsync(scope!, written!, stage, token).ConfigureAwait(false);
                lock (sync) { Require(Phase.Copying); verified = proof; phase = Phase.Copied; }
            }
        }
        catch { lock (sync) phase = Phase.Failed; throw; }
    }
    internal Task ClaimAsync(CancellationToken token)
    { lock (sync) { Require(Phase.Copied); phase = Phase.Claiming; } return ClaimCoreAsync(token); }
    private async Task ClaimCoreAsync(CancellationToken token)
    {
        try
        {
            var proof = noop is not null
                ? await OfdOwnedPackageSave.FinalizeCurrentInputAndSourceAsync(noop, original, noopCopy!, token).ConfigureAwait(false)
                : await OfdOwnedPackageSave.ValidateInputAndClaimCommitAsync(original, verified!, token).ConfigureAwait(false);
            lock (sync) { Require(Phase.Claiming); commit = proof; phase = Phase.Claimed; }
        }
        catch { lock (sync) phase = Phase.Failed; throw; }
    }
    internal Task CommitAsync(CancellationToken token)
    { lock (sync) { Require(Phase.Claimed); phase = Phase.Committing; } return CommitCoreAsync(token); }
    private async Task CommitCoreAsync(CancellationToken token)
    {
        try
        {
            await OfdOwnedPackageSave.CommitActualStageAsync(commit!, stage, token).ConfigureAwait(false);
            lock (sync) phase = Phase.Committed; // no post-rename check can fail a successful save
        }
        catch { lock (sync) phase = Phase.Failed; throw; }
    }
    private void Require(Phase expected)
    { if (phase != expected) throw new InvalidOperationException("PrivateAtomicStageAlreadyConsumed"); }
    public void Dispose()
    {
        lock (sync) { if (phase == Phase.Disposed) return; phase = Phase.Disposed; verified = null; noopCopy = null; commit = null; }
        stage.Dispose();
    }
}
