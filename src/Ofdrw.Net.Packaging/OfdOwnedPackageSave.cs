using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal static partial class OfdOwnedPackageSave
{
    internal static Task<OfdOwnedWrittenSpool> WritePrivateTextEditAsync(OfdOwnedWriteScope scope,
        OfdRetainedTextEditBatch batch, CancellationToken token) => OfdPackageWriter.WritePrivateOwnedAsync(scope, batch, token);
    internal static OfdVerifiedReadback VerifyActualReadback(OfdOwnedWrittenSpool written,
        OfdActualReadback actual, CancellationToken token) => written.VerifyActual(actual, token);
    internal static OfdOwnedCompareScope EnterGeometryComparison(OfdVerifiedReadback verified,
        out OfdDocumentPackage actualR, out OfdReadbackRootAddress[] addresses, CancellationToken token)
        => verified.EnterActualComparison(out actualR, out addresses, token);
    internal static OfdOwnedAtomicStage CreateAtomicStage(OfdOwnedStageScope scope,
        string boundOutputPath, CancellationToken token) => OfdOwnedAtomicStage.CreateChanged(scope, boundOutputPath, token);
    internal static Task<OfdVerifiedStage> CopyVerifiedOutputToStageAsync(OfdOwnedStageScope scope,
        OfdOwnedWrittenSpool written, OfdOwnedAtomicStage stage, CancellationToken token) => written.CopyToActualStageAsync(scope, stage, token);
    internal static Task<OfdOwnedCommitClaim> ValidateInputAndClaimCommitAsync(OfdOwnedInputSnapshot original,
        OfdVerifiedStage stage, CancellationToken token) => original.FinalizeChangedAsync(stage, token);
    internal static OfdOwnedAtomicStage CreateStageNoop(OfdOwnedNoopScope scope,
        OfdOwnedInputSnapshot actualOriginal, string boundOutputPath, CancellationToken token)
        => actualOriginal.CreateNoopStage(scope, boundOutputPath, token);
    internal static Task<OfdOwnedNoopCopyClaim> CopyActualSnapshotNoopAsync(OfdOwnedNoopScope scope,
        OfdOwnedInputSnapshot actualOriginal, OfdOwnedAtomicStage stage, CancellationToken token)
        => actualOriginal.CopyNoopAsync(scope, stage, token);
    internal static Task<OfdOwnedCommitClaim> FinalizeCurrentInputAndSourceAsync(OfdOwnedNoopScope scope,
        OfdOwnedInputSnapshot actualOriginal, OfdOwnedNoopCopyClaim exactCopy, CancellationToken token)
        => actualOriginal.FinalizeNoopAsync(scope, exactCopy, token);
    internal static Task CommitActualStageAsync(OfdOwnedCommitClaim claim,
        OfdOwnedAtomicStage exactStage, CancellationToken token) => exactStage.CommitActualAsync(claim, token);
}
