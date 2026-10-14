using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging;

internal sealed class OfdOwnedWrittenSpool : IDisposable
{
    private readonly OfdOwnedWriteScope scope;
    private readonly OfdOwnedEditCandidate candidate;
    private readonly OfdAccountedTemporaryFile file;
    private readonly OfdPackageWriter.OwnedWriteReceipt receipt;
    private readonly long length;
    private readonly byte[] digest;
    private OfdOwnedZipCensus? actualCensus;
    private int disposed, actualReadCompleted;
    private OfdOwnedWrittenSpool(OfdOwnedWriteScope scope, OfdOwnedEditCandidate candidate, OfdAccountedTemporaryFile file,
        OfdPackageWriter.OwnedWriteReceipt receipt, long length, byte[] digest)
    { this.scope = scope; this.candidate = candidate; this.file = file; this.receipt = receipt; this.length = length; this.digest = digest; }
    internal static OfdOwnedWrittenSpool FromActualWrite(OfdOwnedWriteScope scope, OfdOwnedEditCandidate candidate,
        OfdAccountedTemporaryFile file, OfdPackageWriter.OwnedWriteReceipt receipt, long length, byte[] digest, CancellationToken token)
    {
        candidate.Context.Material.ChargeCopy(1024, token);
        return new(scope, candidate, file, receipt, length, digest);
    }
    internal void EnterActualOutputReader(out Stream actualInput, out OfdControlledEditContext context,
        out OfdOwnedZipCensus census, CancellationToken token)
    {
        try
        {
            if (Volatile.Read(ref disposed) != 0) throw new ObjectDisposedException(nameof(OfdOwnedWrittenSpool));
            context = scope.Owner.BeginOutputRead(this, token);
            census = OfdOwnedZipCensus.Begin(context, token); actualCensus = census;
            file.Position = 0; actualInput = file;
        }
        catch { scope.Owner.Fail(); throw; }
    }
    internal OfdActualReadback CompleteActualOutputReader(OfdDocumentPackage actualR,
        OfdOwnedZipCensus census, CancellationToken token)
    {
        try
        {
            if (!ReferenceEquals(census, actualCensus) || Interlocked.CompareExchange(ref actualReadCompleted, 1, 0) != 0)
                throw new InvalidOperationException("Foreign or repeated actual output Reader completion.");
            scope.Owner.CompleteOutputRead(this, actualR, token);
            candidate.Context.Material.ChargeCopy(512, token);
            return OfdActualReadback.FromActualReader(this, actualR, census, candidate.Context);
        }
        catch { scope.Owner.Fail(); throw; }
    }
    internal OfdVerifiedReadback VerifyActual(OfdActualReadback actual, CancellationToken token)
    {
        try
        {
            actual.ConsumeForActualReceipt(this, out var package, out var census, out var context, token);
            if (!ReferenceEquals(context, candidate.Context) || !ReferenceEquals(census, actualCensus))
                throw new InvalidOperationException("Actual R provenance differs from this written spool.");
            _ = OfdSerializedGuard.CaptureAndObservePayloads(package, context, out var current, token);
            var fonts = receipt.VerifyActual(package, census, current, token);
            current.CompletePrivateIssuance();
            context.Material.ChargeCopy(1024, token);
            var verified = OfdVerifiedReadback.FromActualCompoundPass(this, package, receipt, context);
            scope.Owner.BindVerifiedReadback(this, package, verified, fonts, token);
            return verified;
        }
        catch { scope.Owner.Fail(); throw; }
    }
    internal OfdOwnedCompareScope EnterComparison(OfdVerifiedReadback verified, OfdDocumentPackage actualR,
        CancellationToken token) => scope.Owner.BeginComparison(verified, actualR, token);
    internal async Task<OfdVerifiedStage> CopyToActualStageAsync(OfdOwnedStageScope stageScope,
        OfdOwnedAtomicStage stage, CancellationToken token)
    {
        try
        {
            if (!ReferenceEquals(stageScope.Owner, scope.Owner) || Volatile.Read(ref disposed) != 0)
                throw new InvalidOperationException("Private stage copy has a foreign scope or disposed output.");
            await stage.CopyActualOutputAsync(stageScope, this, file, length, digest, token).ConfigureAwait(false);
            Dispose(); // all required fallible output close/delete happens before claim/rename
            candidate.Context.Material.ChargeCopy(512, token);
            return OfdVerifiedStage.FromActualCopy(stageScope, stage);
        }
        catch { scope.Owner.Fail(); Dispose(); throw; }
    }
    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) != 0) return;
        try { file.Dispose(); candidate.Dispose(); }
        catch { scope.Owner.Fail(); throw; }
    }
}

internal sealed class OfdActualReadback : IDisposable
{
    private readonly OfdOwnedWrittenSpool written;
    private readonly OfdDocumentPackage package;
    private readonly OfdOwnedZipCensus census;
    private readonly OfdControlledEditContext context;
    private int state;
    private OfdActualReadback(OfdOwnedWrittenSpool written, OfdDocumentPackage package,
        OfdOwnedZipCensus census, OfdControlledEditContext context)
    { this.written = written; this.package = package; this.census = census; this.context = context; }
    internal static OfdActualReadback FromActualReader(OfdOwnedWrittenSpool written,
        OfdDocumentPackage package, OfdOwnedZipCensus census, OfdControlledEditContext context) => new(written, package, census, context);
    internal void ConsumeForActualReceipt(OfdOwnedWrittenSpool exact, out OfdDocumentPackage actualR,
        out OfdOwnedZipCensus actualCensus, out OfdControlledEditContext actualContext, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (!ReferenceEquals(exact, written) || Interlocked.CompareExchange(ref state, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or already consumed actual Reader receipt.");
        actualR = package; actualCensus = census; actualContext = context;
    }
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref state, 2, 0) == 0) context.Fail();
    }
}

internal sealed class OfdVerifiedReadback : IDisposable
{
    private readonly OfdOwnedWrittenSpool written;
    private readonly OfdDocumentPackage package;
    private readonly OfdPackageWriter.OwnedWriteReceipt receipt;
    private readonly OfdControlledEditContext context;
    private int state;
    private OfdVerifiedReadback(OfdOwnedWrittenSpool written, OfdDocumentPackage package,
        OfdPackageWriter.OwnedWriteReceipt receipt, OfdControlledEditContext context)
    { this.written = written; this.package = package; this.receipt = receipt; this.context = context; }
    internal static OfdVerifiedReadback FromActualCompoundPass(OfdOwnedWrittenSpool written, OfdDocumentPackage package,
        OfdPackageWriter.OwnedWriteReceipt receipt, OfdControlledEditContext context) => new(written, package, receipt, context);
    internal OfdOwnedCompareScope EnterActualComparison(out OfdDocumentPackage actualR,
        out OfdReadbackRootAddress[] addresses, CancellationToken token)
    {
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (Interlocked.CompareExchange(ref state, 1, 0) != 0) throw new InvalidOperationException("Verified R was already consumed or disposed.");
            var scope = written.EnterComparison(this, package, token);
            addresses = receipt.CopyAddresses(token); actualR = package; return scope;
        }
        catch { context.Fail(); throw; }
    }
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref state, 2, 0) == 0) context.Fail();
    }
}

internal sealed class OfdVerifiedStage
{
    private readonly OfdOwnedStageScope scope;
    private readonly OfdOwnedAtomicStage stage;
    private int state;
    private OfdVerifiedStage(OfdOwnedStageScope scope, OfdOwnedAtomicStage stage) { this.scope = scope; this.stage = stage; }
    internal static OfdVerifiedStage FromActualCopy(OfdOwnedStageScope scope, OfdOwnedAtomicStage stage) => new(scope, stage);
    internal OfdOwnedAtomicStage ConsumeForFinalInput(OfdOwnedModelsLease actualOwner)
    {
        if (!ReferenceEquals(scope.Owner, actualOwner) || Interlocked.CompareExchange(ref state, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or replayed actual stage copy proof.");
        return stage;
    }
}
