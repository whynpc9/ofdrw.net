using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// These claims are private job lifetime authority, never hash-success stamps.
// Only the audited Reader/Packaging/Layout implementations use the stage
// handoffs. An ordinary model/guard pair cannot create an attached source.
internal enum OfdOwnedSavePhase
{
    Bootstrap, ReadingInput, Analyzing, Compiled, Writing, Written,
    ReadingOutput, Comparing, Verified, CopyingStage, CommitClaimed,
    NoopVerified, Completed, Failed, Revoked, Disposed
}

internal sealed class OfdOwnedJobTicket
{
    internal readonly OfdOwnedModelsLease Owner;
    private OfdOwnedJobTicket(OfdOwnedModelsLease owner) { Owner = owner; }
    internal static OfdOwnedJobTicket Issue(OfdOwnedModelsLease owner) => new(owner);
}

internal abstract class OfdOwnedHostIssuer
{
    private OfdOwnedModelsLease? owner;
    protected OfdOwnedHostIssuer() { }
    internal void BindOnce(OfdOwnedModelsLease exactOwner)
    {
        // This rejects public/custom/test-provider implementations. The private
        // Pdf factory and its constructor are still audited trusted code, not a
        // security sandbox against deliberately fabricated friend code.
        var type = GetType();
        if (!type.IsSealed || !type.IsNestedPrivate ||
            type.Assembly.GetName().Name != "Ofdrw.Net.Converter.Pdf" ||
            Interlocked.CompareExchange(ref owner, exactOwner, null) is not null)
            throw new InvalidOperationException("Only one newly issued private geometry host can bind this job.");
    }
    internal void RequireOwner(OfdOwnedModelsLease exactOwner)
    {
        if (!ReferenceEquals(owner, exactOwner))
            throw new InvalidOperationException("Geometry host belongs to another job.");
    }
}

internal abstract class OfdOwnedExpectationIssuer
{
    private OfdOwnedModelsLease? owner;
    protected OfdOwnedExpectationIssuer() { }
    internal void BindOnce(OfdOwnedModelsLease exactOwner)
    {
        if (!GetType().IsSealed || GetType().Assembly.GetName().Name != "Ofdrw.Net.Layout" ||
            Interlocked.CompareExchange(ref owner, exactOwner, null) is not null)
            throw new InvalidOperationException("Geometry expectation was not newly issued by the private compiler.");
    }
    protected void RequireActualReadback(OfdOwnedCompareScope scope, OfdDocumentPackage actual)
    {
        if (owner is null) throw new InvalidOperationException("Unbound geometry expectation.");
        owner.RequireComparison(scope, this, actual);
    }
    protected OfdOwnedStageScope CompleteActualComparison(OfdOwnedCompareScope scope, CancellationToken token)
    {
        if (owner is null) throw new InvalidOperationException("Unbound geometry expectation.");
        // Sole production caller: the sealed Layout comparator's successful
        // tail. Core does not call a virtual comparator or accept a bool/result.
        return owner.CompleteComparison(scope, this, token);
    }
    protected OfdOwnedVerifiedFontSummary RequireVerifiedFontSummary(OfdOwnedCompareScope scope,
        OfdDocumentPackage actualR, OfdFontResource actualFontResource)
    {
        if (owner is null) throw new InvalidOperationException("Unbound geometry expectation.");
        return owner.RequireFontSummary(scope, this, actualR, actualFontResource);
    }
}

// Sole issuer: the private sealed Layout worker's actual successful no-change tail.
internal abstract class OfdOwnedNoopKernelIssuer
{
    private OfdOwnedModelsLease? owner;
    protected OfdOwnedNoopKernelIssuer() { }
    internal void BindOnce(OfdOwnedModelsLease exactOwner)
    {
        var type = GetType();
        if (!type.IsSealed || !type.IsNestedPrivate || type.Assembly.GetName().Name != "Ofdrw.Net.Layout" ||
            Interlocked.CompareExchange(ref owner, exactOwner, null) is not null)
            throw new InvalidOperationException("Only the newly registered private no-change kernel may bind.");
    }
    protected OfdOwnedNoopScope CompleteNoChangedFromActualTail(OfdOwnedAnalysisClaim claim,
        OfdOwnedHostClosure closedHost, CancellationToken token)
    {
        var actual = owner ?? throw new InvalidOperationException("Unregistered no-change kernel.");
        try { return actual.CompleteNoop(claim, closedHost, this, token); }
        catch { actual.Fail(); throw; }
    }
}

internal sealed class OfdOwnedSnapshotCapture : OfdOwnedClaim
{
    private string? hash;
    private long length;
    private int completed;
    private OfdOwnedSnapshotCapture(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedSnapshotCapture Issue(OfdOwnedModelsLease owner) => new(owner);
    internal void Complete(string actualHash, long actualLength, CancellationToken token)
    {
        try
        {
            token.ThrowIfCancellationRequested();
            if (actualHash is null || actualHash.Length != 64 || actualLength < 22 ||
                Interlocked.CompareExchange(ref completed, 1, 0) != 0)
                throw new InvalidDataException("Invalid or repeated actual snapshot capture.");
            foreach (var c in actualHash)
                if (!(c >= '0' && c <= '9') && !(c >= 'a' && c <= 'f'))
                    throw new InvalidDataException("Snapshot SHA256 must be actual lowercase hex.");
            hash = actualHash; length = actualLength;
            Volatile.Write(ref completed, 2);
        }
        catch { Owner.Fail(); throw; }
    }
    internal void RequireCompleted(OfdOwnedModelsLease exactOwner)
    {
        RequireUnconsumed(exactOwner);
        if (Volatile.Read(ref completed) != 2) throw new InvalidOperationException("Incomplete snapshot capture.");
    }
    internal void RequireBytes(string actualHash, long actualLength, OfdOwnedModelsLease exactOwner)
    {
        RequireCompleted(exactOwner);
        if (hash != actualHash || length != actualLength) throw new InvalidDataException("Original snapshot bytes changed.");
    }
}

internal sealed class OfdOwnedNoopScope : OfdOwnedClaim
{
    private OfdOwnedNoopScope(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedNoopScope Issue(OfdOwnedModelsLease owner) => new(owner);
}

internal abstract class OfdOwnedClaim
{
    internal readonly OfdOwnedModelsLease Owner;
    private int consumed;
    private protected OfdOwnedClaim(OfdOwnedModelsLease owner) { Owner = owner; }
    internal void Consume(OfdOwnedModelsLease owner)
    {
        if (!ReferenceEquals(Owner, owner) || Interlocked.CompareExchange(ref consumed, 1, 0) != 0)
            throw new InvalidOperationException("Owned save claim is foreign or already consumed.");
    }
    internal void RequireUnconsumed(OfdOwnedModelsLease owner)
    {
        if (!ReferenceEquals(Owner, owner) || Volatile.Read(ref consumed) != 0)
            throw new InvalidOperationException("Owned save claim is foreign or already consumed.");
    }
}

internal sealed class OfdOwnedReadClaim : OfdOwnedClaim
{
    private OfdOwnedReadClaim(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedReadClaim Issue(OfdOwnedModelsLease owner) => new(owner);
}
internal sealed class OfdOwnedAnalysisClaim : OfdOwnedClaim
{
    private OfdOwnedAnalysisClaim(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedAnalysisClaim Issue(OfdOwnedModelsLease owner) => new(owner);
}
internal sealed class OfdOwnedHostClosure : OfdOwnedClaim
{
    private OfdOwnedHostClosure(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedHostClosure Issue(OfdOwnedModelsLease owner) => new(owner);
}
internal sealed class OfdOwnedWriteScope : OfdOwnedClaim
{
    private OfdOwnedWriteScope(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedWriteScope Issue(OfdOwnedModelsLease owner) => new(owner);
}
internal sealed class OfdOwnedCompareScope : OfdOwnedClaim
{
    private OfdOwnedCompareScope(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedCompareScope Issue(OfdOwnedModelsLease owner) => new(owner);
}
internal sealed class OfdOwnedStageScope : OfdOwnedClaim
{
    private OfdOwnedStageScope(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedStageScope Issue(OfdOwnedModelsLease owner) => new(owner);
}
internal sealed class OfdOwnedCommitClaim : OfdOwnedClaim
{
    private OfdOwnedCommitClaim(OfdOwnedModelsLease owner) : base(owner) { }
    internal static OfdOwnedCommitClaim Issue(OfdOwnedModelsLease owner) => new(owner);
}
// Held only across the synchronous actual rename, never across an await.
// Concurrent abandon/revoke waits for its linearization and cannot mark a
// completed rename as a failed save after the final pre-rename checks.
internal sealed class OfdOwnedCommitGate : IDisposable
{
    private readonly OfdOwnedModelsLease owner;
    private int disposed;
    private OfdOwnedCommitGate(OfdOwnedModelsLease owner) { this.owner = owner; }
    internal static OfdOwnedCommitGate BeginActual(OfdOwnedModelsLease owner) => new(owner);
    internal void RecordActualSuccessfulRename() => owner.RecordSuccessfulRename();
    public void Dispose()
    {
        if (Interlocked.Exchange(ref disposed, 1) == 0) owner.ExitCommitCriticalSection();
    }
}

internal sealed class OfdOwnedTextRootKey
{
    private readonly OfdOwnedModelsLease owner;
    private readonly OfdTextElement original;
    private OfdOwnedTextRootKey(OfdOwnedModelsLease owner, OfdTextElement original)
    { this.owner = owner; this.original = original; }
    internal static OfdOwnedTextRootKey Issue(OfdOwnedModelsLease owner, OfdTextElement original) => new(owner, original);
    internal OfdTextElement ResolveCandidate(OfdOwnedModelsLease exactOwner, OfdOwnedEditCandidate exactCandidate)
    {
        if (!ReferenceEquals(owner, exactOwner)) throw new InvalidOperationException("Foreign root key.");
        owner.RequireWriterCandidate(exactCandidate);
        return exactCandidate.ResolveOwnedClaimedRoot(owner, original);
    }
}

internal enum OfdOwnedPhysicalEntryKind { AcceptedFile, Directory }

internal sealed class OfdOwnedReadAcquisition
{
    private readonly OfdOwnedReadClaim claim;
    private readonly OfdControlledEditContext context;
    private readonly List<(string Name, OfdOwnedPhysicalEntryKind Kind, long Compressed, long Expanded, byte[]? Bytes)> entries = new();
    private readonly HashSet<string> names = new(StringComparer.OrdinalIgnoreCase);
    private int sealedState;
    private readonly object sync = new();
    private OfdOwnedReadAcquisition(OfdOwnedReadClaim claim, OfdControlledEditContext context)
    { this.claim = claim; this.context = context; }
    internal static OfdOwnedReadAcquisition Begin(OfdOwnedReadClaim claim, OfdControlledEditContext context)
    {
        context.Material.ChargeCopy(1024, default);
        return new OfdOwnedReadAcquisition(claim, context);
    }
    internal void Record(string name, OfdOwnedPhysicalEntryKind kind, long compressed, long expanded,
        byte[]? bytes, CancellationToken token)
    {
        lock (sync)
        {
            try { RecordCore(name, kind, compressed, expanded, bytes, token); }
            catch { claim.Owner.Fail(); throw; }
        }
    }
    private void RecordCore(string name, OfdOwnedPhysicalEntryKind kind, long compressed, long expanded,
        byte[]? bytes, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (Volatile.Read(ref sealedState) != 0 || name is null || compressed < 0 || expanded < 0)
            throw new InvalidDataException("Invalid actual ZIP acquisition record.");
        context.Material.ChargeCopy(checked(512L + 8L * name.Length), token);
        context.Material.ChargeWork(checked(16L + 2L * name.Length), token);
        var normalized = name.Replace('\\', '/');
        var unsafeSegment = false; var segmentStart = 0;
        for (var index = 0; index <= normalized.Length; index++)
        {
            if (index != normalized.Length && normalized[index] != '/') continue;
            if (index - segmentStart == 2 && normalized[segmentStart] == '.' && normalized[segmentStart + 1] == '.')
                unsafeSegment = true;
            segmentStart = index + 1;
        }
        if (normalized.Length == 0 || normalized.StartsWith("/", StringComparison.Ordinal) ||
            unsafeSegment || !names.Add(normalized))
            throw new InvalidDataException("Physical ZIP names are empty, unsafe or aliased.");
        if (kind == OfdOwnedPhysicalEntryKind.Directory)
        {
            if (!normalized.EndsWith("/", StringComparison.Ordinal) || expanded != 0 || bytes is not null)
                throw new InvalidDataException("A physical directory cannot hide payload bytes.");
        }
        else if (normalized.EndsWith("/", StringComparison.Ordinal) || bytes is null || bytes.LongLength != expanded)
            throw new InvalidDataException("Actual ZIP payload does not match its physical record.");
        entries.Add((normalized, kind, compressed, expanded, bytes));
    }
    internal void Seal(CancellationToken token)
    {
        lock (sync)
        {
            try
            {
                context.AssertActive(); token.ThrowIfCancellationRequested();
                if (Interlocked.CompareExchange(ref sealedState, 1, 0) != 0)
                    throw new InvalidOperationException("ZIP acquisition was already sealed.");
            }
            catch { claim.Owner.Fail(); throw; }
        }
    }
    internal void RequireActualSource(OfdOwnedReadClaim exactClaim, OfdDocumentPackage package,
        OfdControlledEditContext exactContext, CancellationToken token)
    {
        if (!ReferenceEquals(claim, exactClaim) || !ReferenceEquals(context, exactContext) ||
            Volatile.Read(ref sealedState) != 1)
            throw new InvalidOperationException("Source lacks its actual sealed loader acquisition.");
        var count = 0;
        foreach (var entry in entries)
        {
            token.ThrowIfCancellationRequested(); context.Material.ChargeWork(16, token);
            if (entry.Kind == OfdOwnedPhysicalEntryKind.Directory) continue;
            count++;
            if (!package.PreservedEntries.TryGetValue(entry.Name, out var actual) || !ReferenceEquals(actual, entry.Bytes))
                throw new InvalidDataException("Reader source does not own the actually loaded entry arrays.");
        }
        if (package.PreservedEntries.Count != count || package.SourceIdentity is null)
            throw new InvalidDataException("Reader source inventory is not the actual complete acquisition.");
    }
}

internal sealed class OfdOwnedModelsLease
{
    private readonly object commitSync = new();
    private readonly OfdOwnedIoTelemetry telemetry;
    private readonly long budgetBytes;
    private int inputIoBound, outputIoBound, stageIoBound, logicalURecorded;
    private readonly OfdControlledEditContext context;
    private readonly object jobIdentity = new();
    private int noopBound, snapshotIssued, actualAnalysisStarted;
    private int phase, everBorrowed, intentBound, hostBound, analysisEntered, acquisitionIssued, hostClosed, stageBound, readCompleted;
    private string? inputPath, outputPath;
    private OfdOwnedJobTicket? ticket;
    private OfdOwnedReadClaim? readClaim;
    private OfdOwnedReadAcquisition? acquisition;
    private OfdOwnedSnapshotCapture? snapshotCapture;
    private OfdOwnedNoopKernelIssuer? noopKernel;
    private OfdOwnedNoopScope? noopScope;
    private OfdOwnedAnalysisClaim? analysisClaim;
    private OfdOwnedHostIssuer? host;
    private OfdOwnedHostClosure? closure;
    private OfdOwnedExpectationIssuer? expectation;
    private OfdRetainedTextEditBatch? batch;
    private OfdOwnedWriteScope? writeScope;
    private OfdDocumentPackage? source, readback;
    private OfdSerializedGuard? sourceGuard;
    private OfdOwnedEditCandidate? candidate;
    private OfdOwnedTextRootKey[] rootKeys = Array.Empty<OfdOwnedTextRootKey>();
    private object? writtenIdentity, verifiedIdentity, stageIdentity;
    private OfdOwnedCompareScope? compareScope;
    private OfdOwnedStageScope? stageScope;
    private OfdOwnedCommitClaim? commitClaim;
    private Dictionary<OfdFontResource, OfdOwnedVerifiedFontSummary>? verifiedFonts;
    private int comparisonAdmitted;

    private OfdOwnedModelsLease(OfdModelLimits limits, CancellationToken token)
    {
        var snapshot = limits.Snapshot();
        if (snapshot.MaxMaterializedBytes > 512L * 1024 * 1024)
            throw new ArgumentOutOfRangeException(nameof(limits), "Owned saves retain the 512 MiB ceiling.");
        context = new OfdControlledEditContext(snapshot);
        context.Material.ChargeCopy(4096 + 65_536, token); // ticket + fixed numeric observations and 32 once-fill snapshots
        budgetBytes = snapshot.MaxMaterializedBytes;
        telemetry = OfdOwnedIoTelemetry.CreateAfterPrepayment();
    }
    internal static OfdOwnedJobTicket Begin(OfdModelLimits limits, CancellationToken token)
    {
        if (limits is null) throw new ArgumentNullException(nameof(limits));
        token.ThrowIfCancellationRequested();
        var owner = new OfdOwnedModelsLease(limits, token);
        owner.ticket = OfdOwnedJobTicket.Issue(owner); return owner.ticket;
    }
    private void Require(OfdOwnedSavePhase expected, CancellationToken token = default)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (Volatile.Read(ref everBorrowed) != 0 || Volatile.Read(ref phase) != (int)expected)
            throw new InvalidOperationException("Owned save job is revoked, failed or in a different phase.");
    }
    private void Move(OfdOwnedSavePhase before, OfdOwnedSavePhase after, CancellationToken token = default)
    {
        Require(before, token);
        if (Interlocked.CompareExchange(ref phase, (int)after, (int)before) != (int)before)
            throw new InvalidOperationException("Owned save job already has an active stage.");
    }
    internal void Fail()
    {
      lock (commitSync)
      {
        while (true)
        {
            var before = Volatile.Read(ref phase);
            if (before == (int)OfdOwnedSavePhase.Completed) return;
            if (Interlocked.CompareExchange(ref phase, (int)OfdOwnedSavePhase.Failed, before) != before) continue;
            context.Fail(); return;
        }
      }
    }
    internal void Revoke()
    {
      lock (commitSync)
      {
        if (Volatile.Read(ref phase) == (int)OfdOwnedSavePhase.Completed) return;
        Interlocked.Exchange(ref everBorrowed, 1); context.Fail();
        Interlocked.Exchange(ref phase, (int)OfdOwnedSavePhase.Revoked);
      }
    }
    internal void BindFileIntent(OfdOwnedJobTicket exact, string input, string output, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Bootstrap, token);
        if (!ReferenceEquals(ticket, exact) || Interlocked.CompareExchange(ref intentBound, 1, 0) != 0 || string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output))
            throw new InvalidOperationException("File intent is missing, foreign or already bound.");
        context.Material.ChargeCopy(checked(1024L + 8L * (input.Length + (long)output.Length)), token);
        RequireCanonicalPath(input, token); RequireCanonicalPath(output, token);
        if (string.Equals(input, output, StringComparison.OrdinalIgnoreCase))
            throw new InvalidDataException("Owned save cannot replace its original input or an alias.");
        context.Material.StoreString(input); context.Material.StoreString(output);
        inputPath = input; outputPath = output;
    }
    private void RequireCanonicalPath(string path, CancellationToken token)
    {
        context.Material.ChargeWork(checked(128L + 4L * path.Length), token);
        if (!Path.IsPathRooted(path) || Path.GetFullPath(path) != path)
            throw new InvalidDataException("Owned save requires an exact canonical absolute file intent.");
        string? current = path;
        while (!string.IsNullOrEmpty(current))
        {
            context.Material.ChargeCopy(checked(256L + 4L * current!.Length), token);
            context.Material.ChargeWork(checked(128L + 2L * current.Length), token);
            try
            {
                if ((File.GetAttributes(current) & FileAttributes.ReparsePoint) != 0)
                    throw new InvalidDataException("Owned file intent cannot traverse a symbolic link or reparse point.");
            }
            catch (FileNotFoundException) { }
            catch (DirectoryNotFoundException) { }
            current = Path.GetDirectoryName(current);
        }
    }
    internal void BindHost(OfdOwnedJobTicket exact, OfdOwnedHostIssuer actual)
    {
        Require(OfdOwnedSavePhase.Bootstrap);
        if (!ReferenceEquals(ticket, exact) || Interlocked.CompareExchange(ref hostBound, 1, 0) != 0 || actual is null)
            throw new InvalidOperationException("Geometry host is foreign or already bound.");
        actual.BindOnce(this); host = actual;
    }
    internal void ReserveBootstrap(OfdOwnedJobTicket exact, long bytes, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Bootstrap, token);
        if (!ReferenceEquals(ticket, exact) || bytes < 0) throw new InvalidOperationException("Foreign bootstrap.");
        context.Material.ChargeCopy(bytes, token);
    }
    internal void BindNoopKernel(OfdOwnedJobTicket exact, OfdOwnedNoopKernelIssuer actual)
    {
        Require(OfdOwnedSavePhase.Bootstrap);
        if (!ReferenceEquals(ticket, exact) || actual is null || Interlocked.CompareExchange(ref noopBound, 1, 0) != 0)
            throw new InvalidOperationException("No-change kernel is foreign or already registered.");
        actual.BindOnce(this); noopKernel = actual;
    }
    internal void ReserveBootstrapWork(OfdOwnedJobTicket exact, long units, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Bootstrap, token);
        if (!ReferenceEquals(ticket, exact) || units < 0) throw new InvalidOperationException("Foreign bootstrap work.");
        context.Material.ChargeWork(units, token);
    }
    internal void AdmitBootstrapString(OfdOwnedJobTicket exact, string value, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Bootstrap, token);
        if (!ReferenceEquals(ticket, exact) || value is null) throw new InvalidOperationException("Foreign bootstrap string.");
        context.Material.StoreString(value);
    }
    internal void RequireInitialAcquisition(OfdOwnedReadClaim exact, OfdControlledEditContext actualContext,
        string actualPath, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.ReadingInput, token); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(readClaim, exact) || !ReferenceEquals(context, actualContext) || inputPath != actualPath)
            throw new InvalidOperationException("Foreign initial snapshot acquisition.");
    }
    internal OfdOwnedSnapshotCapture BeginSnapshot(OfdOwnedReadClaim exact)
    {
        Require(OfdOwnedSavePhase.ReadingInput); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(readClaim, exact) || Interlocked.CompareExchange(ref snapshotIssued, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or repeated original snapshot.");
        context.Material.ChargeCopy(512, default);
        snapshotCapture = OfdOwnedSnapshotCapture.Issue(this); return snapshotCapture;
    }
    internal OfdOwnedReadClaim BeginRead(OfdOwnedJobTicket exact, out OfdControlledEditContext actualContext)
    {
        if (!ReferenceEquals(ticket, exact) || inputPath is null || outputPath is null || host is null)
            throw new InvalidOperationException("Initial read requires this fresh configured job and host.");
        Move(OfdOwnedSavePhase.Bootstrap, OfdOwnedSavePhase.ReadingInput);
        context.Material.ChargeCopy(256, default); readClaim = OfdOwnedReadClaim.Issue(this);
        actualContext = context; return readClaim;
    }
    internal OfdOwnedReadAcquisition BeginAcquisition(OfdOwnedReadClaim exact)
    {
        Require(OfdOwnedSavePhase.ReadingInput); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(readClaim, exact) || Interlocked.CompareExchange(ref acquisitionIssued, 1, 0) != 0)
            throw new InvalidOperationException("Actual initial loader acquisition was already issued or is foreign.");
        acquisition = OfdOwnedReadAcquisition.Begin(exact, context); return acquisition;
    }
    internal OfdOwnedModelsLease CompleteRead(OfdOwnedReadClaim exact, OfdDocumentPackage actual,
        OfdOwnedReadAcquisition actualAcquisition, OfdOwnedSnapshotCapture actualSnapshot, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.ReadingInput, token);
        if (!ReferenceEquals(exact, readClaim) || !ReferenceEquals(actualAcquisition, acquisition) || Interlocked.CompareExchange(ref readCompleted, 1, 0) != 0)
            throw new InvalidOperationException("Actual initial Reader completion is foreign or repeated.");
        if (!ReferenceEquals(snapshotCapture, actualSnapshot)) throw new InvalidOperationException("Foreign original snapshot capture.");
        actualSnapshot.RequireCompleted(this);
        actualAcquisition.RequireActualSource(exact, actual, context, token); exact.Consume(this);
        sourceGuard = context.CaptureSource(actual, token); source = actual;
        Move(OfdOwnedSavePhase.ReadingInput, OfdOwnedSavePhase.Analyzing, token); return this;
    }
    internal OfdOwnedAnalysisClaim EnterAnalysis(OfdOwnedHostIssuer actualHost, out OfdDocumentPackage actualSource,
        out OfdControlledEditContext actualContext, out OfdSerializedGuard actualGuard)
    {
        Require(OfdOwnedSavePhase.Analyzing);
        if (!ReferenceEquals(host, actualHost) || Interlocked.CompareExchange(ref analysisEntered, 1, 0) != 0) throw new InvalidOperationException("Foreign or repeated analysis.");
        actualHost.RequireOwner(this); context.Material.ChargeCopy(256, default);
        analysisClaim = OfdOwnedAnalysisClaim.Issue(this);
        actualSource = source!; actualContext = context; actualGuard = sourceGuard!; return analysisClaim;
    }
    internal OfdOwnedHostClosure CompleteHost(OfdOwnedAnalysisClaim exact, OfdOwnedHostIssuer actualHost, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Analyzing, token); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(exact, analysisClaim) || !ReferenceEquals(host, actualHost) ||
            Volatile.Read(ref actualAnalysisStarted) != 1 || Interlocked.CompareExchange(ref hostClosed, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or repeated host closure.");
        actualHost.RequireOwner(this); context.Material.ChargeCopy(256, token);
        closure = OfdOwnedHostClosure.Issue(this); return closure;
    }
    internal void RequireActualAnalysis(OfdOwnedAnalysisClaim exact, OfdOwnedHostIssuer actualHost,
        OfdDocumentPackage actualSource, OfdControlledEditContext actualContext, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Analyzing, token); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(exact, analysisClaim) || !ReferenceEquals(host, actualHost) ||
            !ReferenceEquals(source, actualSource) || !ReferenceEquals(context, actualContext))
            throw new InvalidOperationException("Actual analysis has a foreign claim, host, source or context.");
        actualHost.RequireOwner(this);
        if (Interlocked.CompareExchange(ref actualAnalysisStarted, 1, 0) != 0)
            throw new InvalidOperationException("Actual analysis has already started.");
    }
    internal OfdOwnedWriteScope SealBatch(OfdOwnedAnalysisClaim exact, OfdOwnedHostClosure closed,
        OfdOwnedExpectationIssuer actualExpectation, OfdRetainedTextEditBatch actualBatch,
        OfdTextElement[] completeRoots, out OfdOwnedTextRootKey[] issuedKeys, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Analyzing, token);
        if (!ReferenceEquals(exact, analysisClaim) || !ReferenceEquals(closed, closure) || completeRoots is null || actualExpectation is null)
            throw new InvalidOperationException("Compiled batch lacks actual host closure and analysis ownership.");
        exact.RequireUnconsumed(this); closed.RequireUnconsumed(this);
        actualBatch.RequireUnclaimedBinding(source!, sourceGuard!, context);
        context.Material.ChargeCopy(checked(1024L + 256L * completeRoots.Length), token);
        actualBatch.RequireCompleteAffectedRoots(completeRoots, token);
        rootKeys = new OfdOwnedTextRootKey[completeRoots.Length];
        for (var i = 0; i < completeRoots.Length; i++)
        {
            token.ThrowIfCancellationRequested(); context.Material.ChargeWork(16, token);
            rootKeys[i] = OfdOwnedTextRootKey.Issue(this, completeRoots[i]);
        }
        actualExpectation.BindOnce(this); expectation = actualExpectation; batch = actualBatch;
        exact.Consume(this); closed.Consume(this);
        Move(OfdOwnedSavePhase.Analyzing, OfdOwnedSavePhase.Compiled, token);
        writeScope = OfdOwnedWriteScope.Issue(this);
        // The caller owns only a copied opaque-key array, never source owners.
        context.Material.ChargeCopy(checked(128L + IntPtr.Size * (long)rootKeys.Length), token);
        issuedKeys = rootKeys.ToArray(); return writeScope;
    }
    internal OfdOwnedNoopScope CompleteNoop(OfdOwnedAnalysisClaim exact, OfdOwnedHostClosure closed,
        OfdOwnedNoopKernelIssuer issuer, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Analyzing, token);
        if (!ReferenceEquals(exact, analysisClaim) || !ReferenceEquals(closed, closure) ||
            !ReferenceEquals(issuer, noopKernel) || batch is not null)
            throw new InvalidOperationException("No-change completion lacks actual analysis, host and kernel ownership.");
        exact.RequireUnconsumed(this); closed.RequireUnconsumed(this);
        snapshotCapture!.RequireCompleted(this);
        // Actual successful no-change tail, with a current S, never a cached capture.
        sourceGuard!.AssertUnchanged(source!, context, token);
        context.Material.ChargeCopy(256, token);
        exact.Consume(this); closed.Consume(this);
        Move(OfdOwnedSavePhase.Analyzing, OfdOwnedSavePhase.NoopVerified, token);
        telemetry.RecordActualNoop();
        noopScope = OfdOwnedNoopScope.Issue(this); return noopScope;
    }
    internal OfdControlledEditContext BeginNoopStage(OfdOwnedNoopScope exact,
        OfdOwnedSnapshotCapture snapshot, object actualStage, string destination, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.NoopVerified, token); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(exact, noopScope) || !ReferenceEquals(snapshot, snapshotCapture) ||
            actualStage is null || Interlocked.CompareExchange(ref stageBound, 1, 0) != 0 || outputPath != destination)
            throw new InvalidOperationException("Foreign no-change stage, snapshot or output intent.");
        snapshot.RequireCompleted(this); stageIdentity = actualStage; return context;
    }
    internal void BeginNoopCopy(OfdOwnedNoopScope exact, OfdOwnedSnapshotCapture snapshot,
        object actualStage, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.NoopVerified, token);
        if (!ReferenceEquals(exact, noopScope) || !ReferenceEquals(snapshot, snapshotCapture) || !ReferenceEquals(actualStage, stageIdentity))
            throw new InvalidOperationException("Foreign no-change copy.");
        snapshot.RequireCompleted(this); exact.Consume(this);
        Move(OfdOwnedSavePhase.NoopVerified, OfdOwnedSavePhase.CopyingStage, token);
    }
    internal void AssertFinalNoopSource(OfdOwnedNoopScope exact, OfdOwnedSnapshotCapture snapshot,
        object actualStage, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.CopyingStage, token);
        if (!ReferenceEquals(exact, noopScope) || !ReferenceEquals(snapshot, snapshotCapture) || !ReferenceEquals(actualStage, stageIdentity))
            throw new InvalidOperationException("Foreign final no-change validation.");
        sourceGuard!.AssertUnchanged(source!, context, token);
    }
    internal OfdOwnedEditCandidate Clone(OfdOwnedWriteScope exact, OfdRetainedTextEditBatch actualBatch,
        out OfdSerializedGuard guard, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Compiled, token);
        if (!ReferenceEquals(exact, writeScope) || !ReferenceEquals(actualBatch, batch))
            throw new InvalidOperationException("Private Writer scope or compiler batch is foreign.");
        exact.Consume(this); Move(OfdOwnedSavePhase.Compiled, OfdOwnedSavePhase.Writing, token);
        // Reuse the complete original implementation and all three retained
        // source/candidate frontiers. No extra candidate import is introduced.
        candidate = OfdOwnedEditCandidate.ClonePatchAndClaim(source!, sourceGuard!, context, actualBatch, token, out guard);
        return candidate;
    }
    internal OfdOwnedTextRootKey[] GetWriterRootKeys(OfdOwnedWriteScope exact, OfdOwnedEditCandidate exactCandidate, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Writing, token);
        if (!ReferenceEquals(exact, writeScope) || !ReferenceEquals(exactCandidate, candidate))
            throw new InvalidOperationException("Foreign Writer root projection.");
        context.Material.ChargeCopy(checked(128L + IntPtr.Size * (long)rootKeys.Length), token); return rootKeys.ToArray();
    }
    internal void CompleteWrite(OfdOwnedWriteScope exact, object actualWrittenIdentity, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Writing, token);
        if (!ReferenceEquals(exact, writeScope) || writtenIdentity is not null || actualWrittenIdentity is null)
            throw new InvalidOperationException("Foreign or repeated written spool.");
        writtenIdentity = actualWrittenIdentity; Move(OfdOwnedSavePhase.Writing, OfdOwnedSavePhase.Written, token);
    }
    internal OfdControlledEditContext BeginOutputRead(object actualWrittenIdentity, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Written, token);
        if (!ReferenceEquals(writtenIdentity, actualWrittenIdentity)) throw new InvalidOperationException("Foreign output read.");
        Move(OfdOwnedSavePhase.Written, OfdOwnedSavePhase.ReadingOutput, token); return context;
    }
    internal void CompleteOutputRead(object actualWrittenIdentity, OfdDocumentPackage actualReadback, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.ReadingOutput, token);
        if (!ReferenceEquals(writtenIdentity, actualWrittenIdentity) || readback is not null)
            throw new InvalidOperationException("Foreign or repeated actual output readback.");
        readback = actualReadback;
    }
    internal void BindVerifiedReadback(object actualWrittenIdentity, OfdDocumentPackage actualReadback,
        object actualVerifiedIdentity, Dictionary<OfdFontResource, OfdOwnedVerifiedFontSummary> actualFonts, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.ReadingOutput, token);
        if (!ReferenceEquals(writtenIdentity, actualWrittenIdentity) || !ReferenceEquals(readback, actualReadback) || verifiedIdentity is not null)
            throw new InvalidOperationException("Foreign or repeated serialized readback verification.");
        context.Material.ChargeCopy(checked(512L + 256L * actualFonts.Count), token);
        verifiedFonts = new Dictionary<OfdFontResource, OfdOwnedVerifiedFontSummary>(actualFonts);
        verifiedIdentity = actualVerifiedIdentity;
    }
    internal OfdOwnedCompareScope BeginComparison(object exactVerifiedIdentity, OfdDocumentPackage actualReadback, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.ReadingOutput, token);
        if (!ReferenceEquals(verifiedIdentity, exactVerifiedIdentity) || !ReferenceEquals(readback, actualReadback) || compareScope is not null)
            throw new InvalidOperationException("Actual comparison lacks verified readback ownership.");
        context.Material.ChargeCopy(256, token); Move(OfdOwnedSavePhase.ReadingOutput, OfdOwnedSavePhase.Comparing, token);
        compareScope = OfdOwnedCompareScope.Issue(this); return compareScope;
    }
    internal void RequireComparison(OfdOwnedCompareScope exact, OfdOwnedExpectationIssuer issuer, OfdDocumentPackage actualReadback)
    {
        Require(OfdOwnedSavePhase.Comparing); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(exact, compareScope) || !ReferenceEquals(issuer, expectation) || !ReferenceEquals(actualReadback, readback))
            throw new InvalidOperationException("Comparison does not bind this actual R and compiler expectation.");
        Volatile.Write(ref comparisonAdmitted, 1);
    }
    internal OfdOwnedVerifiedFontSummary RequireFontSummary(OfdOwnedCompareScope exact,
        OfdOwnedExpectationIssuer issuer, OfdDocumentPackage actualR, OfdFontResource font)
    {
        Require(OfdOwnedSavePhase.Comparing); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(exact, compareScope) || !ReferenceEquals(issuer, expectation) || !ReferenceEquals(actualR, readback) ||
            Volatile.Read(ref comparisonAdmitted) != 1 || verifiedFonts is null || !verifiedFonts.TryGetValue(font, out var summary))
            throw new InvalidOperationException("Font lacks this current actual-R compound evidence and admitted comparator.");
        return summary;
    }
    internal OfdOwnedStageScope CompleteComparison(OfdOwnedCompareScope exact, OfdOwnedExpectationIssuer issuer, CancellationToken token)
    {
        RequireComparison(exact, issuer, readback!); token.ThrowIfCancellationRequested(); exact.Consume(this);
        context.Material.ChargeCopy(256, token); Move(OfdOwnedSavePhase.Comparing, OfdOwnedSavePhase.Verified, token);
        stageScope = OfdOwnedStageScope.Issue(this); return stageScope;
    }
    internal OfdControlledEditContext BeginStage(OfdOwnedStageScope exact, object actualStageIdentity, string destination, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Verified, token); exact.RequireUnconsumed(this);
        if (!ReferenceEquals(exact, stageScope) || Interlocked.CompareExchange(ref stageBound, 1, 0) != 0 || outputPath != destination)
            throw new InvalidOperationException("Atomic stage is foreign, repeated or has a different destination intent.");
        stageIdentity = actualStageIdentity; return context;
    }
    internal void BeginStageCopy(OfdOwnedStageScope exact, object actualWrittenIdentity, object actualStageIdentity, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Verified, token);
        if (!ReferenceEquals(stageScope, exact) || !ReferenceEquals(writtenIdentity, actualWrittenIdentity) || !ReferenceEquals(stageIdentity, actualStageIdentity))
            throw new InvalidOperationException("Stage copy has foreign proof owners.");
        exact.Consume(this); Move(OfdOwnedSavePhase.Verified, OfdOwnedSavePhase.CopyingStage, token);
    }
    internal OfdOwnedCommitClaim ClaimCommit(object actualStageIdentity, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.CopyingStage, token);
        if (!ReferenceEquals(stageIdentity, actualStageIdentity) || commitClaim is not null)
            throw new InvalidOperationException("Commit has a foreign or repeated stage.");
        // The actual final-input path/hash check has completed before this
        // Packaging claim. Changed saves retain a fresh whole source-model S
        // here; original-file bytes cannot substitute for this model frontier.
        // Noop already performs its exact scope/snapshot-bound final S.
        if (stageScope is not null) sourceGuard!.AssertUnchanged(source!, context, token);
        else if (noopScope is null) throw new InvalidOperationException("Commit lacks an actual changed or no-change scope.");
        context.Material.ChargeCopy(256, token); Move(OfdOwnedSavePhase.CopyingStage, OfdOwnedSavePhase.CommitClaimed, token);
        commitClaim = OfdOwnedCommitClaim.Issue(this); return commitClaim;
    }
    internal void ConsumeCommit(OfdOwnedCommitClaim exact, object actualStageIdentity, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.CommitClaimed, token);
        if (!ReferenceEquals(commitClaim, exact) || !ReferenceEquals(stageIdentity, actualStageIdentity))
            throw new InvalidOperationException("Commit claim is foreign.");
        exact.Consume(this);
    }
    internal void CompleteCommit(CancellationToken token)
    {
        Move(OfdOwnedSavePhase.CommitClaimed, OfdOwnedSavePhase.Completed, token);
    }
    internal void RecordSuccessfulRename()
    {
        // Actual Packaging rename is the sole caller. All fallible eligibility,
        // cancellation and I/O admission happened before rename.
        Interlocked.Exchange(ref phase, (int)OfdOwnedSavePhase.Completed);
    }
    internal OfdOwnedCommitGate EnterCommitCriticalSection(OfdOwnedCommitClaim exact,
        object actualStageIdentity, CancellationToken token)
    {
        Monitor.Enter(commitSync);
        try
        {
            ConsumeCommit(exact, actualStageIdentity, token);
            context.Material.ChargeCopy(256, token);
            return OfdOwnedCommitGate.BeginActual(this);
        }
        catch { Fail(); Monitor.Exit(commitSync); throw; }
    }
    internal void ExitCommitCriticalSection() => Monitor.Exit(commitSync);
    internal void RequireInputIntent(string actualPath, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (everBorrowed != 0 || inputPath != actualPath) throw new InvalidOperationException("Foreign original input path.");
        RequireCanonicalPath(actualPath, token);
    }
    internal void RequireCurrentDestination(CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        RequireCanonicalPath(outputPath ?? throw new InvalidOperationException("Unbound destination."), token);
    }
    internal OfdOwnedIoObservation BindActualInputObservation(OfdOwnedReadClaim exact,
        OfdOwnedSnapshotCapture capture, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.ReadingInput, token);
        if (!ReferenceEquals(readClaim, exact) || !ReferenceEquals(snapshotCapture, capture) ||
            Interlocked.CompareExchange(ref inputIoBound, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or duplicate actual input observation owner.");
        return telemetry.Input;
    }
    internal OfdOwnedIoObservation BindActualOutputObservation(OfdOwnedWriteScope exact,
        OfdOwnedEditCandidate actualCandidate, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Writing, token);
        if (!ReferenceEquals(writeScope, exact) || !ReferenceEquals(candidate, actualCandidate) ||
            Interlocked.CompareExchange(ref outputIoBound, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or duplicate actual output observation owner.");
        return telemetry.Output;
    }
    internal OfdOwnedIoObservation BindActualStageObservation(object actualStage, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        var actualPhase = Volatile.Read(ref phase);
        if ((actualPhase != (int)OfdOwnedSavePhase.Verified && actualPhase != (int)OfdOwnedSavePhase.NoopVerified) ||
            !ReferenceEquals(stageIdentity, actualStage) || Interlocked.CompareExchange(ref stageIoBound, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or duplicate actual stage observation owner.");
        return telemetry.Stage;
    }
    internal void RecordActualLogicalWriterReservation(OfdOwnedWriteScope exact,
        OfdOwnedEditCandidate actualCandidate, long actualReservation, CancellationToken token)
    {
        Require(OfdOwnedSavePhase.Writing, token);
        if (!ReferenceEquals(writeScope, exact) || !ReferenceEquals(candidate, actualCandidate) || actualReservation < 0 ||
            Interlocked.CompareExchange(ref logicalURecorded, 1, 0) != 0)
            throw new InvalidOperationException("Foreign or repeated actual Writer reservation observation.");
        telemetry.RecordLogicalWriterReservation(actualReservation);
    }
    internal OfdOwnedSaveStatistics CaptureStatistics(OfdOwnedJobTicket exact)
    {
        if (!ReferenceEquals(ticket, exact)) throw new InvalidOperationException("Foreign numeric observation ticket.");
        // No allocation, quota request, virtual I/O/property, model walk or active
        // context requirement. This remains safe after the first failure.
        return telemetry.CaptureActual(Volatile.Read(ref phase), budgetBytes, context.Material);
    }
    internal void RequireWriterCandidate(OfdOwnedEditCandidate actual)
    {
        Require(OfdOwnedSavePhase.Writing);
        if (!ReferenceEquals(candidate, actual)) throw new InvalidOperationException("Foreign Writer candidate.");
    }

}

internal static class OfdOwnedSaveAuthority
{
    private static T Run<T>(OfdOwnedModelsLease owner, Func<T> action)
    {
        // All callers below supply fixed Core lambdas, never caller delegates.
        try { return action(); } catch { owner.Fail(); throw; }
    }
    internal static OfdOwnedJobTicket BeginFreshJob(OfdModelLimits limits, CancellationToken token)
        => OfdOwnedModelsLease.Begin(limits, token);
    internal static void BindFileIntent(OfdOwnedJobTicket ticket, string inputPath, string outputPath, CancellationToken token)
        => Run(ticket.Owner, () => { ticket.Owner.BindFileIntent(ticket, inputPath, outputPath, token); return 0; });
    internal static void BindHost(OfdOwnedJobTicket ticket, OfdOwnedHostIssuer exactHost)
        => Run(ticket.Owner, () => { ticket.Owner.BindHost(ticket, exactHost); return 0; });
    internal static void ReserveBootstrap(OfdOwnedJobTicket ticket, long bytes, CancellationToken token)
        => Run(ticket.Owner, () => { ticket.Owner.ReserveBootstrap(ticket, bytes, token); return 0; });
    internal static OfdOwnedReadClaim BeginInputRead(OfdOwnedJobTicket ticket, out OfdControlledEditContext context)
    {
        try { return ticket.Owner.BeginRead(ticket, out context); } catch { ticket.Owner.Fail(); throw; }
    }
    internal static void BindNoopKernel(OfdOwnedJobTicket ticket, OfdOwnedNoopKernelIssuer actualKernel)
        => Run(ticket.Owner, () => { ticket.Owner.BindNoopKernel(ticket, actualKernel); return 0; });
    internal static void ReserveBootstrapWork(OfdOwnedJobTicket ticket, long units, CancellationToken token)
        => Run(ticket.Owner, () => { ticket.Owner.ReserveBootstrapWork(ticket, units, token); return 0; });
    internal static void AdmitBootstrapString(OfdOwnedJobTicket ticket, string value, CancellationToken token)
        => Run(ticket.Owner, () => { ticket.Owner.AdmitBootstrapString(ticket, value, token); return 0; });
    internal static OfdOwnedSnapshotCapture BeginActualSnapshotCapture(OfdOwnedReadClaim claim)
        => Run(claim.Owner, () => claim.Owner.BeginSnapshot(claim));
    internal static void CompleteActualSnapshotCapture(OfdOwnedSnapshotCapture capture,
        string actualRawZipSha256Hex, long actualLength, CancellationToken token)
        => Run(capture.Owner, () => { capture.Complete(actualRawZipSha256Hex, actualLength, token); return 0; });
    internal static OfdOwnedReadAcquisition BeginLoaderAcquisition(OfdOwnedReadClaim claim)
        => Run(claim.Owner, () => claim.Owner.BeginAcquisition(claim));
    internal static void RecordActualLoadedEntry(OfdOwnedReadAcquisition acquisition, string exactPhysicalName,
        OfdOwnedPhysicalEntryKind kind, long compressedLength, long expandedLength, byte[]? exactAcceptedBuffer, CancellationToken token)
        => acquisition.Record(exactPhysicalName, kind, compressedLength, expandedLength, exactAcceptedBuffer, token);
    internal static void SealActualLoaderAcquisition(OfdOwnedReadAcquisition acquisition, CancellationToken token)
        => acquisition.Seal(token);
    internal static OfdOwnedModelsLease CompleteInputRead(OfdOwnedReadClaim claim, OfdDocumentPackage actualReaderPackage,
        OfdOwnedReadAcquisition actualLoaderAcquisition, OfdOwnedSnapshotCapture actualSnapshotCapture, CancellationToken token)
        => Run(claim.Owner, () => claim.Owner.CompleteRead(claim, actualReaderPackage, actualLoaderAcquisition, actualSnapshotCapture, token));
    internal static OfdOwnedAnalysisClaim EnterAnalysis(OfdOwnedModelsLease lease, OfdOwnedHostIssuer actualHost,
        out OfdDocumentPackage source, out OfdControlledEditContext context, out OfdSerializedGuard sourceGuard)
    {
        try { return lease.EnterAnalysis(actualHost, out source, out context, out sourceGuard); } catch { lease.Fail(); throw; }
    }
    internal static OfdOwnedHostClosure CompleteClosedHost(OfdOwnedAnalysisClaim claim, OfdOwnedHostIssuer exactHost, CancellationToken token)
        => Run(claim.Owner, () => claim.Owner.CompleteHost(claim, exactHost, token));
    internal static void RequireActualAnalysis(OfdOwnedAnalysisClaim claim, OfdOwnedHostIssuer exactHost,
        OfdDocumentPackage actualSource, OfdControlledEditContext actualContext, CancellationToken token)
        => Run(claim.Owner, () => { claim.Owner.RequireActualAnalysis(claim, exactHost, actualSource, actualContext, token); return 0; });
    internal static OfdOwnedWriteScope SealCompiledBatch(OfdOwnedAnalysisClaim claim, OfdOwnedHostClosure actualClosedHost,
        OfdOwnedExpectationIssuer exactExpectation, OfdRetainedTextEditBatch exactBatch,
        OfdTextElement[] completeAffectedPageRoots, out OfdOwnedTextRootKey[] keys, CancellationToken token)
    {
        try { return claim.Owner.SealBatch(claim, actualClosedHost, exactExpectation, exactBatch, completeAffectedPageRoots, out keys, token); }
        catch { claim.Owner.Fail(); throw; }
    }
    internal static OfdOwnedEditCandidate ClonePatchAndClaim(OfdOwnedWriteScope scope, OfdRetainedTextEditBatch exactBatch,
        out OfdSerializedGuard candidateGuard, CancellationToken token)
    {
        try { return scope.Owner.Clone(scope, exactBatch, out candidateGuard, token); } catch { scope.Owner.Fail(); throw; }
    }
    internal static OfdOwnedSaveStatistics CaptureStatistics(OfdOwnedJobTicket ticket) => ticket.Owner.CaptureStatistics(ticket);
    internal static void RevokeForExternalBorrow(OfdOwnedModelsLease lease) => lease.Revoke();
    internal static void Fail(OfdOwnedJobTicket ticket) => ticket.Owner.Fail();
}
