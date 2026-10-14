using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Layout.Editing.Internal;

internal sealed class OfdTextReplacementOperation : IDisposable
{
    internal OfdControlledEditContext Context { get; }
    internal OfdDocumentPackage? Source { get; private set; }
    internal OfdSerializedGuard? SourceGuard { get; private set; }
    internal OfdOwnedEditCandidate? Candidate { get; set; }
    internal OfdSerializedGuard? CandidateGuard { get; set; }
    internal OfdTextReplacementOptions Options { get; }
    private List<OfdTextReplacementPhase> phases = null!;
    private List<OfdKeywordLocatorStatistics> analyses = null!;
    private OfdTextReplacementTraceCell[] trace = Array.Empty<OfdTextReplacementTraceCell>();
    private int traceCount, currentTrace = -1, traceSnapshots;
    private (long Storage, long Copy, long Output, long Work, long References, long Nodes) previous;
    private readonly CancellationToken preparationToken;
    private OfdTextReplacementFitStatistics? fitStatistics;
    internal string Phase { get; private set; } = "CaptureSource";

    internal OfdTextReplacementOperation(OfdDocumentPackage source, OfdTextReplacementOptions options, CancellationToken token)
        : this(source, options, token, new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = options.Limits.MaxMaterializedBytes })) { }

    internal static OfdTextReplacementOperation ForFileSession(OfdTextReplacementFileSession session,
        OfdDocumentPackage source, OfdControlledEditContext context, CancellationToken token)
    {
        session.ClaimPreparation(source, context, token);
        return new OfdTextReplacementOperation(source, session.Options, token, context);
    }

    internal static OfdTextReplacementOperation ForOwnedFileSession(OfdTextReplacementFileSession session,
        OfdDocumentPackage source, OfdControlledEditContext context, CancellationToken token)
    {
        session.ClaimOwnedPreparation(source, context, token);
        return new OfdTextReplacementOperation(source, session.Options, token, context);
    }

    private OfdTextReplacementOperation(OfdDocumentPackage source, OfdTextReplacementOptions options, CancellationToken token,
        OfdControlledEditContext context)
    {
        Options = options; Source = source; preparationToken = token;
        Context = context; Context.AssertActive();
        var baseline = Context.Material;
        previous = (baseline.StorageBytes, baseline.CloneCopyBytes, baseline.OutputUpperBoundBytes,
            baseline.ValidationWorkBytes, baseline.Structure.References, baseline.Structure.Nodes);
        try
        {
            // Reserve bounded phase groups, immutable success/failure reports,
            // and up to three internal trace readbacks before allocating them.
            // Reporting budget exhaustion must never mask its original error.
            Context.Material.ChargeCopy(65_536, token);
            phases = new List<OfdTextReplacementPhase>(16); analyses = new List<OfdKeywordLocatorStatistics>(2);
            trace = new OfdTextReplacementTraceCell[32]; Begin("CaptureSource"); trace[currentTrace].SourceCaptures++;
            SourceGuard = Context.CaptureSource(source, token); Mark("CapturedSource");
        }
        catch (Exception error) { Fail(error); Dispose(); throw; }
    }
    internal void Begin(string phase)
    {
        Context.AssertActive(); FlushTrace(); Phase = phase;
        for (var i = 0; i < traceCount; i++)
            if (trace[i].Name == phase) { currentTrace = i; trace[i].Entries++; return; }
        if (traceCount == trace.Length) throw new InvalidOperationException("Replacement trace phase capacity exceeded.");
        currentTrace = traceCount++; trace[currentTrace].Name = phase; trace[currentTrace].Entries = 1;
    }
    internal void BeforeCallback(string phase)
    {
        Begin(phase);
        if (phase is "OpenProvider" or "MeasureSource" or "MeasureReplacement" or "ProviderStatistics" or "DisposeProvider")
            trace[currentTrace].CallbackBoundaryAttempts++;
    }
    internal void Validate(CancellationToken token)
    { trace[currentTrace].SourceGuardAttempts++; SourceGuard!.AssertUnchanged(Source!, Context, token); }
    internal void CloneCandidate(CancellationToken token)
    { trace[currentTrace].OwnedCloneAttempts++; Candidate = OfdOwnedEditCandidate.CloneForControlledEdit(Source!, SourceGuard!, Context, token); }
    internal void CaptureCandidate(CancellationToken token)
    { trace[currentTrace].CandidateCaptureAttempts++; CandidateGuard = Candidate!.CaptureCandidate(token); }
    internal void RecordRetainedDigests(long bytes, long operations)
    { trace[currentTrace].RetainedFontDigestBytes += bytes; trace[currentTrace].RetainedFontDigestOperations += operations; }
    internal void Mark(string phase)
    {
        Begin(phase); var m = Context.Material;
        if (phases.Count == 16) throw new InvalidOperationException("Replacement phase capacity exceeded.");
        phases.Add(new OfdTextReplacementPhase(phase, m.StorageBytes, m.CloneCopyBytes, m.OutputUpperBoundBytes, m.ValidationWorkBytes,
            m.Structure.References, m.Structure.Nodes));
    }
    internal void RecordAnalysis(OfdKeywordSearchResult result)
    {
        if (analyses.Count == 2) throw new InvalidOperationException("Replacement analysis capacity exceeded.");
        analyses.Add(result.Statistics);
    }
    internal void RecordFit(OfdTextFitStatistics fit)
    {
        fitStatistics = new(fit.Visits, fit.Comparisons, fit.OutputPositions);
    }
    internal OfdTextReplacementStatistics Snapshot()
    {
        return new(phases.ToArray(), analyses.ToArray(), fitStatistics);
    }
    internal void Fail(Exception error)
    {
        Context.Fail();
        try
        {
            var m = Context.Material;
            // If reservation itself failed, avoid allocating an unreserved report.
            if (trace.Length != 0)
            {
                error.Data["Ofdrw.TextReplacement.Phase"] = Phase;
                error.Data["Ofdrw.TextReplacement.Counters"] = new OfdTextReplacementPhase(Phase, m.StorageBytes,
                    m.CloneCopyBytes, m.OutputUpperBoundBytes, m.ValidationWorkBytes, m.Structure.References, m.Structure.Nodes);
                error.Data["Ofdrw.TextReplacement.Trace"] = TraceSnapshot();
                if (fitStatistics is not null) error.Data["Ofdrw.TextReplacement.Fit"] = fitStatistics;
            }
        }
        catch { /* Exception.Data can itself be supplied by an untrusted exception subclass. */ }
    }
    internal OfdDocumentPackage TakePackage(CancellationToken token)
    {
        Begin("TransferPackage"); preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested(); Validate(token);
        if (Candidate is not null)
        {
            trace[currentTrace].CandidateGuardAttempts++;
            CandidateGuard!.AssertUnchanged(Candidate.Package, Context, token);
            return Candidate.Package;
        }
        // No-op copying is explicitly a new bounded operation, not a second
        // clone from the original edit operation or an exact-ZIP publication.
        trace[currentTrace].IndependentNoOpCopyOperations++;
        using var copyContext = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = Options.Limits.MaxMaterializedBytes });
        try
        {
            var guard = copyContext.CaptureSource(Source!, token);
            trace[currentTrace].IndependentNoOpCloneAttempts++;
            using var clone = OfdOwnedEditCandidate.CloneForControlledEdit(Source!, guard, copyContext, token);
            Validate(token); clone.AssertSourceUnchanged(token);
            return clone.Package;
        }
        finally
        {
            ref var cell = ref trace[currentTrace]; var m = copyContext.Material;
            cell.IndependentNoOpStorageBytes = m.StorageBytes; cell.IndependentNoOpCopyBytes = m.CloneCopyBytes;
            cell.IndependentNoOpOutputUpperBoundBytes = m.OutputUpperBoundBytes; cell.IndependentNoOpValidationWorkBytes = m.ValidationWorkBytes;
        }
    }
    public void Dispose()
    {
        Candidate?.Dispose(); Candidate = null; CandidateGuard = null;
        SourceGuard = null; Source = null; phases?.Clear(); analyses?.Clear(); fitStatistics = null; Context.Dispose();
    }
    internal OfdTextReplacementTraceCell[] TraceSnapshot()
    {
        if (++traceSnapshots > 3) throw new InvalidOperationException("Replacement trace readback capacity exceeded.");
        FlushTrace(); var copy = new OfdTextReplacementTraceCell[traceCount]; Array.Copy(trace, copy, traceCount); return copy;
    }
    private void FlushTrace()
    {
        if (currentTrace < 0) return;
        var m = Context.Material;
        if (currentTrace >= 0)
        {
            ref var cell = ref trace[currentTrace];
            cell.StorageBytes += m.StorageBytes - previous.Storage; cell.CopyBytes += m.CloneCopyBytes - previous.Copy;
            cell.OutputUpperBoundBytes += m.OutputUpperBoundBytes - previous.Output; cell.ValidationWorkBytes += m.ValidationWorkBytes - previous.Work;
            cell.TraversalReferences += m.Structure.References - previous.References; cell.TraversalNodes += m.Structure.Nodes - previous.Nodes;
        }
        previous = (m.StorageBytes, m.CloneCopyBytes, m.OutputUpperBoundBytes, m.ValidationWorkBytes, m.Structure.References, m.Structure.Nodes);
    }
}

// Aggregated observed deltas for a fixed bounded set of operation phases.
// Shared clone's internal entry/exit hashes are not counted as observed guards.
// CallbackBoundaryAttempts counts preparation BeforeCallback boundaries only;
// exception-unwind session disposal is separately verified by provider ownership tests.
internal struct OfdTextReplacementTraceCell
{
    public string? Name { get; set; }
    public long Entries { get; set; }
    public long CallbackBoundaryAttempts { get; set; }
    public long SourceCaptures { get; set; }
    public long SourceGuardAttempts { get; set; }
    public long CandidateCaptureAttempts { get; set; }
    public long CandidateGuardAttempts { get; set; }
    public long OwnedCloneAttempts { get; set; }
    public long RetainedFontDigestBytes { get; set; }
    public long RetainedFontDigestOperations { get; set; }
    public long IndependentNoOpCopyOperations { get; set; }
    public long IndependentNoOpCloneAttempts { get; set; }
    public long IndependentNoOpStorageBytes { get; set; }
    public long IndependentNoOpCopyBytes { get; set; }
    public long IndependentNoOpOutputUpperBoundBytes { get; set; }
    public long IndependentNoOpValidationWorkBytes { get; set; }
    public long StorageBytes { get; set; }
    public long CopyBytes { get; set; }
    public long OutputUpperBoundBytes { get; set; }
    public long ValidationWorkBytes { get; set; }
    public long TraversalReferences { get; set; }
    public long TraversalNodes { get; set; }
}
