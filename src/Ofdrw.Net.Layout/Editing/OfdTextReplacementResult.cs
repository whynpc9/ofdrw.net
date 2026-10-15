using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing.Internal;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>A prepared replacement with one mutually exclusive copy or save consumption.</summary>
/// <remarks>Dispose an abandoned result. Inputs must remain stable through consumption.
/// The returned mutable package is independent. Its later public Writer call has a separate budget.</remarks>
public sealed class OfdTextReplacementResult : IDisposable
{
    private readonly object sync = new();
    private OfdTextReplacementOperation? operation;
    private int state; // prepared, consuming, consumed/disposed
    internal OfdTextReplacementResult(OfdTextReplacementOperation operation, OfdTextReplacementObservation[] observations)
    {
        this.operation = operation; Replacements = Array.AsReadOnly(observations);
        RequestedCount = observations.Length; ChangedCount = observations.Count(value => value.Changed);
        Diagnostics = Array.AsReadOnly(Array.Empty<OfdTextReplacementDiagnostic>()); Statistics = operation.Snapshot();
    }
    /// <summary>Whether at least one requested logical string differs from its original text.</summary>
    public bool Changed => ChangedCount != 0;
    /// <summary>Number of authenticated requests in the prepared batch.</summary>
    public int RequestedCount { get; }
    /// <summary>Number of requests whose logical text changes.</summary>
    public int ChangedCount { get; }
    /// <summary>Copied immutable observations; these retain no source package or live hit.</summary>
    public IReadOnlyList<OfdTextReplacementObservation> Replacements { get; }
    /// <summary>Bounded notices produced during replacement preparation.</summary>
    public IReadOnlyList<OfdTextReplacementDiagnostic> Diagnostics { get; }
    /// <summary>Preparation counters; subsequent independent saves have their own budgets.</summary>
    public OfdTextReplacementStatistics Statistics { get; }
    /// <summary>Consumes this result once and transfers its independent package without a second changed-package clone.</summary>
    public OfdDocumentPackage CreatePackageCopy(CancellationToken cancellationToken = default)
    {
        var owned = Claim();
        try { return owned.TakePackage(cancellationToken); }
        catch (Exception error) { owned.Fail(error); throw; }
        finally { owned.Dispose(); CompleteConsumption(); }
    }
    internal OfdTextReplacementOperation Claim()
    {
        lock (sync)
        {
            if (state != 0) throw new InvalidOperationException("Replacement result has already been consumed or disposed.");
            state = 1; var owned = operation!; operation = null; return owned;
        }
    }
    internal void CompleteConsumption()
    {
        lock (sync)
        {
            if (state != 1) throw new InvalidOperationException("Replacement result is not being consumed.");
            state = 2;
        }
    }
    /// <summary>Releases an abandoned prepared operation; disposal during active consumption is rejected.</summary>
    public void Dispose()
    {
        lock (sync)
        {
            if (state == 1) throw new InvalidOperationException("Cannot dispose a replacement result while consumption is active.");
            if (state == 2) return;
            state = 2; operation?.Dispose(); operation = null;
        }
    }
}
