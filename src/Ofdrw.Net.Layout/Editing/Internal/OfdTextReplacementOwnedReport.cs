namespace Ofdrw.Net.Layout.Editing.Internal;

// Copied values only: no hit, package, provider, stream, candidate or Writer.
internal sealed class OfdTextReplacementOwnedReport
{
    internal OfdTextReplacementOwnedReport(OfdTextReplacementObservation[] observations, OfdTextReplacementStatistics statistics)
    {
        Replacements = Array.AsReadOnly(observations); Statistics = statistics;
        RequestedCount = observations.Length; ChangedCount = observations.Count(value => value.Changed);
    }
    internal bool Changed => ChangedCount != 0;
    internal int RequestedCount { get; }
    internal int ChangedCount { get; }
    internal IReadOnlyList<OfdTextReplacementObservation> Replacements { get; }
    internal OfdTextReplacementStatistics Statistics { get; }
}
