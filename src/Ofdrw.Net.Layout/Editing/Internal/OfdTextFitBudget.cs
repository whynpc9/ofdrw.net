using System.Threading;

namespace Ofdrw.Net.Layout.Editing.Internal;

internal sealed class OfdTextFitLimitException : InvalidOperationException
{
    internal OfdTextFitLimitException() : base("The supplied-geometry fit budget was exceeded.") { }
}

internal sealed class OfdTextFitLimits
{
    internal const int HardObjects = 512, HardClusters = 100000, HardRequests = 10000,
        HardNeighbors = 100000, HardOutputClusters = 200000, HardUtf16 = 1000000, HardComparisons = 2000000;
    internal OfdTextFitLimits(int objects = HardObjects, int clusters = HardClusters, int requests = HardRequests,
        int neighbors = HardNeighbors, int outputClusters = HardOutputClusters, int inputUtf16 = HardUtf16,
        int outputUtf16 = HardUtf16, int tokenUtf16 = HardUtf16, int comparisons = HardComparisons)
    {
        Objects = Positive(objects, HardObjects); Clusters = Positive(clusters, HardClusters);
        Requests = Positive(requests, HardRequests); Neighbors = Positive(neighbors, HardNeighbors);
        OutputClusters = Positive(outputClusters, HardOutputClusters); InputUtf16 = Positive(inputUtf16, HardUtf16);
        OutputUtf16 = Positive(outputUtf16, HardUtf16); TokenUtf16 = Positive(tokenUtf16, HardUtf16);
        Comparisons = Positive(comparisons, HardComparisons);
    }
    internal int Objects { get; }
    internal int Clusters { get; }
    internal int Requests { get; }
    internal int Neighbors { get; }
    internal int OutputClusters { get; }
    internal int InputUtf16 { get; }
    internal int OutputUtf16 { get; }
    internal int TokenUtf16 { get; }
    internal int Comparisons { get; }
    private static int Positive(int value, int maximum)
    { if (value <= 0 || value > maximum) throw new ArgumentOutOfRangeException(nameof(value)); return value; }
}

internal sealed class OfdTextFitBudget
{
    private readonly CancellationToken token;
    private readonly OfdTextFitLimits limits;
    private long visits, comparisons, outputPositions, outputUtf16;
    internal OfdTextFitBudget(OfdTextFitLimits limits, CancellationToken token) { this.limits = limits; this.token = token; Check(); }
    internal long CapturedPositions { get; set; }
    internal void Check() => token.ThrowIfCancellationRequested();
    internal void Visit() { Check(); visits = checked(visits + 1); }
    internal void Compare() { Check(); if (comparisons >= limits.Comparisons) throw new OfdTextFitLimitException(); comparisons++; }
    internal void Output(int count, int utf16)
    {
        Check(); outputPositions = Add(outputPositions, count, limits.OutputClusters);
        outputUtf16 = Add(outputUtf16, utf16, limits.OutputUtf16);
    }
    internal OfdTextFitStatistics Snapshot() => new OfdTextFitStatistics(CapturedPositions, visits, outputPositions, outputUtf16, comparisons);
    internal static long Add(long current, long value, long maximum)
    { if (value < 0 || current < 0 || current > maximum || value > maximum - current) throw new OfdTextFitLimitException(); return current + value; }
}
