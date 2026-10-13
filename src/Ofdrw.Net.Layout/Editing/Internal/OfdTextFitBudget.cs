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
    private readonly Action<long>? chargeWork;
    internal OfdTextFitBudget(OfdTextFitLimits limits, CancellationToken token, Action<long>? chargeWork = null)
    { this.limits = limits; this.token = token; this.chargeWork = chargeWork; Check(); }
    internal long CapturedPositions { get; set; }
    internal void Check() => token.ThrowIfCancellationRequested();
    // Logical byte-equivalent work policy, not measured hash/input bytes:
    // visits=16, pair comparisons=32, output position=16 and UTF16 unit=2.
    // The enclosing hook prepays before these actions; null preserves pure use.
    internal void Visit() { Check(); chargeWork?.Invoke(16); visits = checked(visits + 1); }
    internal void Compare() { Check(); if (comparisons >= limits.Comparisons) throw new OfdTextFitLimitException(); chargeWork?.Invoke(32); comparisons++; }
    internal void Output(int count, int utf16)
    {
        Check(); chargeWork?.Invoke(checked(16L * count + 2L * utf16)); outputPositions = Add(outputPositions, count, limits.OutputClusters);
        outputUtf16 = Add(outputUtf16, utf16, limits.OutputUtf16);
    }
    internal OfdTextFitStatistics Snapshot() => new OfdTextFitStatistics(CapturedPositions, visits, outputPositions, outputUtf16, comparisons);
    internal static long Add(long current, long value, long maximum)
    { if (value < 0 || current < 0 || current > maximum || value > maximum - current) throw new OfdTextFitLimitException(); return current + value; }
}
