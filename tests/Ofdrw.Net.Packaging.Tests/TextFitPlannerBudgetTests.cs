using System.Collections;
using Ofdrw.Net.Layout.Editing.Internal;
using static Ofdrw.Net.Packaging.Tests.TextFitPlannerTests;

namespace Ofdrw.Net.Packaging.Tests;

public class TextFitPlannerBudgetTests
{
    [Fact]
    public void ExactInputOutputAndTokenBoundariesPassAndOneLessFails()
    {
        var input = Input();
        // ABC (3), expected B (1), replacement -- (2); 3 original+2 replacement metrics.
        // Tokens: batch12 + object33 + source slots12 + request16 =73 UTF16 units.
        var result = Plan(input, new(clusters: 5, inputUtf16: 6, outputClusters: 4, outputUtf16: 4, tokenUtf16: 73));
        Assert.True(result.FitsSuppliedGeometry); Assert.Equal(7, result.Statistics.CapturedPositions);
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(clusters: 4)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(inputUtf16: 5)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(outputClusters: 3)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(outputUtf16: 3)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(tokenUtf16: 72)));
    }

    [Fact]
    public void CountsAccumulateAcrossObjectsRequestsAndNeighbors()
    {
        var input = Input(new[] { Object(), Object(occurrence: "second", owner: "owner2", page: 1) },
            new[] { Request(), Request(occurrence: "second") },
            new[] { new OfdTextFitNeighbor(2, "one", new(0, 0, 1, 1)), new OfdTextFitNeighbor(2, "two", new(2, 2, 3, 3)) });
        Assert.True(Plan(input, new(objects: 2, requests: 2, neighbors: 2, clusters: 10, outputClusters: 8)).FitsSuppliedGeometry);
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(objects: 1)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(requests: 1)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(neighbors: 1)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(clusters: 9)));
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(outputClusters: 7)));
    }

    [Fact]
    public void EveryRawIntervalCoverageAndGeometricPairConsumesOneSharedBudget()
    {
        // Raw slot pairs: 3; page coverage:1; candidates versus A/C:4;
        // second candidate versus first:1 =>9. No request sort for a singleton.
        var result = Plan(limits: new(comparisons: 9)); Assert.True(result.FitsSuppliedGeometry); Assert.Equal(9, result.Statistics.Comparisons);
        Assert.Throws<OfdTextFitLimitException>(() => Plan(limits: new(comparisons: 8)));
    }

    [Fact]
    public void SortingChargesTheSameComparisonBudgetWithoutWrappingException()
    {
        var input = Input(requests: new[] { Request("-", 2, 1, "C"), Request("-", 0, 1, "A"), Request("-", 1, 1, "B") });
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(comparisons: 4)));
        Assert.True(Plan(input).FitsSuppliedGeometry);
    }

    [Fact]
    public void OversizedDeclaredCountRejectsBeforeReadingOrAllocatingElements()
    {
        var source = new DeclaredList<OfdTextFitSourceCluster>();
        Assert.Throws<OfdTextFitLimitException>(() => Object(sources: source)); Assert.Equal(0, source.Reads);
        var objects = new DeclaredList<OfdTextFitObject>();
        Assert.Throws<OfdTextFitLimitException>(() => Input(objects)); Assert.Equal(0, objects.Reads);
        var metrics = new DeclaredList<OfdTextFitMetric>();
        Assert.Throws<OfdTextFitLimitException>(() => Request(metrics: metrics)); Assert.Equal(0, metrics.Reads);
    }

    [Fact]
    public void LimitsAreFrozenPositiveAndBoundedAndCapacityMathRejectsOverflow()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdTextFitLimits(objects: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdTextFitLimits(comparisons: int.MaxValue));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdTextFitLimits(outputUtf16: -1));
        Assert.Throws<OfdTextFitLimitException>(() => OfdTextFitBudget.Add(long.MaxValue, 1, long.MaxValue));
        Assert.Throws<OfdTextFitLimitException>(() => OfdTextFitBudget.Add(0, -1, 5));
        Assert.Equal(long.MaxValue, OfdTextFitBudget.Add(long.MaxValue - 1, 1, long.MaxValue));
        Assert.Throws<OfdTextFitLimitException>(() => Object(occurrence: new string('x', 257)));
        Assert.Throws<OfdTextFitLimitException>(() => Object(text: new string('x', 1000001), sources: Array.Empty<OfdTextFitSourceCluster>()));
    }

    [Fact]
    public void CancellationCheckedBeforePlanAndBeforeAndDuringCapture()
    {
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdTextFitPlanner.Plan(Input(), new(), canceled.Token));
        var declared = new DeclaredList<OfdTextFitObject>();
        Assert.Throws<OperationCanceledException>(() => new OfdTextFitInput("version", "scope", declared,
            Array.Empty<OfdTextFitRequest>(), Array.Empty<OfdTextFitNeighbor>(), canceled.Token)); Assert.Equal(0, declared.CountReads);
        using var during = new CancellationTokenSource();
        var values = new CancelingList<OfdTextFitObject>(Object(), during);
        Assert.Throws<OperationCanceledException>(() => new OfdTextFitInput("version", "scope", values,
            Array.Empty<OfdTextFitRequest>(), Array.Empty<OfdTextFitNeighbor>(), during.Token)); Assert.Equal(1, values.Reads);
    }

    [Fact]
    public void EveryBudgetVisitCompareOutputAndPublicationCheckObservesCancellation()
    {
        using var source = new CancellationTokenSource(); var budget = new OfdTextFitBudget(new(), source.Token);
        budget.Visit(); budget.Compare(); source.Cancel();
        Assert.Throws<OperationCanceledException>(() => budget.Visit());
        Assert.Throws<OperationCanceledException>(() => budget.Compare());
        Assert.Throws<OperationCanceledException>(() => budget.Output(1, 1));
        Assert.Throws<OperationCanceledException>(() => budget.Check());
    }

    [Fact]
    public void NoOpOutputStillChargedAndEmptyBatchNeedsNoOutputCapacity()
    {
        var input = Input(requests: new[] { Request("B", metrics: Metrics("B")) });
        var result = Plan(input, new(outputClusters: 3, outputUtf16: 3)); Assert.True(result.FitsSuppliedGeometry);
        Assert.Equal(3, result.Statistics.OutputPositions);
        Assert.Throws<OfdTextFitLimitException>(() => Plan(input, new(outputClusters: 2)));
        Assert.Empty(Plan(Input(requests: Array.Empty<OfdTextFitRequest>()), new(outputClusters: 1)).Plans);
    }

    [Fact]
    public void DeletionCannotHideOverlappingRequestsFromOutputPreflight()
    {
        Failure(OfdTextFitFailure.OverlappingRequests, Input(requests: new[] { Request("", 0, 3, "ABC"), Request("", 0, 3, "ABC") }));
    }

    [Fact]
    public void ManyDisjointDeletionsValidateObjectProgressionOnce()
    {
        const int count = 5000;
        var text = new string('A', count);
        var sources = Enumerable.Range(0, count).Select(i => new OfdTextFitSourceCluster(Metric('A', i), i * 3, 4, "slot" + i, 0, 1)).ToArray();
        var obj = Object(text, sources, boundary: new(0, 0, count * 3 + 1, 10), clip: new(0, 0, count * 3 + 1, 10));
        var requests = Enumerable.Range(0, count).Select(i => Request("", i, 1, "A")).ToArray();
        var result = Plan(Input(new[] { obj }, requests)); Assert.True(result.FitsSuppliedGeometry);
        Assert.True(Assert.Single(result.Plans).RetainRootWithCanonicalEmptyTextCode);
        // An upper bound linear in supplied clusters+requests; repeated object scans
        // would exceed it by orders of magnitude, independently of pair limits.
        Assert.InRange(result.Statistics.Visits, 1, 13L * count);
        Assert.InRange(result.Statistics.Comparisons, 1, 3L * count);
    }

    [Fact]
    public void SharedFrozenSnapshotIsNotCountedAsTwoFactoryCopies()
    {
        var obj = Object();
        var result = Plan(Input(new[] { obj, obj }));
        Assert.Equal(OfdTextFitFailure.InvalidIdentity, result.Failure);
        // Input captured 3 top-level positions, object captured 3 once, request2.
        Assert.Equal(8, result.Statistics.CapturedPositions);
    }

    private sealed class DeclaredList<T> : IReadOnlyList<T>
    {
        internal int Reads { get; private set; }
        internal int CountReads { get; private set; }
        public int Count { get { CountReads++; return int.MaxValue; } }
        public T this[int index] { get { Reads++; throw new InvalidOperationException("Elements must not be visited."); } }
        public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
    private sealed class CancelingList<T> : IReadOnlyList<T>
    {
        private readonly T value; private readonly CancellationTokenSource source;
        internal CancelingList(T value, CancellationTokenSource source) { this.value = value; this.source = source; }
        internal int Reads { get; private set; }
        public int Count => 2;
        public T this[int index] { get { Reads++; source.Cancel(); return value; } }
        public IEnumerator<T> GetEnumerator() => throw new InvalidOperationException();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
