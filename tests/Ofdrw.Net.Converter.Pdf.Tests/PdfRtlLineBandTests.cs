using System.Collections;
using System.Diagnostics;
using Ofdrw.Net.Converter.Pdf.Internal;
using Xunit.Abstractions;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class PdfRtlLineBandTests(ITestOutputHelper output)
{
    [Fact]
    public void MatchesTheExhaustiveScan()
    {
        var random = new Random(20261113);
        var conflictCount = 0;
        var clearCount = 0;
        for (var trial = 0; trial < 2_000; trial++)
        {
            var count = random.Next(1, 41);
            var y = new double[count];
            var em = new double[count];
            var position = 800d;
            for (var index = 0; index < count; index++)
            {
                position -= 30 * random.NextDouble();
                y[index] = position;
                em[index] = 0.5 + 49.5 * random.NextDouble();
            }

            var conflicts = AssertMatchesReference(y, em);
            conflictCount += conflicts;
            clearCount += count - conflicts;
        }

        Assert.True(conflictCount >= 1_000, $"Only {conflictCount} conflicts.");
        Assert.True(clearCount >= 1_000, $"Only {clearCount} clear clusters.");
    }

    [Fact]
    public void HandCases()
    {
        Assert.Equal(0, AssertMatchesReference([100], [10]));
        Assert.Equal(0, AssertMatchesReference([100, 0], [10, 10]));
        Assert.Equal(2, AssertMatchesReference([100, 95], [10, 10]));
        Assert.Equal(3, AssertMatchesReference([300, 200, 100], [10, 10, 400]));
        Assert.Equal(3, AssertMatchesReference([300, 200, 100], [400, 10, 10]));
        Assert.Equal(0, AssertMatchesReference([300, 200, 100], [10, 10, 10]));
        Assert.Equal(2, AssertMatchesReference([+0d, -0d], [double.Epsilon, double.Epsilon]));
    }

    [Theory]
    [InlineData(6, 0, 10, false)]
    [InlineData(21.917802363184098, 18.135260228528736, 6.304236891092271, true)]
    [InlineData(-14.99137861956232, -44.39804455689149, 49.01110989554862, false)]
    [InlineData(-double.Epsilon, -6, 10, false)]
    [InlineData(-1e-15, -6, 10, true)]
    public void StrictFloatingPointBoundaries(double above, double below, double wideEm, bool expected)
    {
        // A smaller cluster between the two forces the running reach to decide the
        // conflict for the outer, narrow cluster. Mirror it to exercise both passes.
        var middle = (above + below) / 2;
        var fromBelow = new PdfRtlVisualLineImport.LineBands([above, middle, below], [1, 1, wideEm], default);
        var fromAbove = new PdfRtlVisualLineImport.LineBands([-below, -middle, -above], [wideEm, 1, 1], default);
        Assert.Equal(expected, fromBelow.Conflicts(0));
        Assert.Equal(expected, fromAbove.Conflicts(2));
    }

    [Fact]
    public void ExceptionalCoordinatesAndSizesFollowTheOldComparison()
    {
        Assert.Equal(2, AssertMatchesReference(
            [double.PositiveInfinity, double.PositiveInfinity, 10, 8, 6, double.NegativeInfinity,
                double.NegativeInfinity, double.NaN, double.NaN],
            [double.PositiveInfinity, 10, 10, double.NaN, 1, 10, double.PositiveInfinity, 10, 10]));
        Assert.Equal(0, AssertMatchesReference([double.MaxValue, -double.MaxValue], [1, double.PositiveInfinity]));
        Assert.Equal(2, AssertMatchesReference([0, -double.MaxValue], [1, double.PositiveInfinity]));
        Assert.Equal(2, AssertMatchesReference([double.MaxValue, Math.BitDecrement(double.MaxValue)], [1, double.MaxValue]));
        Assert.Equal(0, AssertMatchesReference([double.Epsilon, 0], [double.Epsilon, double.Epsilon]));
    }

    [Fact]
    public void NearThresholdsAcrossCoordinateScalesMatchTheExhaustiveScan()
    {
        var random = new Random(33);
        for (var trial = 0; trial < 2_000; trial++)
        {
            var scale = Math.ScaleB(1, random.Next(-1074, 1021));
            var center = (random.NextDouble() - 0.5) * scale;
            var size = Math.Max(double.Epsilon, random.NextDouble() * scale);
            var reach = center + 0.6 * size;
            var positions = new[] { Math.BitIncrement(reach), reach, Math.BitDecrement(reach), center };
            Array.Sort(positions);
            Array.Reverse(positions);
            var em = positions.Select(y => y == center ? size : double.Epsilon).ToArray();
            AssertMatchesReference(positions, em);
            AssertMatchesReference(positions.Reverse().Select(y => -y).ToArray(), em.Reverse().ToArray());
        }
    }

    [Theory]
    [InlineData(1)]
    [InlineData(5)]
    public void CancellationIsObservedDuringEitherScan(int cancelAt)
    {
        using var source = new CancellationTokenSource();
        var sizes = new CancellingSizes(source, cancelAt);
        var error = Assert.Throws<OperationCanceledException>(() =>
            new PdfRtlVisualLineImport.LineBands([6, 4, 2, 0], sizes, source.Token));
        Assert.Equal(source.Token, error.CancellationToken);
        Assert.Equal(cancelAt, sizes.Reads);
    }

    [Fact]
    public void ScalesLinearly()
    {
        const int count = 200_000;
        var y = new double[count];
        var em = new double[count];
        for (var index = 0; index < count; index++)
        {
            y[index] = 1_000_000 - index * 2;
            em[index] = 1;
        }

        var stopwatch = Stopwatch.StartNew();
        var bands = new PdfRtlVisualLineImport.LineBands(y, em, default);
        for (var index = 0; index < count; index++)
            Assert.False(bands.Conflicts(index));
        stopwatch.Stop();
        output.WriteLine($"Indexed and queried {count:N0} clusters in {stopwatch.Elapsed.TotalMilliseconds:F2} ms.");
        Assert.True(stopwatch.Elapsed < TimeSpan.FromSeconds(5), $"Elapsed: {stopwatch.Elapsed}.");
    }

    private static int AssertMatchesReference(double[] y, double[] em)
    {
        var bands = new PdfRtlVisualLineImport.LineBands(y, em, default);
        var conflictCount = 0;
        for (var index = 0; index < y.Length; index++)
        {
            var expected = Reference(y, em, index);
            if (expected != bands.Conflicts(index))
                Assert.Fail($"Index {index}, expected {expected}. Y=[{string.Join(", ", y)}]; Em=[{string.Join(", ", em)}].");
            if (expected) conflictCount++;
        }
        return conflictCount;
    }

    private static bool Reference(IReadOnlyList<double> y, IReadOnlyList<double> em, int index)
    {
        for (var other = 0; other < y.Count; other++)
        {
            if (other == index) continue;
            if (Math.Abs(y[other] - y[index]) < 0.6 * Math.Max(em[other], em[index])) return true;
        }
        return false;
    }

    private sealed class CancellingSizes(CancellationTokenSource source, int cancelAt) : IReadOnlyList<double>
    {
        public int Reads { get; private set; }
        public int Count => 4;
        public double this[int index]
        {
            get
            {
                if (++Reads == cancelAt) source.Cancel();
                return 1;
            }
        }

        public IEnumerator<double> GetEnumerator() => throw new NotSupportedException();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
