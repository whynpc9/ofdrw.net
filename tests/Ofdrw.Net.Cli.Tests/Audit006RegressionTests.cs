using System.Reflection;
using System.Runtime.ExceptionServices;

namespace Ofdrw.Net.Cli.Tests;

public sealed class Audit006RegressionTests
{
    [Fact]
    public void Audit006_SafeOverLimitBaseline()
    {
        Assert.Throws<ArgumentException>(() => ParsePages("1-10001"));
    }

    [Theory]
    [InlineData("1", new[] { 0 })]
    [InlineData("3,1-2,3", new[] { 2, 0, 1, 2 })]
    [InlineData(" , 2 - 3 ,,1, ", new[] { 1, 2, 0 })]
    public void Audit006_BoundaryAndOrder(string expression, int[] expected)
    {
        Assert.Equal(expected, ParsePages(expression));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("   ")]
    public void Audit006_MissingPagesRemainUnresolved(string? expression)
    {
        Assert.Null(ParsePages(expression));
    }

    [Fact]
    public void Audit006_ExactCapacity()
    {
        var pages = ParsePages("1-10000")!;
        Assert.Equal(10_000, pages.Count);
        Assert.Equal(0, pages[0]);
        Assert.Equal(9_999, pages[^1]);
    }

    [Fact]
    public void Audit006_ExactCapacityIncludesDuplicatePositions()
    {
        var pages = ParsePages("1-5000,1-5000")!;
        Assert.Equal(10_000, pages.Count);
        Assert.Equal(4_999, pages[4_999]);
        Assert.Equal(0, pages[5_000]);
    }

    [Theory]
    [InlineData("1-5000,1-5000,1")]
    [InlineData("1-9999,1,2")]
    public void Audit006_ManyRanges(string expression)
    {
        Assert.Throws<ArgumentException>(() => ParsePages(expression));
    }

    [Fact]
    public void Audit006_ManySinglePages()
    {
        Assert.Throws<ArgumentException>(() => ParsePages(string.Join(',', Enumerable.Repeat("1", 10_001))));
    }

    [Fact]
    public void Audit006_ExpressionLengthIsBoundedEvenWithoutSelectedPages()
    {
        Assert.Throws<ArgumentException>(() => ParsePages(new string(',', 320_001)));
        Assert.Throws<ArgumentException>(() => ParsePages(new string(' ', 320_001)));
    }

    [Theory]
    [InlineData("2-1")]
    [InlineData("0")]
    [InlineData("-1")]
    [InlineData("2147483648")]
    [InlineData("1-2-3")]
    public void Audit006_InvalidBoundaries(string expression)
    {
        Assert.Throws<ArgumentException>(() => ParsePages(expression));
    }

    [Theory]
    [InlineData("1,2")]
    [InlineData("1,1")]
    [InlineData("1-2")]
    public void Audit006_SvgCapacity(string expression)
    {
        Assert.Throws<ArgumentException>(() => Invoke("ParsePages", [typeof(string), typeof(int)], [expression, 1]));
    }

    [Fact]
    public void Audit006_SvgSinglePageRetainsOneBasedInput()
    {
        Assert.Equal(new[] { 4 }, (IReadOnlyList<int>)Invoke("ParsePages", [typeof(string), typeof(int)], ["5", 1])!);
    }

    [Fact]
    [Trait("Stage", "GuardedOnly")]
    public void Audit006_LargeRange()
    {
        // Call the same pure preflight used before any materialization.
        Assert.Throws<ArgumentException>(() => Invoke("GetPageRangeCount",
            [typeof(int), typeof(int), typeof(int)], [0, int.MaxValue, 10_000]));
        Assert.Throws<ArgumentException>(() => ParsePages("1-2147483647"));
    }

    [Fact]
    [Trait("Stage", "GuardedOnly")]
    public void Audit006_NearIntMaxRange()
    {
        Assert.Equal(new[] { int.MaxValue - 2, int.MaxValue - 1 }, ParsePages("2147483646-2147483647"));
    }

    private static IReadOnlyList<int>? ParsePages(string? expression)
        => (IReadOnlyList<int>?)Invoke("ParsePages", [typeof(string)], [expression]);

    private static object? Invoke(string method, Type[] types, object?[] arguments)
    {
        var parser = typeof(global::Cli).GetMethod(method, BindingFlags.NonPublic | BindingFlags.Static,
            binder: null, types: types, modifiers: null)!;
        try
        {
            return parser.Invoke(null, arguments);
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            ExceptionDispatchInfo.Capture(ex.InnerException).Throw();
            throw;
        }
    }
}
