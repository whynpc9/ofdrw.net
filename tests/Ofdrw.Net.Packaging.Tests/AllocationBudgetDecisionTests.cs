namespace Ofdrw.Net.Packaging.Tests;

public sealed class AllocationBudgetDecisionTests
{
    [Theory]
    [InlineData(1000, 999, false, false)]
    [InlineData(1000, 999, true, false)]
    [InlineData(-10, 38, false, false)]
    [InlineData(-10, 38, true, false)]
    [InlineData(1000, 1000, false, true)]
    [InlineData(1000, 1048, false, true)]
    [InlineData(1000, 5095, false, true)]
    [InlineData(1000, 5096, false, false)]
    public void BudgetRejectsInvalidCountersAndPreservesBoundary(long before, long after, bool warmup, bool accepted)
    {
        var item = new AllocationCase
        {
            Operation = "target", ValidNoGcInterval = true, SamplesRecorded = 9,
            Samples = Enumerable.Range(0, 9).Select(index =>
                new AllocationSample(index + 1, index == 0, 1000, 1048, Environment.CurrentManagedThreadId)).ToArray()
        };
        var index = warmup ? 0 : 1;
        item.Samples[index] = item.Samples[index] with { Before = before, After = after };
        Assert.Equal(accepted, AllocationProbe.Accept(item));
        Assert.Equal(before, item.Samples[index].Before);
        Assert.Equal(after, item.Samples[index].After); // Invalid observations are retained, never clamped.
    }
}
