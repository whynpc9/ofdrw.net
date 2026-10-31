using System.Diagnostics;
using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Core.Tests;

public sealed class SupervisorExecutableBusyRetryTests
{
    [Fact]
    public void AcceptedSpawnIsReturnedBeforeLaterCancellationOrBudgetChecks()
    {
        using var cancellation = new CancellationTokenSource();
        var accepted = false;
        var calls = 0;
        var error = SupervisorExecutableBusyRetry.Start(() =>
        {
            calls++; accepted = true; cancellation.Cancel(); return 0;
        }, () => { if (accepted) throw new TimeoutException(); },
            () => throw new InvalidOperationException("Acceptance must not request a retry delay."), cancellation.Token);
        Assert.Equal(0, error.Error);
        Assert.Equal(1, calls);
        Assert.Equal(1, error.Attempts);
        Assert.Equal(0, error.FirstError);
        Assert.Equal("Accepted", error.Termination);
    }

    [Fact]
    public void OtherAdmissionErrorsRemainAnImmediateFailure()
    {
        var calls = 0;
        var error = SupervisorExecutableBusyRetry.Start(() => { calls++; return 13; },
            () => { }, () => TimeSpan.FromSeconds(10), default);
        Assert.Equal(13, error.Error);
        Assert.Equal(1, calls);
    }

    [Fact]
    public void BusyAdmissionCanRecoverWithoutRepeatingTheAcceptedSpawn()
    {
        var calls = 0;
        var error = SupervisorExecutableBusyRetry.Start(() => ++calls < 3 ? 26 : 0,
            () => { }, () => TimeSpan.FromSeconds(10), default);
        Assert.Equal(0, error.Error);
        Assert.Equal(3, calls);
        Assert.Equal(3, error.Attempts);
        Assert.Equal(26, error.FirstError);
        Assert.Equal("Accepted", error.Termination);
    }

    [Fact]
    public void PersistentBusyAdmissionStillFailsAfterFiniteAttempts()
    {
        var calls = 0;
        var error = SupervisorExecutableBusyRetry.Start(() => { calls++; return 26; },
            () => { }, () => TimeSpan.FromSeconds(10), default);
        Assert.Equal(26, error.Error);
        Assert.Equal(7, calls);
        Assert.Equal(7, error.Attempts);
        Assert.Equal(26, error.FirstError);
        Assert.Equal("BusyAttemptLimit", error.Termination);
    }

    [Fact]
    public void BusyThenOtherAdmissionErrorStopsImmediately()
    {
        var calls = 0;
        var result = SupervisorExecutableBusyRetry.Start(() => ++calls == 1 ? 26 : 13,
            () => { }, () => TimeSpan.FromSeconds(10), default);
        Assert.Equal(13, result.Error);
        Assert.Equal(2, calls);
        Assert.Equal(26, result.FirstError);
        Assert.Equal("OtherError", result.Termination);
    }

    [Fact]
    public void AlreadyCanceledStartupDoesNotAttemptASpawn()
    {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        var calls = 0;
        Assert.ThrowsAny<OperationCanceledException>(() => SupervisorExecutableBusyRetry.Start(
            () => { calls++; return 0; }, () => { }, () => TimeSpan.FromSeconds(10), cancellation.Token));
        Assert.Equal(0, calls);
    }

    [Fact]
    public void CancellationAfterBusyAdmissionPreventsAnotherAttempt()
    {
        using var cancellation = new CancellationTokenSource();
        var calls = 0;
        var failure = Assert.ThrowsAny<OperationCanceledException>(() => SupervisorExecutableBusyRetry.Start(
            () => { calls++; cancellation.Cancel(); return 26; }, () => { },
            () => TimeSpan.FromSeconds(10), cancellation.Token));
        Assert.Equal(1, calls);
        Assert.Equal(1, failure.Data["OwnedProcessSupervisor.spawnAttempts"]);
        Assert.Equal(26, failure.Data["OwnedProcessSupervisor.firstSpawnError"]);
        Assert.Equal("Cancellation", failure.Data["OwnedProcessSupervisor.spawnTermination"]);
    }

    [Fact]
    public void BudgetExpiredDuringBusyAdmissionPreventsAnotherAttempt()
    {
        var calls = 0;
        var failure = Assert.Throws<TimeoutException>(() => SupervisorExecutableBusyRetry.Start(
            () => { calls++; return 26; }, () => { if (calls > 0) throw new TimeoutException(); },
            () => TimeSpan.Zero, default));
        Assert.Equal(1, calls);
        Assert.Equal(26, failure.Data["OwnedProcessSupervisor.firstSpawnError"]);
        Assert.Equal("Deadline", failure.Data["OwnedProcessSupervisor.spawnTermination"]);
    }

    [Fact]
    public void BusyBackoffConsumesTheOriginalDeadline()
    {
        var elapsed = new Stopwatch(); var calls = 0; var budget = TimeSpan.FromMilliseconds(5);
        Assert.Throws<TimeoutException>(() => SupervisorExecutableBusyRetry.Start(
            () => { calls++; elapsed.Start(); return 26; },
            () => { if (elapsed.Elapsed >= budget) throw new TimeoutException(); },
            () => budget - elapsed.Elapsed, default));
        Assert.InRange(calls, 1, 6);
        Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(1));
    }
}
