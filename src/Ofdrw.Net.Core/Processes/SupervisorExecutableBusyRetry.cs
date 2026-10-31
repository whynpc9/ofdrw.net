using System;
using System.Threading;

namespace Ofdrw.Net.Core.Processes;

internal readonly struct SupervisorSpawnAdmission
{
    internal SupervisorSpawnAdmission(int error, int attempts, int? firstError, string termination)
    { Error = error; Attempts = attempts; FirstError = firstError; Termination = termination; }
    internal int Error { get; }
    internal int Attempts { get; }
    internal int? FirstError { get; }
    internal string Termination { get; }
    internal void AttachTo(Exception exception)
        => SupervisorExecutableBusyRetry.Attach(exception, Attempts, FirstError, Termination);
}

// Only the private Linux supervisor's failed posix_spawn admission uses this
// policy. Acceptance returns before further cancellation/budget checks, so the
// caller first records the accepted PID and can clean it up.
internal static class SupervisorExecutableBusyRetry
{
    internal static SupervisorSpawnAdmission Start(Func<int> spawn, Action checkBudget,
        Func<TimeSpan> remainingBudget, CancellationToken cancellationToken)
    {
        const int executableBusy = 26;
        const int maximumRetries = 6;
        var attempts = 0;
        int? firstError = null;
        try
        {
            for (var retry = 0; ; retry++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                checkBudget();
                attempts++;
                var error = spawn();
                if (firstError is null) firstError = error;
                if (error != executableBusy)
                    return new SupervisorSpawnAdmission(error, attempts, firstError, error == 0 ? "Accepted" : "OtherError");
                cancellationToken.ThrowIfCancellationRequested();
                checkBudget();
                if (retry == maximumRetries)
                    return new SupervisorSpawnAdmission(error, attempts, firstError, "BusyAttemptLimit");
                var delay = (int)Math.Max(1, Math.Min(1 << retry,
                    Math.Ceiling(remainingBudget().TotalMilliseconds)));
                if (cancellationToken.WaitHandle.WaitOne(delay))
                    cancellationToken.ThrowIfCancellationRequested();
            }
        }
        catch (OperationCanceledException failure)
        { Attach(failure, attempts, firstError, "Cancellation"); throw; }
        catch (TimeoutException failure)
        { Attach(failure, attempts, firstError, "Deadline"); throw; }
    }

    internal static void Attach(Exception exception, int attempts, int? firstError, string termination)
    {
        exception.Data["OwnedProcessSupervisor.spawnAttempts"] = attempts;
        exception.Data["OwnedProcessSupervisor.firstSpawnError"] = firstError.HasValue ? (object)firstError.Value : "none";
        exception.Data["OwnedProcessSupervisor.spawnTermination"] = termination;
    }
}
