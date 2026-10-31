using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Core.Processes;

internal static class UnixLaunchOwnershipProbe
{
    internal static async Task<int> RunAsync(string directory)
    {
        if (OperatingSystem.IsWindows()) return 44;
        if (Path.TrimEndingDirectorySeparator(Path.GetTempPath()) != Path.TrimEndingDirectorySeparator(directory)) return 45;
        var marker = Path.Combine(directory, "target-started");
        var type = typeof(OwnedProcessGroup).GetNestedType("UnixGroup", BindingFlags.NonPublic)!;
        var gate = type.GetField("LaunchLock", BindingFlags.Static | BindingFlags.NonPublic)!.GetValue(null)!;
        using var entered = new ManualResetEventSlim();
        var completed = new TaskCompletionSource<ExternalProcessResult>(TaskCreationOptions.RunContinuationsAsynchronously);
        var worker = new Thread(() =>
        {
            try
            {
                var start = new ProcessStartInfo("/bin/sh");
                foreach (var value in new[] { "-c", "printf started > \"$1\"", "supervisor-probe", marker })
                    start.ArgumentList.Add(value);
                entered.Set();
                completed.SetResult(ExternalProcessRunner.RunAsync(start, TimeSpan.FromSeconds(10), default).GetAwaiter().GetResult());
            }
            catch (Exception failure) { completed.SetException(failure); }
        }) { IsBackground = true };
        bool workerAnnouncedStartup, observedWorkerWaiting, markerBefore;
        string[] directoriesBefore;
        Monitor.Enter(gate);
        try
        {
            worker.Start();
            workerAnnouncedStartup = entered.Wait(TimeSpan.FromSeconds(5));
            // A finite concurrency observation: the worker announced startup
            // and is waiting. ThreadState does not identify the awaited lock;
            // source review separately verifies the extraction boundary.
            observedWorkerWaiting = workerAnnouncedStartup && SpinWait.SpinUntil(
                () => (worker.ThreadState & System.Threading.ThreadState.WaitSleepJoin) != 0,
                TimeSpan.FromSeconds(5));
            directoriesBefore = Directory.GetDirectories(directory, "ofdrw-supervisor-*");
            markerBefore = File.Exists(marker);
        }
        finally { Monitor.Exit(gate); }
        var result = await completed.Task.WaitAsync(TimeSpan.FromSeconds(15));
        var workerExited = worker.Join(TimeSpan.FromSeconds(2));
        var directoriesAfter = Directory.GetDirectories(directory, "ofdrw-supervisor-*");
        var markerAfter = File.Exists(marker) && File.ReadAllText(marker) == "started";
        var core = typeof(OwnedProcessGroup).Assembly.Location;
        Console.WriteLine(JsonSerializer.Serialize(new
        {
            workerAnnouncedStartup, observedWorkerWaiting, directoriesBefore, markerBefore, workerExited,
            directoriesAfter, markerAfter, result.ExitCode, result.Output, result.Error,
            corePath = core, coreSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(core))).ToLowerInvariant()
        }));
        return workerAnnouncedStartup && observedWorkerWaiting && directoriesBefore.Length == 0 && !markerBefore &&
               workerExited && directoriesAfter.Length == 0 && markerAfter && result.ExitCode == 0 ? 0 : 43;
    }
}
