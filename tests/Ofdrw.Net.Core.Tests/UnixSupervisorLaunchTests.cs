using System.Diagnostics;
using System.Reflection;
using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Core.Tests;

public sealed class UnixSupervisorLaunchTests
{
    [UnixSupervisorFact]
    public async Task ObservedConcurrentStartupDoesNotPublishSupervisorWhileLaunchOwnerIsHeld()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-launch-window-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet")
        { UseShellExecute = false, RedirectStandardOutput = true, RedirectStandardError = true };
        var runtimeConfig = Path.Combine(AppContext.BaseDirectory, "Ofdrw.Net.Core.Tests.runtimeconfig.json");
        foreach (var value in new[] { "exec", "--runtimeconfig", runtimeConfig,
                     typeof(ProcessFixtureMarker).Assembly.Location, "owned-launch-window-probe", directory })
            start.ArgumentList.Add(value);
        start.Environment["TMPDIR"] = directory;
        using var process = Process.Start(start)!;
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        try
        {
            await process.WaitForExitAsync().WaitAsync(TimeSpan.FromSeconds(20));
            Assert.True(process.ExitCode == 0, (await output) + "\n" + (await error));
        }
        finally
        {
            if (!process.HasExited) { process.Kill(entireProcessTree: true); process.WaitForExit(3000); }
            Directory.Delete(directory, true);
        }
    }

    [UnixSupervisorFact]
    public async Task CancellationWhileWaitingForLaunchOwnershipIsBounded()
    {
        using var held = new HeldLaunchLock();
        using var cancellation = new CancellationTokenSource(TimeSpan.FromMilliseconds(150));
        var marker = Path.Combine(Path.GetTempPath(), "ofdrw-unstarted-target-" + Guid.NewGuid().ToString("N"));
        var elapsed = Stopwatch.StartNew();
        var running = Task.Run(() => ExternalProcessRunner.RunAsync(MarkerCommand(marker),
            TimeSpan.FromSeconds(10), cancellation.Token));
        try
        {
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => running.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(2));
            Assert.False(File.Exists(marker));
        }
        finally
        {
            held.Dispose(); await ObserveCompletion(running);
            var targetStarted = File.Exists(marker); File.Delete(marker);
            Assert.False(targetStarted);
        }
        var normal = await ExternalProcessRunner.RunAsync(new ProcessStartInfo("/usr/bin/true"), TimeSpan.FromSeconds(10), default);
        Assert.Equal(0, normal.ExitCode);
    }

    [UnixSupervisorFact]
    public async Task WaitingForLaunchOwnershipConsumesTheOriginalStartupDeadline()
    {
        using var held = new HeldLaunchLock();
        var marker = Path.Combine(Path.GetTempPath(), "ofdrw-unstarted-target-" + Guid.NewGuid().ToString("N"));
        var elapsed = Stopwatch.StartNew();
        var running = Task.Run(() => ExternalProcessRunner.RunAsync(MarkerCommand(marker),
            TimeSpan.FromMilliseconds(200), default));
        try
        {
            var error = await Assert.ThrowsAsync<TimeoutException>(() => running.WaitAsync(TimeSpan.FromSeconds(2)));
            Assert.Contains("during process startup", error.Message);
            Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(2));
            Assert.False(File.Exists(marker));
        }
        finally
        {
            held.Dispose(); await ObserveCompletion(running);
            var targetStarted = File.Exists(marker); File.Delete(marker);
            Assert.False(targetStarted);
        }
        var normal = await ExternalProcessRunner.RunAsync(new ProcessStartInfo("/usr/bin/true"), TimeSpan.FromSeconds(10), default);
        Assert.Equal(0, normal.ExitCode);
    }

    private static ProcessStartInfo MarkerCommand(string marker)
    {
        var start = new ProcessStartInfo("/bin/sh");
        foreach (var value in new[] { "-c", "printf started > \"$1\"", "supervisor-probe", marker })
            start.ArgumentList.Add(value);
        return start;
    }

    private static async Task ObserveCompletion(Task running)
    {
        try { await running.WaitAsync(TimeSpan.FromSeconds(5)); }
        catch (OperationCanceledException) { }
        catch (TimeoutException) { }
        Assert.True(running.IsCompleted, "Startup remained pending after launch ownership was released.");
    }

    // Hold real launch ownership from a different thread, without starting a
    // process or delaying extraction through a production test hook.
    private sealed class HeldLaunchLock : IDisposable
    {
        private readonly ManualResetEventSlim _ready = new();
        private readonly ManualResetEventSlim _release = new();
        private readonly Thread _owner;
        private bool _disposed;
        internal HeldLaunchLock()
        {
            var type = typeof(OwnedProcessGroup).GetNestedType("UnixGroup", BindingFlags.NonPublic)!;
            var gate = type.GetField("LaunchLock", BindingFlags.Static | BindingFlags.NonPublic)!.GetValue(null)!;
            _owner = new Thread(() =>
            {
                lock (gate) { _ready.Set(); _release.Wait(); }
            }) { IsBackground = true };
            _owner.Start();
            if (!_ready.Wait(TimeSpan.FromSeconds(5)))
            {
                _release.Set();
                throw new TimeoutException("The test launch owner did not acquire its gate.");
            }
        }
        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            _release.Set();
            Assert.True(_owner.Join(TimeSpan.FromSeconds(5)));
            _ready.Dispose(); _release.Dispose();
        }
    }

    private sealed class UnixSupervisorFactAttribute : FactAttribute
    {
        public UnixSupervisorFactAttribute()
        {
            if (OperatingSystem.IsWindows()) Skip = "Unix native supervisor; Windows retains the Job backend.";
        }
    }
}
