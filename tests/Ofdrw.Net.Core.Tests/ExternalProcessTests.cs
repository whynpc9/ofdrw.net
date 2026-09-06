using System.Diagnostics;
using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Core.Tests;

public sealed class ExternalProcessTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-process-tests-" + Guid.NewGuid().ToString("N"));

    public ExternalProcessTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task Runner_ShouldDrainBothPipesWithoutUnboundedCapture()
    {
        var result = await ExternalProcessRunner.RunAsync(CreateStart("flood"), TimeSpan.FromSeconds(15), default);
        Assert.Equal(7, result.ExitCode);
        Assert.Equal(65_536, result.Output.Length);
        Assert.Equal(65_536, result.Error.Length);
    }

    [Fact]
    public async Task Runner_ShouldCancelStartedProcessAndDescendant()
    {
        var pidFile = Path.Combine(_directory, "pid");
        using var cancellation = new CancellationTokenSource();
        var running = ExternalProcessRunner.RunAsync(CreateStart("tree", pidFile, RuntimeConfig),
            TimeSpan.FromSeconds(20), cancellation.Token);
        await WaitForFileAsync(pidFile + ".child", running);
        var parent = int.Parse(await File.ReadAllTextAsync(pidFile));
        var child = int.Parse(await File.ReadAllTextAsync(pidFile + ".child"));
        var elapsed = Stopwatch.StartNew();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => running);
        Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(8));
        await AssertProcessExitedAsync(parent);
        await AssertProcessExitedAsync(child);
    }

    [Fact]
    public async Task Runner_ShouldEnforceDeadlineAndReapProcess()
    {
        var pidFile = Path.Combine(_directory, "pid");
        var running = ExternalProcessRunner.RunAsync(CreateStart("wait", pidFile), TimeSpan.FromSeconds(2), default);
        await WaitForFileAsync(pidFile, running);
        var pid = int.Parse(await File.ReadAllTextAsync(pidFile));
        await Assert.ThrowsAsync<TimeoutException>(() => running);
        await AssertProcessExitedAsync(pid);
    }

    private static string RuntimeConfig => Path.Combine(
        Path.GetDirectoryName(typeof(ExternalProcessTests).Assembly.Location)!,
        "Ofdrw.Net.Core.Tests.runtimeconfig.json");

    private static ProcessStartInfo CreateStart(params string[] arguments)
    {
        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet");
        foreach (var argument in new[]
        {
            "exec", "--runtimeconfig", RuntimeConfig,
            typeof(ProcessFixtureMarker).Assembly.Location
        }.Concat(arguments))
        {
            start.ArgumentList.Add(argument);
        }
        return start;
    }

    private static async Task WaitForFileAsync(string path, Task running)
    {
        var elapsed = Stopwatch.StartNew();
        while (!File.Exists(path) || new FileInfo(path).Length == 0)
        {
            if (running.IsCompleted) await running;
            Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(10), "Process fixture did not start.");
            await Task.Delay(10);
        }
    }

    private static async Task AssertProcessExitedAsync(int pid)
    {
        for (var attempt = 0; attempt < 100; attempt++)
        {
            try
            {
                using var process = Process.GetProcessById(pid);
                if (process.HasExited) return;
            }
            catch (ArgumentException) { return; }
            await Task.Delay(20);
        }

        Assert.Fail($"Process {pid} survived cancellation/timeout.");
    }

    public void Dispose() => Directory.Delete(_directory, recursive: true);
}
