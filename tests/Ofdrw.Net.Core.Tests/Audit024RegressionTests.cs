using System.Diagnostics;
using Ofdrw.Net.Core.Processes;

namespace Ofdrw.Net.Core.Tests;

public sealed class Audit024RegressionTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-audit024-" + Guid.NewGuid().ToString("N"));

    public Audit024RegressionTests() => Directory.CreateDirectory(_directory);

    [Theory]
    [InlineData(0)]
    [InlineData(7)]
    public async Task Audit024_IsolatedIgnoredSigchldPreservesRealStatus(int exitCode)
    {
        // Signal disposition is changed only in this isolated fixture process.
        // Windows exercises the same fixture through the unchanged Job backend.
        var info = Start("owned-signal-probe", exitCode.ToString(System.Globalization.CultureInfo.InvariantCulture), RuntimeConfig);
        info.RedirectStandardOutput = true;
        info.RedirectStandardError = true;
        using var process = Process.Start(info)!;
        var output = process.StandardOutput.ReadToEndAsync();
        var error = process.StandardError.ReadToEndAsync();
        try
        {
            Assert.True(process.WaitForExit(20000), "Isolated signal probe exceeded its deadline.");
            Assert.True(process.ExitCode == 0, await error);
            Assert.Equal($"actual={exitCode};stdout=normal output;stderr=normal error", (await output).Trim());
        }
        finally
        {
            if (!process.HasExited) { process.Kill(entireProcessTree: true); process.WaitForExit(3000); }
        }
    }

    [Fact]
    public async Task Audit024_ClosedStandardDescriptorsDoNotCollideWithProtocol()
    {
        var report = Path.Combine(_directory, "closed-stdio-report");
        using var process = Process.Start(Start("owned-signal-probe", "7", RuntimeConfig, "closed-stdio", report))!;
        try
        {
            Assert.True(process.WaitForExit(20000), "Closed-descriptor fixture exceeded its deadline.");
            Assert.Equal(0, process.ExitCode);
            Assert.Equal("actual=7;stdout=normal output;stderr=normal error", await File.ReadAllTextAsync(report));
        }
        finally
        {
            if (!process.HasExited) { process.Kill(entireProcessTree: true); process.WaitForExit(3000); }
        }
    }

    [Fact]
    public async Task Audit024_NormalExit()
    {
        var result = await ExternalProcessRunner.RunAsync(Start("echo"), TimeSpan.FromSeconds(10), default);
        Assert.Equal(7, result.ExitCode);
        Assert.Equal("normal output", result.Output.Trim());
        Assert.Equal("normal error", result.Error.Trim());
    }

    [Fact]
    public Task Audit024_ParentExitsFirstTimeout() => ExerciseCleanup("parent-exits-first", false);

    [Fact]
    public Task Audit024_ParentExitsFirstCancellation() => ExerciseCleanup("parent-exits-first", true);

    [Fact]
    public Task Audit024_ExistingTreeControl() => ExerciseCleanup("tree", true);

    [Fact]
    public Task Audit024_BoundedCleanup() => ExerciseCleanup("parent-exits-first", false);

    [Fact]
    public Task Audit024_StableOutputOwnedTree() => ExerciseCleanup("stable-output", false);

    [Fact]
    public async Task Audit024_NoCollateralProcess()
    {
        var unrelatedFile = Path.Combine(_directory, "unrelated");
        using var unrelated = Process.Start(Start("wait", unrelatedFile))!;
        try
        {
            await WaitForFile(unrelatedFile, Task.CompletedTask);
            await ExerciseCleanup("parent-exits-first", true);
            Assert.False(unrelated.HasExited, "An independent same-executable fixture must survive cleanup.");
        }
        finally
        {
            if (!unrelated.HasExited) unrelated.Kill();
            await unrelated.WaitForExitAsync();
        }
    }

    private async Task ExerciseCleanup(string mode, bool cancel)
    {
        var pidFile = Path.Combine(_directory, Guid.NewGuid().ToString("N"));
        using var cancellation = new CancellationTokenSource();
        var elapsed = Stopwatch.StartNew();
        var running = ExternalProcessRunner.RunAsync(Start(mode, pidFile, RuntimeConfig, pidFile + ".output"),
            TimeSpan.FromSeconds(cancel || mode == "stable-output" ? 15 : 2), cancellation.Token,
            mode == "stable-output" ? pidFile + ".output" : null);
        try
        {
            await WaitForFile(pidFile + ".child", running);
            var parent = int.Parse(await File.ReadAllTextAsync(pidFile));
            var child = int.Parse(await File.ReadAllTextAsync(pidFile + ".child"));
            if (mode == "parent-exits-first") await AssertExited(parent, allowZombie: true);
            elapsed.Restart();
            if (mode == "stable-output")
            {
                var result = await running;
                Assert.Equal(0, result.ExitCode);
                Assert.Equal("completed output", await File.ReadAllTextAsync(pidFile + ".output"));
            }
            else if (cancel)
            {
                cancellation.Cancel();
                await Assert.ThrowsAnyAsync<OperationCanceledException>(() => running);
            }
            else await Assert.ThrowsAsync<TimeoutException>(() => running);
            Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(8), "Cleanup exceeded its bounded deadline.");
            await AssertExited(parent);
            await AssertExited(child);
        }
        finally
        {
            cancellation.Cancel();
            try { await running; } catch (OperationCanceledException) { } catch (TimeoutException) { }
            // Baseline is intentionally red: clean only PIDs published by this fixture.
            foreach (var path in new[] { pidFile + ".child", pidFile })
            {
                if (!File.Exists(path) || !int.TryParse(await File.ReadAllTextAsync(path), out var pid)) continue;
                try
                {
                    using var process = Process.GetProcessById(pid);
                    if (!process.HasExited) process.Kill();
                    process.WaitForExit(3000);
                }
                catch (ArgumentException) { }
            }
        }
    }

    private static string RuntimeConfig => Path.Combine(Path.GetDirectoryName(typeof(Audit024RegressionTests).Assembly.Location)!,
        "Ofdrw.Net.Core.Tests.runtimeconfig.json");

    private static ProcessStartInfo Start(params string[] args)
    {
        var start = new ProcessStartInfo(Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet") { UseShellExecute = false };
        foreach (var arg in new[] { "exec", "--runtimeconfig", RuntimeConfig, typeof(ProcessFixtureMarker).Assembly.Location }.Concat(args))
            start.ArgumentList.Add(arg);
        return start;
    }

    private static async Task WaitForFile(string path, Task running)
    {
        var elapsed = Stopwatch.StartNew();
        while (!File.Exists(path) || new FileInfo(path).Length == 0)
        {
            if (running.IsFaulted) await running;
            Assert.True(elapsed.Elapsed < TimeSpan.FromSeconds(10), "Fixture did not publish its PID.");
            await Task.Delay(10);
        }
    }

    private static async Task AssertExited(int pid, bool allowZombie = false)
    {
        for (var i = 0; i < 100; i++)
        {
            try
            {
                using var process = Process.GetProcessById(pid);
                if (process.HasExited) return;
                // The native launcher deliberately retains an exited leader as
                // a waitable zombie to reserve its group ID until cleanup.
                if (allowZombie && !OperatingSystem.IsWindows())
                {
                    using var query = Process.Start(new ProcessStartInfo("/bin/ps", $"-p {pid} -o stat=")
                    { UseShellExecute = false, RedirectStandardOutput = true })!;
                    var state = await query.StandardOutput.ReadToEndAsync();
                    await query.WaitForExitAsync();
                    if (state.TrimStart().StartsWith("Z", StringComparison.Ordinal)) return;
                }
            }
            catch (ArgumentException) { return; }
            await Task.Delay(20);
        }
        Assert.Fail($"Owned process {pid} survived cleanup.");
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
