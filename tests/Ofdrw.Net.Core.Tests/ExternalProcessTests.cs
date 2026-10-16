using System.Diagnostics;
using System.ComponentModel;
using System.Text;
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Runner_ShouldPreserveArgumentVector(bool useRawArguments)
    {
        var arguments = new[] { "", "one two", "中文 空格", "back\\slash", "quote\"x", "trail\\", "two\"\"quotes", "tab\tvalue", "'$(echo unsafe);*" };
        var start = CreateStart(new[] { "arguments" }.Concat(arguments).ToArray());
        if (useRawArguments)
        {
            start.Arguments = string.Join(" ", start.ArgumentList.Select(QuoteArgument));
            start.ArgumentList.Clear();
        }
        var result = await ExternalProcessRunner.RunAsync(start, TimeSpan.FromSeconds(10), default);
        Assert.Equal(0, result.ExitCode);
        var actual = result.Output.Replace("\r", "").Split('\n').SkipLast(1)
            .Select(value => Encoding.UTF8.GetString(Convert.FromBase64String(value)));
        Assert.Equal(arguments, actual);
    }

    [Fact]
    public async Task Runner_ShouldPreserveWorkingDirectoryAndEnvironment()
    {
        var directory = Path.Combine(_directory, "中文 directory");
        Directory.CreateDirectory(directory);
        var marker = Guid.NewGuid().ToString("N");
        await File.WriteAllTextAsync(Path.Combine(directory, "cwd-marker"), marker);
        var start = CreateStart("context");
        start.StandardOutputEncoding = Encoding.UTF8;
        start.WorkingDirectory = directory;
        start.Environment["OFDRW_AUDIT024_VALUE"] = "value 空格 \" \\";
        var result = await ExternalProcessRunner.RunAsync(start, TimeSpan.FromSeconds(10), default);
        var lines = result.Output.Replace("\r", "").Split('\n');
        Assert.Equal(marker, await File.ReadAllTextAsync(Path.Combine(lines[0], "cwd-marker")));
        Assert.Equal(start.Environment["OFDRW_AUDIT024_VALUE"], lines[1]);
        Assert.Equal(0, result.ExitCode);
    }

    [Fact]
    public async Task Runner_ShouldReportMissingExecutableAsNativeStartFailure()
    {
        var start = new ProcessStartInfo(Path.Combine(_directory, "missing-executable"));
        await Assert.ThrowsAsync<Win32Exception>(() => ExternalProcessRunner.RunAsync(start, TimeSpan.FromSeconds(10), default));
    }

    [Fact]
    public async Task Runner_ShouldRejectCancellationBeforeStarting()
    {
        var pidFile = Path.Combine(_directory, "must-not-start");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ExternalProcessRunner.RunAsync(
            CreateStart("wait", pidFile), TimeSpan.FromSeconds(10), cancellation.Token));
        Assert.False(File.Exists(pidFile));
    }

    private static string QuoteArgument(string argument)
    {
        var result = new StringBuilder("\"");
        var slashes = 0;
        foreach (var character in argument)
        {
            if (character == '\\') { slashes++; continue; }
            result.Append('\\', character == '\"' ? slashes * 2 + 1 : slashes).Append(character);
            slashes = 0;
        }
        return result.Append('\\', slashes * 2).Append('\"').ToString();
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
