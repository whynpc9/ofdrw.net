using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Core.Processes;

internal sealed class ExternalProcessResult
{
    internal ExternalProcessResult(int exitCode, string output, string error)
    {
        ExitCode = exitCode;
        Output = output;
        Error = error;
    }

    internal int ExitCode { get; }
    internal string Output { get; }
    internal string Error { get; }
}

internal static class ExternalProcessRunner
{
    private const int MaximumCapturedCharacters = 65_536;

    internal static async Task<ExternalProcessResult> RunAsync(
        ProcessStartInfo startInfo,
        TimeSpan timeout,
        CancellationToken cancellationToken,
        string? expectedOutputPath = null)
    {
        if (timeout <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(timeout));
        }

        cancellationToken.ThrowIfCancellationRequested();
        startInfo.UseShellExecute = false;
        startInfo.CreateNoWindow = true;
        startInfo.RedirectStandardOutput = true;
        startInfo.RedirectStandardError = true;
        using var process = new Process { StartInfo = startInfo };
        if (!process.Start())
        {
            throw new InvalidOperationException($"{startInfo.FileName} could not be started.");
        }

        var output = CaptureAsync(process.StandardOutput);
        var error = CaptureAsync(process.StandardError);
        var pipes = Task.WhenAll(output, error);
        var elapsed = Stopwatch.StartNew();
        var stableSince = Stopwatch.StartNew();
        long previousLength = -1;
        var stoppedAfterOutput = false;
        try
        {
            // Include pipe completion in the deadline: a descendant may inherit
            // the handles even when its parent has already exited.
            while (!process.HasExited || !pipes.IsCompleted)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (elapsed.Elapsed >= timeout)
                {
                    throw new TimeoutException($"{startInfo.FileName} exceeded {timeout}.");
                }

                // Preserve the Office-wrapper compatibility path, but only for
                // callers explicitly supplying an expected output file.
                if (!process.HasExited && !string.IsNullOrWhiteSpace(expectedOutputPath) &&
                    File.Exists(expectedOutputPath))
                {
                    var length = new FileInfo(expectedOutputPath).Length;
                    if (length != previousLength)
                    {
                        previousLength = length;
                        stableSince.Restart();
                    }
                    else if (length > 0 && stableSince.Elapsed >= TimeSpan.FromSeconds(2))
                    {
                        KillTree(process);
                        stoppedAfterOutput = true;
                    }
                }

                await Task.Delay(25, cancellationToken).ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            await pipes.ConfigureAwait(false);
            var exitCode = stoppedAfterOutput ? 0 : process.ExitCode;
            return new ExternalProcessResult(exitCode, output.Result, error.Result);
        }
        catch
        {
            KillTree(process);
            // Termination is bounded. Closing the process/streams below releases
            // inherited pipes even if an external program cannot be cleaned up.
            process.WaitForExit(5000);
            try
            {
                if (await Task.WhenAny(pipes, Task.Delay(1000)).ConfigureAwait(false) == pipes)
                {
                    await pipes.ConfigureAwait(false);
                }
            }
            catch (IOException)
            {
                // The original cancellation/timeout remains the reported cause.
            }

            throw;
        }
    }

    private static async Task<string> CaptureAsync(StreamReader reader)
    {
        var captured = new StringBuilder();
        var buffer = new char[4096];
        int read;
        while ((read = await reader.ReadAsync(buffer, 0, buffer.Length).ConfigureAwait(false)) > 0)
        {
            var remaining = MaximumCapturedCharacters - captured.Length;
            if (remaining > 0)
            {
                captured.Append(buffer, 0, Math.Min(read, remaining));
            }
        }

        return captured.ToString();
    }

    private static void KillTree(Process process)
    {
        try
        {
            if (process.HasExited)
            {
                return;
            }

            // Keep netstandard2.0 compatibility while using full descendant
            // termination on modern .NET (including the CLI/test hosts).
            var killTree = typeof(Process).GetMethod("Kill", new[] { typeof(bool) });
            if (killTree is not null)
            {
                killTree.Invoke(process, new object[] { true });
                return;
            }

            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                using var killer = Process.Start(new ProcessStartInfo
                {
                    FileName = "taskkill", Arguments = $"/PID {process.Id} /T /F",
                    UseShellExecute = false, CreateNoWindow = true
                });
                killer?.WaitForExit(3000);
            }
            else
            {
                KillUnixChildren(process.Id);
            }

            if (!process.HasExited)
            {
                process.Kill();
            }
        }
        catch (Exception exception) when (exception is InvalidOperationException or
            System.ComponentModel.Win32Exception or TargetInvocationException)
        {
            // The process may exit while its descendants are being enumerated.
            try { if (!process.HasExited) process.Kill(); }
            catch (InvalidOperationException) { }
            catch (System.ComponentModel.Win32Exception) { }
        }
    }

    private static void KillUnixChildren(int parent)
    {
        using var query = Process.Start(new ProcessStartInfo
        {
            FileName = "pgrep", Arguments = $"-P {parent}",
            UseShellExecute = false, RedirectStandardOutput = true, CreateNoWindow = true
        });
        if (query is null) return;
        var children = query.StandardOutput.ReadToEnd();
        query.WaitForExit(1000);
        foreach (var value in children.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
        {
            if (int.TryParse(value, out var pid))
            {
                try
                {
                    using var child = Process.GetProcessById(pid);
                    KillTree(child);
                }
                catch (ArgumentException) { }
            }
        }
    }
}
