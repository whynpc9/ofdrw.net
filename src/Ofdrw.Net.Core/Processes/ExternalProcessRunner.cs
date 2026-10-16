using System;
using System.Diagnostics;
using System.IO;
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
        var elapsed = Stopwatch.StartNew();
        var process = OwnedProcessGroup.Start(startInfo, timeout, elapsed, cancellationToken);

        var output = CaptureAsync(process.StandardOutput);
        var error = CaptureAsync(process.StandardError);
        var pipes = Task.WhenAll(output, error);
        var stableSince = Stopwatch.StartNew();
        long previousLength = -1;
        var stoppedAfterOutput = false;
        Exception? failure = null;
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
                        process.Terminate();
                        stoppedAfterOutput = true;
                    }
                }

                await Task.Delay(25, cancellationToken).ConfigureAwait(false);
            }

            cancellationToken.ThrowIfCancellationRequested();
            await pipes.ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (elapsed.Elapsed >= timeout) throw new TimeoutException($"{startInfo.FileName} exceeded {timeout}.");
            var exitCode = stoppedAfterOutput ? 0 : process.ExitCode;
            return new ExternalProcessResult(exitCode, output.Result, error.Result);
        }
        catch (Exception exception)
        {
            failure = exception;
            try
            {
                process.Terminate();
                process.WaitForExit(5000);
                if (await Task.WhenAny(pipes, Task.Delay(1000)).ConfigureAwait(false) == pipes)
                    await pipes.ConfigureAwait(false);
            }
            catch (Exception cleanupException)
            {
                exception.Data["OwnedProcessCleanup"] = cleanupException.ToString();
            }
            throw;
        }
        finally
        {
            try { process.Dispose(); }
            catch (Exception cleanupException) when (failure != null)
            {
                failure.Data["OwnedProcessDisposal"] = cleanupException.ToString();
            }
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

}
