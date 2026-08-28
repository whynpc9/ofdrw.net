using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Docx.Internal;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

namespace Ofdrw.Net.Converter.Docx.Converters;

/// <summary>
/// Converts DOCX to PDF using Microsoft Word on macOS when available, then
/// headless LibreOffice, with a cross-platform in-process fallback.
/// </summary>
public sealed class DocxToPdfConverter : IDocxToPdfConverter
{
    private static readonly SemaphoreSlim MicrosoftWordLock = new(1, 1);
    private static readonly SemaphoreSlim SharedLibreOfficeProfileGate = new(1, 1);
    private readonly DocxConversionOptions _options;
    private readonly Func<bool> _canUseMicrosoftWord;
    private readonly Func<bool> _canUseLibreOffice;

    /// <summary>
    /// Initializes a converter with default options.
    /// </summary>
    public DocxToPdfConverter()
        : this(new DocxConversionOptions())
    {
    }

    /// <summary>
    /// Initializes a converter with the supplied options.
    /// </summary>
    public DocxToPdfConverter(DocxConversionOptions options)
    {
        _options = options ?? throw new ArgumentNullException(nameof(options));
        _canUseMicrosoftWord = CanUseMicrosoftWord;
        _canUseLibreOffice = CanUseLibreOffice;
        if (_options.ProcessTimeout <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "ProcessTimeout must be greater than zero.");
        }

        if (_options.MaxInputBytes <= 0 || _options.MaxExpandedBytes <= 0 ||
            _options.MaxPackagePartCount <= 0 || _options.MaxDocumentElements <= 0 ||
            _options.MaxEmbeddedImageBytes <= 0 || _options.MaxEmbeddedImagePixels <= 0)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "BuiltIn resource limits must be positive.");
        }
    }

    internal DocxToPdfConverter(
        DocxConversionOptions options,
        Func<bool> canUseMicrosoftWord,
        Func<bool> canUseLibreOffice)
        : this(options)
    {
        _canUseMicrosoftWord = canUseMicrosoftWord ??
            throw new ArgumentNullException(nameof(canUseMicrosoftWord));
        _canUseLibreOffice = canUseLibreOffice ??
            throw new ArgumentNullException(nameof(canUseLibreOffice));
    }

    /// <inheritdoc />
    public async Task ConvertAsync(
        Stream docxInput,
        Stream pdfOutput,
        CancellationToken cancellationToken = default)
    {
        _ = await ConvertWithResultAsync(docxInput, pdfOutput, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>
    /// Converts DOCX to PDF and reports the engine attempts and rendering degradations.
    /// </summary>
    public async Task<DocxConversionResult> ConvertWithResultAsync(
        Stream docxInput,
        Stream pdfOutput,
        CancellationToken cancellationToken = default)
    {
        if (docxInput is null)
        {
            throw new ArgumentNullException(nameof(docxInput));
        }

        if (pdfOutput is null)
        {
            throw new ArgumentNullException(nameof(pdfOutput));
        }

        if (!docxInput.CanRead)
        {
            throw new ArgumentException("The DOCX input stream must be readable.", nameof(docxInput));
        }

        if (!pdfOutput.CanWrite)
        {
            throw new ArgumentException("The PDF output stream must be writable.", nameof(pdfOutput));
        }

        var stagingDirectory = Path.Combine(Path.GetTempPath(), $"ofdrw-docx-stage-{Guid.NewGuid():N}");
        var stagingInputPath = Path.Combine(stagingDirectory, "input.docx");
        var diagnostics = new List<DocxConversionDiagnostic>();
        var attempted = new List<DocxConversionEngine>();
        Exception? lastFailure = null;

        Directory.CreateDirectory(stagingDirectory);
        try
        {
            using (var inputFile = File.Create(stagingInputPath))
            {
                await CopyInputWithLimitAsync(docxInput, inputFile, cancellationToken).ConfigureAwait(false);
            }

            foreach (var engine in ResolveEngines(diagnostics))
            {
                cancellationToken.ThrowIfCancellationRequested();
                attempted.Add(engine);
                var workDirectory = CreateWorkDirectory(engine);
                var inputPath = Path.Combine(workDirectory, "input.docx");
                var outputPath = Path.Combine(workDirectory, "input.pdf");
                Directory.CreateDirectory(workDirectory);
                try
                {
                    File.Copy(stagingInputPath, inputPath, overwrite: true);
                    var engineDiagnostics = await RunEngineAsync(
                        engine,
                        inputPath,
                        outputPath,
                        workDirectory,
                        cancellationToken).ConfigureAwait(false);
                    ValidatePdfOutput(engine, outputPath);

                    using var pdfFile = File.OpenRead(outputPath);
                    await pdfFile.CopyToAsync(pdfOutput, 81920, cancellationToken).ConfigureAwait(false);
                    diagnostics.AddRange(engineDiagnostics);
                    return new DocxConversionResult(engine, attempted, diagnostics);
                }
                catch (OperationCanceledException)
                {
                    throw;
                }
                catch (Exception ex) when (CanFallback(engine, ex))
                {
                    lastFailure = ex;
                    diagnostics.Add(new DocxConversionDiagnostic(
                        "DOCX_ENGINE_FAILED",
                        $"{GetEngineName(engine)} failed; Auto is trying the next available engine.",
                        DocxConversionDiagnosticSeverity.Warning));
                }
                finally
                {
                    TryDeleteDirectory(workDirectory);
                }
            }

            throw new InvalidOperationException(
                "No available DOCX conversion engine produced a valid PDF.",
                lastFailure);
        }
        finally
        {
            TryDeleteDirectory(stagingDirectory);
        }
    }

    private IEnumerable<DocxConversionEngine> ResolveEngines(
        IList<DocxConversionDiagnostic> diagnostics)
    {
        if (_options.Engine != DocxConversionEngine.Auto)
        {
            if (_options.Engine == DocxConversionEngine.MicrosoftWord && !_canUseMicrosoftWord())
            {
                throw new PlatformNotSupportedException(
                    "Microsoft Word DOCX rendering requires macOS, Microsoft Word in /Applications, " +
                    "and its local application container.");
            }

            yield return _options.Engine;
            yield break;
        }

        if (_canUseMicrosoftWord())
        {
            yield return DocxConversionEngine.MicrosoftWord;
        }

        if (_canUseLibreOffice())
        {
            yield return DocxConversionEngine.LibreOffice;
        }
        else
        {
            diagnostics.Add(new DocxConversionDiagnostic(
                "DOCX_LIBREOFFICE_UNAVAILABLE",
                "LibreOffice is unavailable; Auto will use the in-process renderer."));
        }

        yield return DocxConversionEngine.BuiltIn;
    }

    private async Task<IReadOnlyList<DocxConversionDiagnostic>> RunEngineAsync(
        DocxConversionEngine engine,
        string inputPath,
        string outputPath,
        string workDirectory,
        CancellationToken cancellationToken)
    {
        if (engine == DocxConversionEngine.BuiltIn)
        {
            cancellationToken.ThrowIfCancellationRequested();
            using var timeout = CancellationTokenSource.CreateLinkedTokenSource(cancellationToken);
            timeout.CancelAfter(_options.ProcessTimeout);
            try
            {
                return await Task.Run(
                    () => new BuiltInDocxRenderer(_options).Convert(inputPath, outputPath, timeout.Token),
                    timeout.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (!cancellationToken.IsCancellationRequested)
            {
                throw new TimeoutException(
                    $"BuiltIn DOCX conversion exceeded {_options.ProcessTimeout}.");
            }
        }

        ProcessResult result;
        if (engine == DocxConversionEngine.MicrosoftWord)
        {
            result = await RunMicrosoftWordAsync(
                inputPath,
                outputPath,
                workDirectory,
                cancellationToken).ConfigureAwait(false);
        }
        else
        {
            result = await RunLibreOfficeAsync(
                inputPath,
                workDirectory,
                Path.Combine(workDirectory, "profile"),
                outputPath,
                cancellationToken).ConfigureAwait(false);
        }

        if (result.ExitCode != 0)
        {
            throw new InvalidOperationException(
                $"{GetEngineName(engine)} DOCX conversion failed with exit code " +
                $"{result.ExitCode}: {result.Error}");
        }

        return Array.Empty<DocxConversionDiagnostic>();
    }

    private static void ValidatePdfOutput(DocxConversionEngine engine, string outputPath)
    {
        if (!File.Exists(outputPath) || new FileInfo(outputPath).Length < 5)
        {
            throw new InvalidOperationException($"{GetEngineName(engine)} did not produce a PDF.");
        }

        using var stream = File.OpenRead(outputPath);
        var signature = new byte[5];
        if (stream.Read(signature, 0, signature.Length) != signature.Length ||
            signature[0] != '%' || signature[1] != 'P' || signature[2] != 'D' ||
            signature[3] != 'F' || signature[4] != '-')
        {
            throw new InvalidOperationException($"{GetEngineName(engine)} produced an invalid PDF.");
        }
    }

    private bool CanFallback(DocxConversionEngine engine, Exception exception)
    {
        if (_options.Engine != DocxConversionEngine.Auto || engine == DocxConversionEngine.BuiltIn)
        {
            return false;
        }

        return exception is not InvalidDataException &&
            exception is not NotSupportedException &&
            exception is not ArgumentException;
    }

    private async Task CopyInputWithLimitAsync(
        Stream input,
        Stream output,
        CancellationToken cancellationToken)
    {
        var buffer = new byte[81920];
        long total = 0;
        while (true)
        {
            var read = await input.ReadAsync(buffer, 0, buffer.Length, cancellationToken)
                .ConfigureAwait(false);
            if (read == 0)
            {
                break;
            }

            total = checked(total + read);
            if (total > _options.MaxInputBytes)
            {
                throw new InvalidDataException(
                    $"The DOCX input exceeds the configured {_options.MaxInputBytes} byte limit.");
            }

            await output.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }
    }

    private static bool CanUseMicrosoftWord()
    {
        return RuntimeInformation.IsOSPlatform(OSPlatform.OSX) &&
            File.Exists("/usr/bin/osascript") &&
            Directory.Exists("/Applications/Microsoft Word.app") &&
            Directory.Exists(GetMicrosoftWordTempRoot());
    }

    private bool CanUseLibreOffice()
    {
        try
        {
            var executable = LibreOfficeExecutableResolver.Resolve(_options.LibreOfficePath);
            if (Path.IsPathRooted(executable) ||
                executable.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
                executable.IndexOf(Path.AltDirectorySeparatorChar) >= 0)
            {
                return File.Exists(executable);
            }

            var pathValue = Environment.GetEnvironmentVariable("PATH") ?? string.Empty;
            foreach (var directory in pathValue.Split(Path.PathSeparator))
            {
                if (string.IsNullOrWhiteSpace(directory))
                {
                    continue;
                }

                if (File.Exists(Path.Combine(directory, executable)) ||
                    File.Exists(Path.Combine(directory, executable + ".exe")) ||
                    File.Exists(Path.Combine(directory, executable + ".com")))
                {
                    return true;
                }
            }

            return false;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return false;
        }
    }

    private static string CreateWorkDirectory(DocxConversionEngine engine)
    {
        var root = engine == DocxConversionEngine.MicrosoftWord
            ? GetMicrosoftWordTempRoot()
            : Path.GetTempPath();
        return Path.Combine(root, $"ofdrw-docx-{Guid.NewGuid():N}");
    }

    private static string GetMicrosoftWordTempRoot()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
            "Library",
            "Containers",
            "com.microsoft.Word",
            "Data",
            "tmp");
    }

    private async Task<ProcessResult> RunMicrosoftWordAsync(
        string inputPath,
        string outputPath,
        string workingDirectory,
        CancellationToken cancellationToken)
    {
        var arguments = ProcessArguments.Join(
            "-e", "on run argv",
            "-e", "set sourcePath to item 1 of argv",
            "-e", "set targetPath to item 2 of argv",
            "-e", "set sourceFile to POSIX file sourcePath",
            "-e", "set targetFile to POSIX file targetPath",
            "-e", "tell application \"Microsoft Word\"",
            "-e", "open sourceFile",
            "-e", "set convertedDocument to active document",
            "-e", "try",
            "-e", "save as convertedDocument file name targetFile file format format PDF",
            "-e", "on error errorMessage number errorNumber",
            "-e", "close convertedDocument saving no",
            "-e", "error errorMessage number errorNumber",
            "-e", "end try",
            "-e", "close convertedDocument saving no",
            "-e", "end tell",
            "-e", "end run",
            inputPath,
            outputPath);

        await MicrosoftWordLock.WaitAsync(cancellationToken).ConfigureAwait(false);
        try
        {
            return await RunProcessAsync(
                "/usr/bin/osascript",
                arguments,
                workingDirectory,
                "Microsoft Word",
                cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            MicrosoftWordLock.Release();
        }
    }

    private async Task<ProcessResult> RunLibreOfficeAsync(
        string inputPath,
        string workDirectory,
        string profileDirectory,
        string expectedOutputPath,
        CancellationToken cancellationToken)
    {
        var executable = LibreOfficeExecutableResolver.Resolve(_options.LibreOfficePath);
        var processWorkingDirectory = LibreOfficeExecutableResolver.ResolveWorkingDirectory(executable);
        var isolateUserProfile = _options.IsolateUserProfile
            ?? LibreOfficeExecutableResolver.ShouldIsolateUserProfile(executable);

        var gateEntered = false;
        if (!isolateUserProfile)
        {
            await SharedLibreOfficeProfileGate.WaitAsync(cancellationToken).ConfigureAwait(false);
            gateEntered = true;
        }

        try
        {
            string[] argumentParts;
            if (isolateUserProfile)
            {
                Directory.CreateDirectory(profileDirectory);
                LibreOfficeFontStager.Stage(profileDirectory, _options);
                var profileUri = new Uri(profileDirectory + Path.DirectorySeparatorChar).AbsoluteUri;
                argumentParts = new[]
                {
                    $"-env:UserInstallation={profileUri}",
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--norestore",
                    "--nolockcheck",
                    "--nodefault",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    workDirectory,
                    inputPath
                };
            }
            else
            {
                argumentParts = new[]
                {
                    "--headless",
                    "--nologo",
                    "--nofirststartwizard",
                    "--norestore",
                    "--nolockcheck",
                    "--nodefault",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    workDirectory,
                    inputPath
                };
            }

            var arguments = ProcessArguments.Join(argumentParts);
            return await RunProcessAsync(
                executable,
                arguments,
                processWorkingDirectory,
                "LibreOffice",
                cancellationToken,
                expectedOutputPath).ConfigureAwait(false);
        }
        finally
        {
            if (gateEntered)
            {
                SharedLibreOfficeProfileGate.Release();
            }
        }
    }

    private async Task<ProcessResult> RunProcessAsync(
        string executable,
        string arguments,
        string workingDirectory,
        string processName,
        CancellationToken cancellationToken,
        string? expectedOutputPath = null)
    {
        using var process = new Process
        {
            StartInfo = new ProcessStartInfo
            {
                FileName = executable,
                Arguments = arguments,
                WorkingDirectory = workingDirectory,
                UseShellExecute = false,
                CreateNoWindow = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            }
        };

        try
        {
            if (!process.Start())
            {
                throw new InvalidOperationException($"{processName} could not be started.");
            }
        }
        catch (Win32Exception ex)
        {
            var message = processName == "LibreOffice"
                ? "LibreOffice was not found. Install LibreOffice, put soffice on PATH, " +
                    "set OFDRW_LIBREOFFICE_PATH, or configure DocxConversionOptions.LibreOfficePath."
                : "Microsoft Word automation could not be started. Ensure Word is installed and " +
                    "allow the host process to control Microsoft Word when macOS requests permission.";
            throw new InvalidOperationException(message, ex);
        }

        var outputTask = process.StandardOutput.ReadToEndAsync();
        var errorTask = process.StandardError.ReadToEndAsync();
        var elapsed = Stopwatch.StartNew();
        long lastPdfLength = -1;
        var pdfStableSince = Stopwatch.StartNew();

        try
        {
            while (!process.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();

                // Some LibreOffice wrappers write the PDF but never exit. Treat a
                // non-empty, size-stable output file as success and stop the tree.
                if (!string.IsNullOrWhiteSpace(expectedOutputPath) &&
                    File.Exists(expectedOutputPath))
                {
                    var length = new FileInfo(expectedOutputPath).Length;
                    if (length > 0)
                    {
                        if (length != lastPdfLength)
                        {
                            lastPdfLength = length;
                            pdfStableSince.Restart();
                        }
                        else if (pdfStableSince.Elapsed >= TimeSpan.FromSeconds(2))
                        {
                            TryKill(process);
                            break;
                        }
                    }
                }

                if (elapsed.Elapsed >= _options.ProcessTimeout)
                {
                    throw new TimeoutException(
                        $"{processName} DOCX conversion exceeded {_options.ProcessTimeout}.");
                }

                await Task.Delay(100, cancellationToken).ConfigureAwait(false);
            }
        }
        catch
        {
            TryKill(process);
            throw;
        }

        if (!process.HasExited)
        {
            if (!process.WaitForExit(5000))
            {
                TryKill(process);
                process.WaitForExit(5000);
            }
        }
        else
        {
            process.WaitForExit();
        }

        var exitCode = -1;
        try
        {
            if (process.HasExited)
            {
                exitCode = process.ExitCode;
            }
        }
        catch
        {
            exitCode = -1;
        }

        if (exitCode != 0 &&
            !string.IsNullOrWhiteSpace(expectedOutputPath) &&
            File.Exists(expectedOutputPath) &&
            new FileInfo(expectedOutputPath).Length > 0)
        {
            exitCode = 0;
        }

        return new ProcessResult(
            exitCode,
            await outputTask.ConfigureAwait(false),
            await errorTask.ConfigureAwait(false));
    }

    private static string GetEngineName(DocxConversionEngine engine)
    {
        return engine switch
        {
            DocxConversionEngine.MicrosoftWord => "Microsoft Word",
            DocxConversionEngine.LibreOffice => "LibreOffice",
            DocxConversionEngine.BuiltIn => "BuiltIn",
            _ => engine.ToString()
        };
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (process.HasExited)
            {
                return;
            }

            var pid = process.Id;
            if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
            {
                TryKillWindowsProcessTree(pid);
                return;
            }

            TryKillUnixProcessTree(pid);
        }
        catch
        {
            // Best effort cleanup after cancellation or timeout.
        }
    }

    private static void TryKillWindowsProcessTree(int pid)
    {
        using var killer = Process.Start(new ProcessStartInfo
        {
            FileName = "taskkill",
            Arguments = $"/PID {pid} /T /F",
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        });
        killer?.WaitForExit(10000);
    }

    private static void TryKillUnixProcessTree(int pid)
    {
        try
        {
            using var pkill = Process.Start(new ProcessStartInfo
            {
                FileName = "pkill",
                Arguments = $"-KILL -P {pid}",
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            });
            pkill?.WaitForExit(3000);
        }
        catch
        {
            // ignored
        }

        try
        {
            using var target = Process.GetProcessById(pid);
            if (!target.HasExited)
            {
                target.Kill();
            }
        }
        catch
        {
            // ignored
        }
    }

    private static void TryDeleteDirectory(string path)
    {
        try
        {
            if (Directory.Exists(path))
            {
                Directory.Delete(path, recursive: true);
            }
        }
        catch
        {
            // A failed cleanup must not hide the conversion result.
        }
    }

    private sealed class ProcessResult
    {
        internal ProcessResult(int exitCode, string output, string error)
        {
            ExitCode = exitCode;
            Output = output;
            Error = error;
        }

        internal int ExitCode { get; }

        internal string Output { get; }

        internal string Error { get; }
    }
}
