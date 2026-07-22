using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Converter.Abstractions.Interfaces;
using Ofdrw.Net.Converter.Docx.Internal;

namespace Ofdrw.Net.Converter.Docx.Converters;

/// <summary>
/// Converts DOCX to PDF using Microsoft Word on macOS when available, with
/// isolated headless LibreOffice as the cross-platform fallback.
/// </summary>
public sealed class DocxToPdfConverter : IDocxToPdfConverter
{
    private static readonly SemaphoreSlim MicrosoftWordLock = new(1, 1);
    private readonly DocxConversionOptions _options;

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
        if (_options.ProcessTimeout <= TimeSpan.Zero)
        {
            throw new ArgumentOutOfRangeException(nameof(options), "ProcessTimeout must be greater than zero.");
        }
    }

    /// <inheritdoc />
    public async Task ConvertAsync(
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

        var engine = ResolveEngine();
        var workDirectory = CreateWorkDirectory(engine);
        var profileDirectory = Path.Combine(workDirectory, "profile");
        var inputPath = Path.Combine(workDirectory, "input.docx");
        var outputPath = Path.Combine(workDirectory, "input.pdf");

        Directory.CreateDirectory(workDirectory);
        try
        {
            using (var inputFile = File.Create(inputPath))
            {
                await docxInput.CopyToAsync(inputFile, 81920, cancellationToken).ConfigureAwait(false);
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
                Directory.CreateDirectory(profileDirectory);
                LibreOfficeFontStager.Stage(profileDirectory, _options);
                result = await RunLibreOfficeAsync(
                    inputPath,
                    workDirectory,
                    profileDirectory,
                    cancellationToken).ConfigureAwait(false);
            }

            if (result.ExitCode != 0)
            {
                throw new InvalidOperationException(
                    $"{GetEngineName(engine)} DOCX to PDF conversion failed with exit code " +
                    $"{result.ExitCode}: {result.Error}");
            }

            if (!File.Exists(outputPath) || new FileInfo(outputPath).Length == 0)
            {
                throw new InvalidOperationException(
                    $"{GetEngineName(engine)} did not produce a PDF. " +
                    $"Output: {result.Output} Error: {result.Error}");
            }

            using var pdfFile = File.OpenRead(outputPath);
            await pdfFile.CopyToAsync(pdfOutput, 81920, cancellationToken).ConfigureAwait(false);
        }
        finally
        {
            TryDeleteDirectory(workDirectory);
        }
    }

    private DocxConversionEngine ResolveEngine()
    {
        if (_options.Engine != DocxConversionEngine.Auto)
        {
            if (_options.Engine == DocxConversionEngine.MicrosoftWord && !CanUseMicrosoftWord())
            {
                throw new PlatformNotSupportedException(
                    "Microsoft Word DOCX rendering requires macOS, Microsoft Word in /Applications, " +
                    "and its local application container.");
            }

            return _options.Engine;
        }

        return CanUseMicrosoftWord()
            ? DocxConversionEngine.MicrosoftWord
            : DocxConversionEngine.LibreOffice;
    }

    private static bool CanUseMicrosoftWord()
    {
        return RuntimeInformation.IsOSPlatform(OSPlatform.OSX) &&
            File.Exists("/usr/bin/osascript") &&
            Directory.Exists("/Applications/Microsoft Word.app") &&
            Directory.Exists(GetMicrosoftWordTempRoot());
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
        string workingDirectory,
        string profileDirectory,
        CancellationToken cancellationToken)
    {
        var executable = LibreOfficeExecutableResolver.Resolve(_options.LibreOfficePath);
        var profileUri = new Uri(profileDirectory + Path.DirectorySeparatorChar).AbsoluteUri;
        var arguments = ProcessArguments.Join(
            $"-env:UserInstallation={profileUri}",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            workingDirectory,
            inputPath);

        return await RunProcessAsync(
            executable,
            arguments,
            workingDirectory,
            "LibreOffice",
            cancellationToken).ConfigureAwait(false);
    }

    private async Task<ProcessResult> RunProcessAsync(
        string executable,
        string arguments,
        string workingDirectory,
        string processName,
        CancellationToken cancellationToken)
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

        try
        {
            while (!process.HasExited)
            {
                cancellationToken.ThrowIfCancellationRequested();
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

        process.WaitForExit();
        return new ProcessResult(
            process.ExitCode,
            await outputTask.ConfigureAwait(false),
            await errorTask.ConfigureAwait(false));
    }

    private static string GetEngineName(DocxConversionEngine engine)
    {
        return engine == DocxConversionEngine.MicrosoftWord ? "Microsoft Word" : "LibreOffice";
    }

    private static void TryKill(Process process)
    {
        try
        {
            if (!process.HasExited)
            {
                process.Kill();
            }
        }
        catch
        {
            // Best effort cleanup after cancellation or timeout.
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
