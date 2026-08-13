using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal static class LibreOfficeExecutableResolver
{
    internal static string Resolve(string? configuredPath)
    {
        if (!string.IsNullOrWhiteSpace(configuredPath))
        {
            return Normalize(ValidateExplicitPath(configuredPath!));
        }

        var environmentPath = Environment.GetEnvironmentVariable("OFDRW_LIBREOFFICE_PATH");
        if (!string.IsNullOrWhiteSpace(environmentPath))
        {
            return Normalize(ValidateExplicitPath(environmentPath!));
        }

        foreach (var candidate in GetKnownPaths())
        {
            if (File.Exists(candidate))
            {
                return Normalize(candidate);
            }
        }

        return RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? "soffice.com"
            : "soffice";
    }

    internal static string ResolveWorkingDirectory(string executable)
    {
        var directory = Path.GetDirectoryName(executable);
        if (string.IsNullOrWhiteSpace(directory))
        {
            return Environment.CurrentDirectory;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return directory;
        }

        if (IsProgramDirectory(directory))
        {
            return directory;
        }

        foreach (var programDirectory in GetUnixProgramDirectories())
        {
            if (Directory.Exists(programDirectory))
            {
                return programDirectory;
            }
        }

        return directory;
    }

    /// <summary>
    /// LibreOffice Portable / SecureUserConfig builds hang when forced onto a fresh
    /// <c>-env:UserInstallation</c> profile. Default isolation is therefore disabled for those builds.
    /// </summary>
    internal static bool ShouldIsolateUserProfile(string executable)
    {
        if (string.IsNullOrWhiteSpace(executable))
        {
            return false;
        }

        var normalized = executable.Replace(Path.AltDirectorySeparatorChar, Path.DirectorySeparatorChar);

        if (normalized.IndexOf("LibreOfficePortable", StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return false;
        }

        var portableMarker =
            $"{Path.DirectorySeparatorChar}App{Path.DirectorySeparatorChar}libreoffice{Path.DirectorySeparatorChar}program";
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows) &&
            normalized.IndexOf(portableMarker, StringComparison.OrdinalIgnoreCase) >= 0)
        {
            return false;
        }

        return !HasSecureUserConfig(normalized);
    }

    private static string ValidateExplicitPath(string path)
    {
        var trimmed = path.Trim();
        if (Directory.Exists(trimmed))
        {
            return ResolveExecutableFromDirectory(trimmed);
        }

        if ((Path.IsPathRooted(trimmed) ||
             trimmed.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
             trimmed.IndexOf(Path.AltDirectorySeparatorChar) >= 0) &&
            !File.Exists(trimmed))
        {
            throw new FileNotFoundException("The configured LibreOffice executable does not exist.", trimmed);
        }

        return trimmed;
    }

    private static string ResolveExecutableFromDirectory(string directory)
    {
        foreach (var candidate in new[]
                 {
                     Path.Combine(directory, "soffice.com"),
                     Path.Combine(directory, "soffice.exe"),
                     Path.Combine(directory, "soffice"),
                     Path.Combine(directory, "libreoffice"),
                     Path.Combine(directory, "program", "soffice.com"),
                     Path.Combine(directory, "program", "soffice.exe"),
                     Path.Combine(directory, "program", "soffice"),
                     Path.Combine(directory, "program", "libreoffice")
                 })
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        throw new FileNotFoundException(
            "The configured LibreOffice directory does not contain a recognizable executable.",
            directory);
    }

    private static string Normalize(string path)
    {
        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows) || !File.Exists(path))
        {
            return path;
        }

        // Prefer soffice.com: soffice.exe is a GUI subsystem binary and often hangs when
        // stdout/stderr are redirected from ProcessStartInfo.
        var directory = Path.GetDirectoryName(path) ?? string.Empty;
        var comPath = Path.Combine(directory, "soffice.com");
        return File.Exists(comPath) ? comPath : path;
    }

    private static bool HasSecureUserConfig(string executable)
    {
        try
        {
            var directory = Path.GetDirectoryName(executable);
            if (string.IsNullOrWhiteSpace(directory))
            {
                return false;
            }

            var iniPath = Path.Combine(directory, "soffice.ini");
            if (!File.Exists(iniPath))
            {
                return false;
            }

            foreach (var line in File.ReadLines(iniPath))
            {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("SecureUserConfig", StringComparison.OrdinalIgnoreCase) &&
                    trimmed.IndexOf("=true", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return true;
                }
            }
        }
        catch
        {
            // ignored
        }

        return false;
    }

    private static bool IsProgramDirectory(string directory)
    {
        var name = Path.GetFileName(directory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        return string.Equals(name, "program", StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<string> GetKnownPaths()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            yield return "/Applications/LibreOffice.app/Contents/MacOS/soffice";
            yield break;
        }

        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            if (!string.IsNullOrWhiteSpace(programFiles))
            {
                yield return Path.Combine(programFiles, "LibreOffice", "program", "soffice.com");
                yield return Path.Combine(programFiles, "LibreOffice", "program", "soffice.exe");
            }

            var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
            if (!string.IsNullOrWhiteSpace(programFilesX86))
            {
                yield return Path.Combine(programFilesX86, "LibreOffice", "program", "soffice.com");
                yield return Path.Combine(programFilesX86, "LibreOffice", "program", "soffice.exe");
            }

            yield break;
        }

        yield return "/usr/bin/libreoffice";
        yield return "/usr/bin/soffice";
        foreach (var programDirectory in GetUnixProgramDirectories())
        {
            yield return Path.Combine(programDirectory, "soffice");
        }
    }

    private static IEnumerable<string> GetUnixProgramDirectories()
    {
        yield return "/usr/lib/libreoffice/program";
        yield return "/usr/lib64/libreoffice/program";
        yield return "/opt/libreoffice/program";

        if (!Directory.Exists("/opt"))
        {
            yield break;
        }

        foreach (var entry in Directory.EnumerateDirectories("/opt"))
        {
            if (entry.IndexOf("libreoffice", StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            var programDirectory = Path.Combine(entry, "program");
            if (Directory.Exists(programDirectory))
            {
                yield return programDirectory;
            }
        }
    }
}
