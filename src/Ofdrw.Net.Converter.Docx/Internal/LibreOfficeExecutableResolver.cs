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
            return ValidateExplicitPath(configuredPath!);
        }

        var environmentPath = Environment.GetEnvironmentVariable("OFDRW_LIBREOFFICE_PATH");
        if (!string.IsNullOrWhiteSpace(environmentPath))
        {
            return ValidateExplicitPath(environmentPath!);
        }

        foreach (var candidate in GetKnownPaths())
        {
            if (File.Exists(candidate))
            {
                return candidate;
            }
        }

        return RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
            ? "soffice.exe"
            : "soffice";
    }

    private static string ValidateExplicitPath(string path)
    {
        var trimmed = path.Trim();
        if ((Path.IsPathRooted(trimmed) || trimmed.IndexOf(Path.DirectorySeparatorChar) >= 0 ||
             trimmed.IndexOf(Path.AltDirectorySeparatorChar) >= 0) && !File.Exists(trimmed))
        {
            throw new FileNotFoundException("The configured LibreOffice executable does not exist.", trimmed);
        }

        return trimmed;
    }

    private static IEnumerable<string> GetKnownPaths()
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            yield return "/Applications/LibreOffice.app/Contents/MacOS/soffice";
            yield break;
        }

        if (!RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            yield break;
        }

        var programFiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        if (!string.IsNullOrWhiteSpace(programFiles))
        {
            yield return Path.Combine(programFiles, "LibreOffice", "program", "soffice.exe");
        }

        var programFilesX86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        if (!string.IsNullOrWhiteSpace(programFilesX86))
        {
            yield return Path.Combine(programFilesX86, "LibreOffice", "program", "soffice.exe");
        }
    }
}
