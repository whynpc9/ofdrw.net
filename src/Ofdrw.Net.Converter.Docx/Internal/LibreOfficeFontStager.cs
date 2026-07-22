using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal static class LibreOfficeFontStager
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".otf",
        ".ttc",
        ".ttf"
    };

    internal static void Stage(string profileDirectory, DocxConversionOptions options)
    {
        var directories = new List<string>();
        foreach (var configuredDirectory in options.FontDirectories)
        {
            if (string.IsNullOrWhiteSpace(configuredDirectory))
            {
                continue;
            }

            var fullPath = Path.GetFullPath(configuredDirectory);
            if (!Directory.Exists(fullPath))
            {
                throw new DirectoryNotFoundException($"The configured font directory does not exist: {fullPath}");
            }

            directories.Add(fullPath);
        }

        if (options.UseInstalledMicrosoftOfficeFonts && RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
        {
            AddIfPresent(directories, "/Applications/Microsoft Word.app/Contents/Resources/DFonts");
            var userApplications = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.UserProfile),
                "Applications",
                "Microsoft Word.app",
                "Contents",
                "Resources",
                "DFonts");
            AddIfPresent(directories, userApplications);
        }

        if (directories.Count == 0)
        {
            return;
        }

        var targetDirectory = Path.Combine(profileDirectory, "user", "fonts");
        Directory.CreateDirectory(targetDirectory);
        var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var sourceDirectory in directories.Distinct(StringComparer.OrdinalIgnoreCase))
        {
            foreach (var sourcePath in Directory.EnumerateFiles(sourceDirectory))
            {
                if (!SupportedExtensions.Contains(Path.GetExtension(sourcePath)))
                {
                    continue;
                }

                var targetName = MakeUniqueName(Path.GetFileName(sourcePath), usedNames);
                var targetPath = Path.Combine(targetDirectory, targetName);
                if (!TryCreateSymbolicLink(sourcePath, targetPath))
                {
                    File.Copy(sourcePath, targetPath, overwrite: false);
                }
            }
        }
    }

    private static void AddIfPresent(ICollection<string> directories, string path)
    {
        if (Directory.Exists(path))
        {
            directories.Add(path);
        }
    }

    private static string MakeUniqueName(string fileName, ISet<string> usedNames)
    {
        if (usedNames.Add(fileName))
        {
            return fileName;
        }

        var stem = Path.GetFileNameWithoutExtension(fileName);
        var extension = Path.GetExtension(fileName);
        for (var suffix = 2; ; suffix++)
        {
            var candidate = $"{stem}-{suffix}{extension}";
            if (usedNames.Add(candidate))
            {
                return candidate;
            }
        }
    }

    private static bool TryCreateSymbolicLink(string sourcePath, string targetPath)
    {
        if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
        {
            return false;
        }

        try
        {
            return Symlink(sourcePath, targetPath) == 0;
        }
        catch (DllNotFoundException)
        {
            return false;
        }
        catch (EntryPointNotFoundException)
        {
            return false;
        }
    }

    [DllImport("libc", EntryPoint = "symlink", SetLastError = true, CharSet = CharSet.Ansi)]
    private static extern int Symlink(string target, string linkPath);
}
