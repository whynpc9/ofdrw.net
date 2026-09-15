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

    // Windows\Fonts is huge. When a directory contains many faces, only stage
    // CJK families the Word templates actually name (宋体/黑体/等线/微软雅黑).
    private static readonly HashSet<string> PreferredCjkFontFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        "simsun.ttc",
        "simsunb.ttf",
        "simhei.ttf",
        "simkai.ttf",
        "simfang.ttf",
        "msyh.ttc",
        "msyhbd.ttc",
        "msyhl.ttc",
        "Deng.ttf",
        "Dengb.ttf",
        "Dengl.ttf",
        "NotoSansCJKsc-Regular.ttf",
        "Ofdrw-CI-NotoSansCJKsc-Regular.ttf"
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
            var files = Directory.EnumerateFiles(sourceDirectory)
                .Where(path => SupportedExtensions.Contains(Path.GetExtension(path)))
                .ToList();
            var copyAll = files.Count <= 32;
            foreach (var sourcePath in files)
            {
                var fileName = Path.GetFileName(sourcePath);
                if (!copyAll && !PreferredCjkFontFiles.Contains(fileName))
                {
                    continue;
                }

                var targetName = MakeUniqueName(fileName, usedNames);
                var targetPath = Path.Combine(targetDirectory, targetName);
                if (AlreadyStaged(sourcePath, targetPath))
                {
                    continue;
                }

                if (!TryCreateSymbolicLink(sourcePath, targetPath))
                {
                    File.Copy(sourcePath, targetPath, overwrite: true);
                }
            }
        }
    }

    private static bool AlreadyStaged(string sourcePath, string targetPath)
    {
        if (!File.Exists(targetPath))
        {
            return false;
        }

        return new FileInfo(sourcePath).Length == new FileInfo(targetPath).Length;
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
