using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Ofdrw.Net.Converter.Pdf.Internal;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal static class LibreOfficeFontStager
{
    private static readonly HashSet<string> SupportedExtensions = new(StringComparer.OrdinalIgnoreCase)
    {
        ".otf",
        ".ttc",
        ".ttf"
    };

    // Windows\Fonts and other platform stores are huge. Those directories still
    // only stage CJK families the Word templates actually name.
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
            // Configured FontDirectories are copied in full. The CJK short-list only
            // applies to huge platform stores such as Windows\Fonts so LibreOffice is
            // not asked to ingest hundreds of unrelated faces.
            var copyAll = !ShouldRestrictToPreferredCjk(sourceDirectory, files.Count);
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

                if (File.Exists(targetPath))
                {
                    File.Delete(targetPath);
                }

                if (!TryCreateSymbolicLink(sourcePath, targetPath))
                {
                    File.Copy(sourcePath, targetPath, overwrite: true);
                }
            }
        }
    }

    internal static bool ShouldRestrictToPreferredCjk(string directory, int fontFileCount) =>
        fontFileCount > 32 && IsPlatformFontDirectory(directory);

    internal static bool IsPlatformFontDirectory(string directory)
    {
        if (string.IsNullOrWhiteSpace(directory))
        {
            return false;
        }

        var fullPath = NormalizeDirectory(directory);
        foreach (var platform in CjkViewerFontLoader.PlatformDirectories())
        {
            var platformPath = NormalizeDirectory(platform);
            if (fullPath.Equals(platformPath, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }

            if (fullPath.StartsWith(platformPath + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                fullPath.StartsWith(platformPath + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
            {
                return true;
            }
        }

        return false;
    }

    private static string NormalizeDirectory(string directory) =>
        Path.GetFullPath(directory).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);

    private static bool AlreadyStaged(string sourcePath, string targetPath)
    {
        if (!File.Exists(targetPath))
        {
            return false;
        }

        if (new FileInfo(sourcePath).Length != new FileInfo(targetPath).Length)
        {
            return false;
        }

        return FilesHaveSameContent(sourcePath, targetPath);
    }

    private static bool FilesHaveSameContent(string leftPath, string rightPath)
    {
        using var left = File.OpenRead(leftPath);
        using var right = File.OpenRead(rightPath);
        var leftBuffer = new byte[81920];
        var rightBuffer = new byte[81920];
        while (true)
        {
            var leftRead = left.Read(leftBuffer, 0, leftBuffer.Length);
            var rightRead = right.Read(rightBuffer, 0, rightBuffer.Length);
            if (leftRead != rightRead)
            {
                return false;
            }

            if (leftRead == 0)
            {
                return true;
            }

            for (var index = 0; index < leftRead; index++)
            {
                if (leftBuffer[index] != rightBuffer[index])
                {
                    return false;
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
