using System;
using System.Collections.Generic;
using System.IO;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Loads a TrueType face for OFD fonts that named SimSun/SimHei but did not embed a file.
/// PDFsharp cannot use Windows .ttc collections, so this prefers .ttf stand-ins.
/// </summary>
internal static class CjkViewerFontLoader
{
    private static readonly string[] HeiFiles = { "simhei.ttf", "Deng.ttf" };
    private static readonly string[] SongFiles = { "simsun.ttf", "simhei.ttf", "Deng.ttf" };
    private static readonly string[] NotoFiles =
    {
        "NotoSansCJKsc-Regular.ttf",
        "Ofdrw-CI-NotoSansCJKsc-Regular.ttf"
    };

    internal static byte[]? TryRead(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName)) return null;
        var name = StripStyleSuffix(fontName!.Trim());
        string[] files;
        if (name.Equals("SimSun", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("NSimSun", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("宋体", StringComparison.OrdinalIgnoreCase))
        {
            files = SongFiles;
        }
        else if (name.Equals("SimHei", StringComparison.OrdinalIgnoreCase) ||
                 name.Equals("黑体", StringComparison.OrdinalIgnoreCase) ||
                 name.StartsWith("Noto Sans CJK", StringComparison.OrdinalIgnoreCase))
        {
            files = name.StartsWith("Noto Sans CJK", StringComparison.OrdinalIgnoreCase)
                ? Concat(NotoFiles, HeiFiles)
                : HeiFiles;
        }
        else
        {
            return null;
        }

        foreach (var directory in EnumerateDirectories())
        {
            foreach (var file in files)
            {
                var path = Path.Combine(directory, file);
                if (File.Exists(path)) return File.ReadAllBytes(path);
            }
        }

        return null;
    }

    private static string StripStyleSuffix(string family)
    {
        var separator = family.IndexOf('|');
        return separator < 0 ? family : family.Substring(0, separator);
    }

    private static string[] Concat(string[] first, string[] second)
    {
        var result = new string[first.Length + second.Length];
        first.CopyTo(result, 0);
        second.CopyTo(result, first.Length);
        return result;
    }

    private static IEnumerable<string> EnumerateDirectories()
    {
        var windows = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
        if (Directory.Exists(windows)) yield return windows;

        foreach (var directory in new[]
                 {
                     "/usr/share/fonts/opentype/noto",
                     "/usr/share/fonts/truetype/noto",
                     "/usr/share/fonts/truetype/wqy"
                 })
        {
            if (Directory.Exists(directory)) yield return directory;
        }
    }
}
