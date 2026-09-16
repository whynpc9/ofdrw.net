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
    private static readonly string[] PreferredFiles =
    {
        "Ofdrw-CI-NotoSansCJKsc-Regular.ttf",
        "NotoSansCJKsc-Regular.ttf",
        "simhei.ttf",
        "simsun.ttf",
        "Deng.ttf"
    };

    internal static byte[]? TryRead(string? fontName)
    {
        if (string.IsNullOrWhiteSpace(fontName)) return null;
        var name = StripStyleSuffix(fontName!.Trim());
        if (!IsCjkFaceName(name)) return null;

        foreach (var directory in PlatformDirectories())
        {
            foreach (var file in PreferredFiles)
            {
                var path = Path.Combine(directory, file);
                if (File.Exists(path)) return File.ReadAllBytes(path);
            }

            try
            {
                foreach (var path in Directory.GetFiles(directory, "Ofdrw-CI-Noto*.ttf"))
                    return File.ReadAllBytes(path);
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            {
            }
        }

        return null;
    }

    internal static IEnumerable<string> PlatformDirectories()
    {
        var home = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        var windows = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
        foreach (var directory in new[]
                 {
                     string.IsNullOrEmpty(windows) ? null : Path.Combine(windows, "Fonts"),
                     string.IsNullOrEmpty(home) ? null : Path.Combine(home, ".fonts"),
                     string.IsNullOrEmpty(home) ? null : Path.Combine(home, "Library", "Fonts"),
                     "/usr/share/fonts",
                     "/usr/local/share/fonts",
                     "/usr/share/fonts/opentype/noto",
                     "/usr/share/fonts/truetype/noto",
                     "/usr/share/fonts/truetype/wqy",
                     "/usr/share/fonts/noto-cjk",
                     "/System/Library/Fonts",
                     "/Library/Fonts"
                 })
        {
            if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                yield return directory;
        }
    }

    private static bool IsCjkFaceName(string name) =>
        name.Equals("SimSun", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("NSimSun", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("宋体", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("新宋体", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("SimHei", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("黑体", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("DengXian", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("等线", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("KaiTi", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("楷体", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("FangSong", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("仿宋", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Microsoft YaHei", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("微软雅黑", StringComparison.OrdinalIgnoreCase) ||
        name.StartsWith("Noto Sans CJK", StringComparison.OrdinalIgnoreCase);

    private static string StripStyleSuffix(string family)
    {
        var separator = family.IndexOf('|');
        return separator < 0 ? family : family.Substring(0, separator);
    }
}
