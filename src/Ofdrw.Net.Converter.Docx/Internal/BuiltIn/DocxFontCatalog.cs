using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Internal;
using SixLabors.Fonts;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

/// <summary>Per-conversion mappings from configured font families to immutable content identities.</summary>
internal sealed class DocxFontCatalog
{
    // Word templates name 宋体/SimSun, but Windows ships that face as simsun.ttc.
    // FontDescription and PDFsharp RegisterFace only accept TTF/OTF, so the catalog
    // never sees SimSun and the host resolver substitutes a Latin face (Agency FB).
    // Alias each CJK family group onto an available TTF in the same group, then
    // fill empty groups from a CJK TTF substitute (SimHei, DengXian, Noto, …).
    private static readonly string[][] CjkAliasGroups =
    {
        new[] { "SimSun", "NSimSun", "宋体", "新宋体" },
        new[] { "SimHei", "黑体" },
        new[] { "DengXian", "等线" },
        new[] { "KaiTi", "楷体" },
        new[] { "FangSong", "仿宋" },
        new[] { "Microsoft YaHei", "Microsoft YaHei UI", "微软雅黑" },
        new[] { "Noto Sans CJK SC", "Noto Sans CJK SC Regular" }
    };

    private static readonly string[] CjkSubstituteOrder =
    {
        "SimSun", "NSimSun", "SimHei", "Noto Sans CJK SC", "DengXian", "KaiTi", "FangSong", "Microsoft YaHei"
    };

    private static readonly HashSet<string> ViewerLocalCjkFamilies = new(StringComparer.OrdinalIgnoreCase)
    {
        "SimSun", "NSimSun", "宋体", "新宋体",
        "SimHei", "黑体",
        "DengXian", "等线",
        "KaiTi", "楷体",
        "FangSong", "仿宋",
        "Microsoft YaHei", "Microsoft YaHei UI", "微软雅黑"
    };

    private static readonly HashSet<string> PreferredCjkCollectionFiles = new(StringComparer.OrdinalIgnoreCase)
    {
        "simsun.ttc",
        "msyh.ttc",
        "msyhbd.ttc",
        "msyhl.ttc"
    };

    private readonly Dictionary<string, List<(byte[] Data, FontStyle Style)>> _fonts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _resolved = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> _authentic = new(StringComparer.OrdinalIgnoreCase);

    internal DocxFontCatalog(DocxConversionOptions options, IList<DocxConversionDiagnostic> diagnostics, CancellationToken cancellationToken)
    {
        PdfFontRegistry.EnsureInstalled();
        var configured = options.FontDirectories
            .Where(directory => !string.IsNullOrWhiteSpace(directory))
            .Select(directory => directory.Trim())
            .ToList();
        foreach (var directory in configured)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!Directory.Exists(directory))
            {
                diagnostics.Add(new DocxConversionDiagnostic("DOCX_FONT_DIRECTORY_MISSING", "A configured font directory does not exist."));
                continue;
            }
            LoadDirectory(directory, options, diagnostics, cancellationToken);
        }

        // Empty FontDirectories (Linux ARM without Windows\\Fonts) still needs Noto/WQY.
        if (configured.Count == 0)
        {
            foreach (var directory in DefaultPlatformFontDirectories())
                LoadDirectory(directory, options, diagnostics, cancellationToken);
        }

        ApplyCjkAliases();
    }

    internal static IEnumerable<string> DefaultPlatformFontDirectories()
    {
        foreach (var directory in new[]
                 {
                     Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts"),
                     "/usr/share/fonts",
                     "/usr/local/share/fonts",
                     "/usr/share/fonts/opentype/noto",
                     "/usr/share/fonts/truetype/noto",
                     "/usr/share/fonts/truetype/wqy",
                     "/usr/share/fonts/noto-cjk",
                     "/System/Library/Fonts",
                     "/Library/Fonts",
                     Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".fonts")
                 })
        {
            if (!string.IsNullOrWhiteSpace(directory) && Directory.Exists(directory))
                yield return directory;
        }
    }

    internal string Resolve(string family, bool bold, bool italic)
    {
        var key = family + $"\u001f{bold}\u001f{italic}";
        if (_resolved.TryGetValue(key, out var resolved)) return resolved;
        if (!TryGetFaces(family, out var faces)) return family;
        var style = (bold ? FontStyle.Bold : FontStyle.Regular) | (italic ? FontStyle.Italic : FontStyle.Regular);
        var face = faces.FirstOrDefault(candidate => candidate.Style == style);
        if (face.Data is null) face = faces.FirstOrDefault(candidate => candidate.Style == FontStyle.Regular);
        if (face.Data is null) face = faces[0];
        resolved = PdfFontRegistry.RegisterFontFace(face.Data, bold, italic);
        _resolved.Add(key, resolved);
        return resolved;
    }

    internal string ResolveFallbackFamily(IEnumerable<string> families)
    {
        string? first = null;
        string? systemMatch = null;
        foreach (var candidate in families)
        {
            if (string.IsNullOrWhiteSpace(candidate)) continue;
            var family = candidate.Trim();
            first ??= family;
            if (_fonts.ContainsKey(family)) return family;
            if (systemMatch is null && SystemFonts.TryGet(family, out _))
                systemMatch = family;
        }

        // Windows SystemFonts can see SimSun via simsun.ttc. Returning that name
        // without catalog bytes lets PDFsharp pick a Latin substitute.
        if (systemMatch is not null && !IsKnownCjkFamily(systemMatch))
            return systemMatch;

        foreach (var name in CjkSubstituteOrder)
        {
            if (_fonts.ContainsKey(name)) return name;
        }

        // A host resolver may know names outside SystemFonts. Preserve its
        // opportunity to resolve the first configured family as a last resort.
        return first ?? GlobalFontSettings.FontResolver.DefaultFontName;
    }

    /// <summary>
    /// Viewer-local CJK names stay unembedded, matching name-only OFD like linux1.
    /// Catalog TTF/TTC bytes are still used to measure advances so CJK does not collapse.
    /// </summary>
    internal bool ShouldEmbed(string family)
    {
        if (string.IsNullOrWhiteSpace(family)) return false;
        var baseName = StripStyleSuffix(family);
        if (IsViewerLocalCjkFamily(baseName)) return false;
        if (baseName.StartsWith("Noto Sans CJK", StringComparison.OrdinalIgnoreCase)) return false;
        return _authentic.Contains(baseName) || !IsKnownCjkFamily(baseName);
    }

    internal static bool IsViewerLocalCjkFamily(string family) =>
        ViewerLocalCjkFamilies.Contains(StripStyleSuffix(family));

    private void LoadDirectory(
        string directory,
        DocxConversionOptions options,
        IList<DocxConversionDiagnostic> diagnostics,
        CancellationToken cancellationToken)
    {
        foreach (var path in EnumerateFontFiles(directory).OrderBy(candidate => candidate, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var extension = Path.GetExtension(path);
            var isCollection = extension.Equals(".ttc", StringComparison.OrdinalIgnoreCase);
            var isFace = extension.Equals(".ttf", StringComparison.OrdinalIgnoreCase) ||
                         extension.Equals(".otf", StringComparison.OrdinalIgnoreCase);
            if (!isFace && !isCollection) continue;
            if (isCollection && !IsPreferredCjkCollection(Path.GetFileName(path))) continue;
            try
            {
                if (new FileInfo(path).Length > options.MaxEmbeddedFontBytes)
                    throw new InvalidDataException("Configured font exceeds MaxEmbeddedFontBytes.");
                var bytes = File.ReadAllBytes(path);
                if (isCollection)
                {
                    foreach (var face in OpenTypeCollection.ExtractFaces(bytes))
                        AddFace(fileStem: null, face);
                }
                else
                {
                    AddFace(Path.GetFileNameWithoutExtension(path), bytes);
                }
            }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException or InvalidDataException)
            {
                diagnostics.Add(new DocxConversionDiagnostic("DOCX_FONT_LOAD_FAILED", $"Configured font '{Path.GetFileName(path)}' could not be loaded: {exception.Message}"));
            }
        }
    }

    private static bool IsPreferredCjkCollection(string fileName)
    {
        if (PreferredCjkCollectionFiles.Contains(fileName)) return true;
        var name = fileName.ToLowerInvariant();
        return name.Contains("cjk") || name.Contains("noto") || name.Contains("wqy") ||
               name.Contains("sourcehan") || name.Contains("simsun") || name.Contains("msyh") ||
               name.Contains("uming") || name.Contains("ukai") || name.Contains("droid");
    }

    private static IEnumerable<string> EnumerateFontFiles(string directory)
    {
        var pending = new Stack<string>();
        pending.Push(directory);
        while (pending.Count > 0)
        {
            var current = pending.Pop();
            string[] files;
            string[] children;
            try { files = Directory.GetFiles(current); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            { files = Array.Empty<string>(); }
            try { children = Directory.GetDirectories(current); }
            catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
            { children = Array.Empty<string>(); }
            foreach (var file in files) yield return file;
            foreach (var child in children) pending.Push(child);
        }
    }

    private static string StripStyleSuffix(string family)
    {
        var separator = family.IndexOf('|');
        return separator < 0 ? family : family.Substring(0, separator);
    }

    private bool TryGetFaces(string family, out List<(byte[] Data, FontStyle Style)> faces)
    {
        if (_fonts.TryGetValue(family, out faces!)) return true;
        if (IsKnownCjkFamily(family))
        {
            foreach (var name in CjkSubstituteOrder)
            {
                if (_fonts.TryGetValue(name, out faces!)) return true;
            }
        }

        faces = null!;
        return false;
    }

    private void ApplyCjkAliases()
    {
        foreach (var group in CjkAliasGroups)
        {
            List<(byte[] Data, FontStyle Style)>? source = null;
            foreach (var name in group)
            {
                if (_fonts.TryGetValue(name, out var faces))
                {
                    source = faces;
                    break;
                }
            }

            if (source is null) continue;
            foreach (var name in group)
            {
                if (!_fonts.ContainsKey(name)) _fonts[name] = source;
            }
        }

        List<(byte[] Data, FontStyle Style)>? substitute = null;
        foreach (var name in CjkSubstituteOrder)
        {
            if (_fonts.TryGetValue(name, out var faces))
            {
                substitute = faces;
                break;
            }
        }

        if (substitute is null) return;
        foreach (var group in CjkAliasGroups)
        {
            if (group.Any(name => _fonts.ContainsKey(name))) continue;
            foreach (var name in group)
            {
                if (!_fonts.ContainsKey(name)) _fonts[name] = substitute;
            }
        }
    }

    private void AddFace(string? fileStem, byte[] bytes)
    {
        using var stream = new MemoryStream(bytes, writable: false);
        var description = FontDescription.LoadDescription(stream);
        if (!string.IsNullOrWhiteSpace(fileStem))
            Add(fileStem, bytes, description.Style);
        Add(description.FontFamilyInvariantCulture, bytes, description.Style);
        AddLocalizedFamily(description, bytes);
    }

    private void AddLocalizedFamily(FontDescription description, byte[] bytes)
    {
        try
        {
            var localized = description.FontFamily(CultureInfo.GetCultureInfo("zh-CN"));
            if (!string.IsNullOrWhiteSpace(localized))
                Add(localized, bytes, description.Style);
        }
        catch (CultureNotFoundException)
        {
        }
    }

    private static bool IsKnownCjkFamily(string family)
    {
        foreach (var group in CjkAliasGroups)
        {
            foreach (var name in group)
            {
                if (name.Equals(family, StringComparison.OrdinalIgnoreCase)) return true;
            }
        }

        return false;
    }

    private void Add(string family, byte[] data, FontStyle style)
    {
        if (string.IsNullOrWhiteSpace(family)) return;
        if (!_fonts.TryGetValue(family, out var faces)) _fonts[family] = faces = new();
        if (!faces.Any(face => ReferenceEquals(face.Data, data))) faces.Add((data, style));
        _authentic.Add(family);
    }
}
