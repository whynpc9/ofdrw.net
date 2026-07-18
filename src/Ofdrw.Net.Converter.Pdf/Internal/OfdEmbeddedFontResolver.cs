using System;
using System.Collections.Generic;
using PdfSharpCore.Fonts;
using PdfSharpCore.Utils;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class OfdEmbeddedFontResolver : IFontResolver
{
    private static readonly object Sync = new();
    private static readonly OfdEmbeddedFontResolver Instance = new();

    private readonly Dictionary<string, string> _faceByFamily =
        new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _fontByFace =
        new(StringComparer.Ordinal);
    private readonly FontResolver _fallback = new();

    private OfdEmbeddedFontResolver()
    {
    }

    public string DefaultFontName => _fallback.DefaultFontName;

    public static void Register(
        string familyName,
        byte[] fontData,
        bool bold,
        bool italic)
    {
        if (string.IsNullOrWhiteSpace(familyName) || fontData.Length == 0)
        {
            return;
        }

        lock (Sync)
        {
            TryInstall();
            if (!ReferenceEquals(GlobalFontSettings.FontResolver, Instance))
            {
                return;
            }

            var faceName = $"ofd:{familyName}:{(bold ? "b" : "r")}:{(italic ? "i" : "n")}";
            Instance._faceByFamily[BuildFamilyKey(familyName, bold, italic)] = faceName;
            Instance._fontByFace[faceName] = fontData;
        }
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        lock (Sync)
        {
            if (_faceByFamily.TryGetValue(
                BuildFamilyKey(familyName, isBold, isItalic),
                out var exactFace))
            {
                return new FontResolverInfo(exactFace);
            }

            if (_faceByFamily.TryGetValue(
                BuildFamilyKey(familyName, false, false),
                out var regularFace))
            {
                return new FontResolverInfo(regularFace, isBold, isItalic);
            }
        }

        return _fallback.ResolveTypeface(familyName, isBold, isItalic);
    }

    public byte[] GetFont(string faceName)
    {
        lock (Sync)
        {
            if (_fontByFace.TryGetValue(faceName, out var fontData))
            {
                return fontData;
            }
        }

        return _fallback.GetFont(faceName);
    }

    public static void EnsureInstalled()
    {
        lock (Sync)
        {
            TryInstall();
        }
    }

    private static void TryInstall()
    {
        if (GlobalFontSettings.FontResolver is not null)
        {
            return;
        }

        try
        {
            GlobalFontSettings.FontResolver = Instance;
        }
        catch (InvalidOperationException)
        {
            // A host may have initialized PDFsharp's global font cache already.
            // In that case its resolver remains authoritative.
        }
    }

    private static string BuildFamilyKey(string familyName, bool bold, bool italic)
    {
        return $"{familyName}\u001f{bold}\u001f{italic}";
    }
}
