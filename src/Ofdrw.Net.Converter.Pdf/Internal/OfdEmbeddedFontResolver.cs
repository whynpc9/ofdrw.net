using System;
using System.Collections.Generic;
using System.IO;
using Ofdrw.Net.Core.IO;
using SixLabors.Fonts;
using PdfSharpCore.Fonts;
using PdfSharpCore.Utils;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class OfdEmbeddedFontResolver : IFontResolver
{
    private static readonly object InstallSync = new();
    private static readonly object Sync = new();
    private static readonly OfdEmbeddedFontResolver Instance = new();
    private readonly Dictionary<string, FontResolverInfo> _faceByFamily = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, byte[]> _fontByFace = new(StringComparer.Ordinal);
    private readonly Lazy<IFontResolver> _fallback = new(() => new FontResolver());
    private long _registeredBytes;
    private static long _maximumBytes = 256L * 1024 * 1024;
    private static int _maximumFaces = 4096;

    private OfdEmbeddedFontResolver() { }
    public string DefaultFontName => _fallback.Value.DefaultFontName;

    internal static long MaximumBytes
    {
        get { lock (Sync) return _maximumBytes; }
        set
        {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            lock (Sync)
            {
                if (value < Instance._registeredBytes) throw new InvalidOperationException("The font budget cannot be set below already registered data.");
                _maximumBytes = value;
            }
        }
    }

    internal static int MaximumFaces
    {
        get { lock (Sync) return _maximumFaces; }
        set
        {
            if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value));
            lock (Sync)
            {
                if (value < Instance._faceByFamily.Count) throw new InvalidOperationException("The face budget cannot be set below already registered faces.");
                _maximumFaces = value;
            }
        }
    }

    internal static long RegisteredBytes { get { lock (Sync) return Instance._registeredBytes; } }

    // Legacy named registration is immutable. Converters use RegisterFace and
    // retain the returned content identity in their own document-local mappings.
    public static void Register(string familyName, byte[] fontData, bool bold, bool italic)
    {
        if (string.IsNullOrWhiteSpace(familyName)) throw new ArgumentException("A font family name is required.", nameof(familyName));
        var family = RegisterFace(fontData, bold, italic);
        lock (Sync)
        {
            var info = Instance._faceByFamily[BuildFamilyKey(family, bold, italic)];
            var key = BuildFamilyKey(familyName, bold, italic);
            if (Instance._faceByFamily.TryGetValue(key, out var existing) && existing.FaceName != info.FaceName)
                throw new InvalidOperationException("A different font is already registered under this name. Use RegisterFontFace and its returned family for document-local fonts.");
            AddFace(key, info);
        }
    }

    internal static string RegisterFace(byte[] fontData, bool bold, bool italic)
    {
        if (fontData is null) throw new ArgumentNullException(nameof(fontData));
        if (fontData.Length == 0) throw new ArgumentException("Font data must not be empty.", nameof(fontData));
        EnsureInstalled();
        if (!ReferenceEquals(GlobalFontSettings.FontResolver, Instance) && GlobalFontSettings.FontResolver is not DelegatingResolver)
            throw new InvalidOperationException("Embedded fonts require PdfFontRegistry.EnsureInstalled() before the host creates its first PDFsharp font. The host's existing resolver was preserved.");
        if (fontData.Length > MaximumBytes) throw new InvalidOperationException("Font exceeds the configured process-wide byte budget.");
        var snapshot = (byte[])fontData.Clone();
        var identity = BinaryIdentity.Hash(snapshot);
        var family = $"ofd-font-{identity}-{(bold ? 'b' : 'r')}{(italic ? 'i' : 'n')}";
        var faceName = "ofd:" + identity;
        var key = BuildFamilyKey(family, bold, italic);
        using var stream = new MemoryStream(snapshot, writable: false);
        var actualStyle = FontDescription.LoadDescription(stream).Style;
        lock (Sync)
        {
            if (Instance._faceByFamily.ContainsKey(key)) return family;
            if (Instance._faceByFamily.Count >= _maximumFaces)
                throw new InvalidOperationException("The configured process-wide font face budget has been reached.");
            if (!Instance._fontByFace.ContainsKey(faceName))
            {
                if (fontData.Length > _maximumBytes - Instance._registeredBytes)
                    throw new InvalidOperationException("The configured process-wide embedded font byte budget has been reached.");
                Instance._fontByFace.Add(faceName, snapshot);
                Instance._registeredBytes += fontData.Length;
            }
            AddFace(key, new FontResolverInfo(faceName,
                bold && (actualStyle & FontStyle.Bold) == 0,
                italic && (actualStyle & FontStyle.Italic) == 0));
        }
        return family;
    }

    private static void AddFace(string key, FontResolverInfo info)
    {
        if (!Instance._faceByFamily.ContainsKey(key) && Instance._faceByFamily.Count >= _maximumFaces)
            throw new InvalidOperationException("The configured process-wide font face budget has been reached.");
        Instance._faceByFamily[key] = info;
    }

    public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        lock (Sync)
        {
            if (_faceByFamily.TryGetValue(BuildFamilyKey(familyName, isBold, isItalic), out var exact)) return exact;
            if (_faceByFamily.TryGetValue(BuildFamilyKey(familyName, false, false), out var regular))
                return new FontResolverInfo(regular.FaceName, isBold, isItalic);
        }
        return _fallback.Value.ResolveTypeface(familyName, isBold, isItalic);
    }

    public byte[] GetFont(string faceName)
    {
        lock (Sync)
        {
            if (_fontByFace.TryGetValue(faceName, out var bytes))
                return OpenTypeFontIdentity.WithUniqueNames(bytes, faceName.Substring("ofd:".Length));
        }
        return _fallback.Value.GetFont(faceName);
    }

    internal static byte[] GetOriginalFont(string faceName)
    {
        lock (Sync)
        {
            if (Instance._fontByFace.TryGetValue(faceName, out var bytes)) return (byte[])bytes.Clone();
        }
        return GlobalFontSettings.FontResolver.GetFont(faceName);
    }

    public static void EnsureInstalled()
    {
        // Never hold the resolver dictionary lock while entering PDFsharp's
        // font-factory lock; ResolveTypeface is called in the reverse direction.
        lock (InstallSync)
        {
            var existing = GlobalFontSettings.FontResolver;
            if (ReferenceEquals(existing, Instance) || existing.GetType() != typeof(FontResolver)) return;
            try { GlobalFontSettings.FontResolver = Instance; }
            catch (InvalidOperationException) { /* Preserve an already initialized host resolver. */ }
        }
    }

    internal static IFontResolver CreateResolver(IFontResolver fallback)
    {
        return new DelegatingResolver(fallback ?? throw new ArgumentNullException(nameof(fallback)));
    }

    private sealed class DelegatingResolver : IFontResolver
    {
        private readonly IFontResolver _host;
        internal DelegatingResolver(IFontResolver host) => _host = host;
        public string DefaultFontName => _host.DefaultFontName;
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            lock (Sync)
            {
                if (Instance._faceByFamily.TryGetValue(BuildFamilyKey(familyName, isBold, isItalic), out var exact)) return exact;
                if (Instance._faceByFamily.TryGetValue(BuildFamilyKey(familyName, false, false), out var regular))
                    return new FontResolverInfo(regular.FaceName, isBold, isItalic);
            }
            return _host.ResolveTypeface(familyName, isBold, isItalic);
        }
        public byte[] GetFont(string faceName)
        {
            lock (Sync)
            {
                if (Instance._fontByFace.ContainsKey(faceName)) return Instance.GetFont(faceName);
            }
            return _host.GetFont(faceName);
        }
    }

    private static string BuildFamilyKey(string familyName, bool bold, bool italic) => $"{familyName}\u001f{bold}\u001f{italic}";
}
