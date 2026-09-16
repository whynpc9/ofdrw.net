using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class DocumentFontContext
{
    private readonly IReadOnlyList<OfdFontResource> _fonts;
    private readonly Dictionary<OfdFontResource, string> _families = new();
    internal Dictionary<string, SixLabors.Fonts.FontFamily> OutlineFonts { get; } = new();

    internal DocumentFontContext(IReadOnlyList<OfdFontResource> fonts, IFontResolver? fallbackResolver = null)
    {
        _fonts = fonts;
        PdfFontRegistry.EnsureInstalled();
        fallbackResolver ??= GlobalFontSettings.FontResolver;
        foreach (var font in fonts)
        {
            if (font.Data.Length > 0)
            {
                _families[font] = PdfFontRegistry.RegisterFontFace(font.Data, font.Bold, font.Italic);
                continue;
            }

            // Name-only discovery is optional. Unusable host fonts must not make
            // an otherwise renderable document fail before per-element fallback.
            try
            {
                var local = CjkViewerFontLoader.TryRead(font.FontName) ?? CjkViewerFontLoader.TryRead(font.FamilyName);
                if (local is null && (font.Bold || font.Italic))
                {
                    // Only probe requested styles: plain name-only text can use the
                    // host directly without copying its fonts into our registry.
                    var face = fallbackResolver.ResolveTypeface(font.FontName, font.Bold, font.Italic);
                    if (face is not null) local = fallbackResolver.GetFont(face.FaceName);
                }
                if (IsStandaloneFont(local))
                    _families[font] = PdfFontRegistry.RegisterFontFace(local!, font.Bold, font.Italic);
            }
            catch (Exception exception) when (exception is not OutOfMemoryException &&
                                               exception is not OperationCanceledException)
            {
                // Includes malformed font parsing, host I/O, and optional registry
                // budget exhaustion. Embedded OFD fonts above remain strict.
            }
        }
    }

    private static bool IsStandaloneFont(byte[]? data) => data is { Length: >= 12 } &&
        ((data[0] == 0 && data[1] == 1 && data[2] == 0 && data[3] == 0) ||
         (data[0] == 'O' && data[1] == 'T' && data[2] == 'T' && data[3] == 'O'));

    internal string Resolve(OfdTextElement text, out OfdFontResource? resource)
    {
        resource = !string.IsNullOrEmpty(text.FontResourceId)
            ? _fonts.FirstOrDefault(font => font.Id == text.FontResourceId) : null;
        resource ??= _fonts.FirstOrDefault(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase)
                && !font.Bold && !font.Italic)
            ?? _fonts.FirstOrDefault(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase));
        return resource is not null && _families.TryGetValue(resource, out var family)
            ? family : string.IsNullOrWhiteSpace(text.FontName) ? "Arial" : text.FontName;
    }
}
