using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal sealed class DocumentFontContext
{
    private readonly IReadOnlyList<OfdFontResource> _fonts;
    private readonly Dictionary<OfdFontResource, string> _families = new();
    internal Dictionary<string, SixLabors.Fonts.FontFamily> OutlineFonts { get; } = new();

    internal DocumentFontContext(IReadOnlyList<OfdFontResource> fonts)
    {
        _fonts = fonts;
        PdfFontRegistry.EnsureInstalled();
        foreach (var font in fonts.Where(font => font.Data.Length > 0))
            _families[font] = PdfFontRegistry.RegisterFontFace(font.Data, font.Bold, font.Italic);
    }

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
