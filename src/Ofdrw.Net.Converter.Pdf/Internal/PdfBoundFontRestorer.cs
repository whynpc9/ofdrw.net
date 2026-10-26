using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Ofdrw.Net.Core.Fonts;
using PdfSharpCore.Pdf;
using PdfSharpCore.Pdf.IO;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// PDFsharpCore 1.3.67 does not close recursively over nested TrueType composite
/// components when it subsets FontFile2. Keep the verified, identity-renamed
/// complete face for selected, bounded font profiles, including direct-PUA
/// symbol aliases. This also honors fonts that prohibit subsetting.
/// CIDToGIDMap and ToUnicode remain the ones emitted for the selected text.
/// </summary>
internal static class PdfBoundFontRestorer
{
    internal static byte[] Restore(byte[] pdf, IReadOnlyDictionary<string, byte[]> faces)
    {
        if (faces.Count == 0) return pdf;

        using var source = new MemoryStream(pdf, writable: false);
        using var parsed = PdfReader.Open(source, PdfDocumentOpenMode.Modify);
        var fullFaces = faces.ToDictionary(pair => pair.Key,
            pair => OpenTypeFontIdentity.WithUniqueNames(pair.Value, pair.Key), StringComparer.Ordinal);
        var restored = new HashSet<string>(StringComparer.Ordinal);
        foreach (var descriptor in parsed.Internals.GetAllObjects().OfType<PdfDictionary>()
                     .Where(value => value.Elements.ContainsKey("/FontFile2")))
        {
            var fontName = descriptor.Elements.GetName("/FontName");
            if (string.IsNullOrEmpty(fontName)) continue;
            var identity = fullFaces.Keys.FirstOrDefault(key =>
                fontName.IndexOf("Ofdrw-" + key, StringComparison.OrdinalIgnoreCase) >= 0);
            if (identity is null) continue;

            var fontFile = descriptor.Elements.GetDictionary("/FontFile2");
            if (fontFile?.Stream is null)
                throw new InvalidDataException("Selected verified PDF font has no embedded FontFile2 stream.");
            fontFile.Elements.Remove("/Filter");
            fontFile.Elements.Remove("/DecodeParms");
            fontFile.Stream.Value = fullFaces[identity];
            fontFile.Elements.SetInteger("/Length1", fullFaces[identity].Length);
            fontFile.Stream.Zip();
            restored.Add(identity);
        }
        if (restored.Count != fullFaces.Count)
            throw new InvalidDataException("Selected verified PDF face was not found among the embedded font descriptors.");

        using var output = new MemoryStream();
        parsed.Save(output, false);
        return output.ToArray();
    }
}
