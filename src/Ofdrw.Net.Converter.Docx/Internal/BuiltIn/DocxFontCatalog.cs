using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Ofdrw.Net.Converter.Pdf;
using SixLabors.Fonts;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

/// <summary>Per-conversion mappings from configured font families to immutable content identities.</summary>
internal sealed class DocxFontCatalog
{
    private readonly Dictionary<string, List<(byte[] Data, FontStyle Style)>> _fonts = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, string> _resolved = new(StringComparer.OrdinalIgnoreCase);

    internal DocxFontCatalog(DocxConversionOptions options, IList<DocxConversionDiagnostic> diagnostics, CancellationToken cancellationToken)
    {
        PdfFontRegistry.EnsureInstalled();
        foreach (var directory in options.FontDirectories)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (string.IsNullOrWhiteSpace(directory) || !Directory.Exists(directory))
            {
                diagnostics.Add(new DocxConversionDiagnostic("DOCX_FONT_DIRECTORY_MISSING", "A configured font directory does not exist."));
                continue;
            }
            foreach (var path in Directory.EnumerateFiles(directory).OrderBy(path => path, StringComparer.Ordinal))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var extension = Path.GetExtension(path);
                if (!extension.Equals(".ttf", StringComparison.OrdinalIgnoreCase) && !extension.Equals(".otf", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    if (new FileInfo(path).Length > options.MaxEmbeddedFontBytes)
                        throw new InvalidDataException("Configured font exceeds MaxEmbeddedFontBytes.");
                    var bytes = File.ReadAllBytes(path);
                    using var stream = new MemoryStream(bytes, writable: false);
                    var description = FontDescription.LoadDescription(stream);
                    Add(Path.GetFileNameWithoutExtension(path), bytes, description.Style);
                    Add(description.FontFamilyInvariantCulture, bytes, description.Style);
                }
                catch (Exception exception) when (exception is IOException or UnauthorizedAccessException)
                {
                    diagnostics.Add(new DocxConversionDiagnostic("DOCX_FONT_LOAD_FAILED", $"Configured font '{Path.GetFileName(path)}' could not be loaded: {exception.Message}"));
                }
            }
        }
    }

    internal string Resolve(string family, bool bold, bool italic)
    {
        var key = family + $"\u001f{bold}\u001f{italic}";
        if (_resolved.TryGetValue(key, out var resolved)) return resolved;
        if (!_fonts.TryGetValue(family, out var faces)) return family;
        var style = (bold ? FontStyle.Bold : FontStyle.Regular) | (italic ? FontStyle.Italic : FontStyle.Regular);
        var face = faces.FirstOrDefault(candidate => candidate.Style == style);
        if (face.Data is null) face = faces.FirstOrDefault(candidate => candidate.Style == FontStyle.Regular);
        if (face.Data is null) face = faces[0];
        resolved = PdfFontRegistry.RegisterFontFace(face.Data, bold, italic);
        _resolved.Add(key, resolved);
        return resolved;
    }

    private void Add(string family, byte[] data, FontStyle style)
    {
        if (!_fonts.TryGetValue(family, out var faces)) _fonts[family] = faces = new();
        if (!faces.Any(face => ReferenceEquals(face.Data, data))) faces.Add((data, style));
    }
}
