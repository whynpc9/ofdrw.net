using Ofdrw.Net.Converter.Pdf.Internal;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Coordinates the process-wide PDFsharp font resolver used by Ofdrw.Net converters.
/// </summary>
public static class PdfFontRegistry
{
    /// <summary>
    /// Ensures the shared resolver is installed before a PDF document initializes its font cache.
    /// </summary>
    public static void EnsureInstalled()
    {
        OfdEmbeddedFontResolver.EnsureInstalled();
    }

    /// <summary>
    /// Registers a font face for subsequent PDF rendering.
    /// </summary>
    public static void RegisterFont(
        string familyName,
        byte[] fontData,
        bool bold = false,
        bool italic = false)
    {
        OfdEmbeddedFontResolver.Register(familyName, fontData, bold, italic);
    }
}
