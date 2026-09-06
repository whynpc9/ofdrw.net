using Ofdrw.Net.Converter.Pdf.Internal;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Coordinates the process-wide PDFsharp font resolver used by Ofdrw.Net converters.
/// </summary>
public static class PdfFontRegistry
{
    /// <summary>
    /// Composes embedded-font support with a host resolver. Assign the result to
    /// GlobalFontSettings.FontResolver at startup, before creating any XFont.
    /// </summary>
    public static IFontResolver CreateResolver(IFontResolver fallback) => OfdEmbeddedFontResolver.CreateResolver(fallback);

    /// <summary>Maximum font payload bytes retained by the process-wide registry. Set before conversion.</summary>
    public static long MaximumRegisteredFontBytes
    {
        get => OfdEmbeddedFontResolver.MaximumBytes;
        set => OfdEmbeddedFontResolver.MaximumBytes = value;
    }

    /// <summary>Maximum registered face aliases; existing PDFsharp cache entries are never evicted while in use.</summary>
    public static int MaximumRegisteredFaces
    {
        get => OfdEmbeddedFontResolver.MaximumFaces;
        set => OfdEmbeddedFontResolver.MaximumFaces = value;
    }

    /// <summary>Current retained payload size; identical font bytes are counted once.</summary>
    public static long RegisteredFontBytes => OfdEmbeddedFontResolver.RegisteredBytes;

    /// <summary>Registers immutable bytes and returns a content-specific family to use for this document.</summary>
    public static string RegisterFontFace(byte[] fontData, bool bold = false, bool italic = false)
    {
        return OfdEmbeddedFontResolver.RegisterFace(fontData, bold, italic);
    }

    internal static byte[] GetOriginalFont(string faceName) => OfdEmbeddedFontResolver.GetOriginalFont(faceName);
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
