namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Selects the rendering engine used for DOCX to PDF conversion.
/// </summary>
public enum DocxConversionEngine
{
    /// <summary>
    /// Prefer Microsoft Word on macOS when it is installed, then LibreOffice,
    /// and finally the in-process renderer.
    /// </summary>
    Auto,

    /// <summary>
    /// Use Microsoft Word for macOS through its AppleScript interface.
    /// </summary>
    MicrosoftWord,

    /// <summary>
    /// Use LibreOffice in headless mode.
    /// </summary>
    LibreOffice,

    /// <summary>
    /// Use the cross-platform in-process Open XML renderer. This engine is a
    /// predictable fallback and does not promise pixel parity with Microsoft Word.
    /// </summary>
    BuiltIn
}
