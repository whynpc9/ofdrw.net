namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Selects the rendering engine used for DOCX to PDF conversion.
/// </summary>
public enum DocxConversionEngine
{
    /// <summary>
    /// Prefer Microsoft Word on macOS when it is installed, otherwise use LibreOffice.
    /// </summary>
    Auto,

    /// <summary>
    /// Use Microsoft Word for macOS through its AppleScript interface.
    /// </summary>
    MicrosoftWord,

    /// <summary>
    /// Use LibreOffice in headless mode.
    /// </summary>
    LibreOffice
}
