using System;
using System.Collections.Generic;

namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Configures the process used for DOCX rendering.
/// </summary>
public sealed class DocxConversionOptions
{
    /// <summary>
    /// Gets or sets the rendering engine. Auto prefers Microsoft Word on macOS when installed,
    /// then falls back to LibreOffice.
    /// </summary>
    public DocxConversionEngine Engine { get; set; } = DocxConversionEngine.Auto;

    /// <summary>
    /// Gets or sets the LibreOffice executable path. When omitted, the converter checks
    /// <c>OFDRW_LIBREOFFICE_PATH</c>, common installation locations, and then <c>soffice</c> on PATH.
    /// </summary>
    public string? LibreOfficePath { get; set; }

    /// <summary>
    /// Gets or sets the maximum time allowed for one conversion process.
    /// </summary>
    public TimeSpan ProcessTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Gets additional font directories to expose to the isolated LibreOffice profile.
    /// Font files are linked on Unix-like systems and copied on Windows.
    /// </summary>
    public IList<string> FontDirectories { get; } = new List<string>();

    /// <summary>
    /// Gets or sets whether macOS Microsoft Word private fonts should be made available
    /// when Word is installed. The fonts are referenced in place and are never bundled.
    /// </summary>
    public bool UseInstalledMicrosoftOfficeFonts { get; set; } = true;
}
