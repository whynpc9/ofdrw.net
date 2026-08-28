using System;
using System.Collections.Generic;

namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Configures the process used for DOCX rendering.
/// </summary>
public sealed class DocxConversionOptions
{
    /// <summary>
    /// Gets or sets the rendering engine used for PDF and dual-layer OFD output.
    /// Native OFD output does not invoke a rendering engine. Auto prefers Microsoft Word
    /// on macOS when installed, then LibreOffice, and finally the in-process renderer.
    /// </summary>
    public DocxConversionEngine Engine { get; set; } = DocxConversionEngine.Auto;

    /// <summary>
    /// Gets or sets the DOCX-to-OFD materialization mode. The default writes native
    /// OFD text objects directly from OpenXML and skips the PDF rendering stage.
    /// </summary>
    public DocxToOfdMode OfdMode { get; set; } = DocxToOfdMode.Native;

    /// <summary>
    /// Gets or sets the LibreOffice executable path. When omitted, the converter checks
    /// <c>OFDRW_LIBREOFFICE_PATH</c>, common installation locations, and then <c>soffice</c> on PATH.
    /// </summary>
    public string? LibreOfficePath { get; set; }

    /// <summary>
    /// Gets or sets the maximum time allowed for one conversion attempt.
    /// </summary>
    public TimeSpan ProcessTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Gets additional font directories for DOCX rendering. Font files are linked or copied
    /// into an isolated LibreOffice profile and registered with the BuiltIn PDF renderer.
    /// </summary>
    public IList<string> FontDirectories { get; } = new List<string>();

    /// <summary>
    /// Gets or sets whether macOS Microsoft Word private fonts should be made available
    /// when Word is installed. The fonts are referenced in place and are never bundled.
    /// </summary>
    public bool UseInstalledMicrosoftOfficeFonts { get; set; } = true;

    /// <summary>
    /// Gets or sets whether each LibreOffice conversion should use an isolated
    /// <c>-env:UserInstallation</c> profile. When <c>null</c>, isolation is enabled for
    /// regular installs and disabled automatically for LibreOffice Portable /
    /// SecureUserConfig builds (those hang on a fresh empty profile). When isolation is
    /// disabled, conversions are serialized through an internal gate.
    /// </summary>
    public bool? IsolateUserProfile { get; set; }

    /// <summary>
    /// Gets or sets how the in-process renderer handles DOCX features it cannot render faithfully.
    /// </summary>
    public UnsupportedDocxFeatureBehavior UnsupportedFeatureBehavior { get; set; } =
        UnsupportedDocxFeatureBehavior.Placeholder;

    /// <summary>
    /// Gets or sets the maximum accepted DOCX input size in bytes.
    /// </summary>
    public long MaxInputBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>
    /// Gets or sets the maximum total uncompressed size of all package parts.
    /// </summary>
    public long MaxExpandedBytes { get; set; } = 256L * 1024 * 1024;

    /// <summary>
    /// Gets or sets the maximum number of entries in the DOCX package.
    /// </summary>
    public int MaxPackagePartCount { get; set; } = 4096;

    /// <summary>
    /// Gets or sets the maximum number of document XML elements processed by BuiltIn.
    /// </summary>
    public int MaxDocumentElements { get; set; } = 1_000_000;

    /// <summary>
    /// Gets or sets the maximum compressed size of one embedded image in bytes.
    /// </summary>
    public long MaxEmbeddedImageBytes { get; set; } = 32L * 1024 * 1024;

    /// <summary>
    /// Gets or sets the maximum decoded pixel count of one embedded image.
    /// </summary>
    public long MaxEmbeddedImagePixels { get; set; } = 40_000_000;

    /// <summary>
    /// Gets font family names considered when DOCX styles do not resolve an explicit font.
    /// </summary>
    public IList<string> FontFallbackFamilies { get; } = new List<string>
    {
        "Noto Sans CJK SC",
        "Microsoft YaHei",
        "PingFang SC",
        "SimSun",
        "Arial",
        "DejaVu Sans"
    };
}
