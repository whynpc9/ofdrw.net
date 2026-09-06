using System;

namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Options that control PDF to OFD conversion fidelity.
/// </summary>
public sealed class PdfToOfdOptions
{
    /// <summary>Maximum input bytes, enforced while staging the PDF.</summary>
    public long MaxInputBytes { get; set; } = 128L * 1024 * 1024;

    /// <summary>Maximum source or selected output page count.</summary>
    public int MaxPageCount { get; set; } = 10_000;

    /// <summary>Maximum decoded pixels in any one rendered page.</summary>
    public long MaxRasterizedPixelsPerPage { get; set; } = 40_000_000;

    /// <summary>Maximum accumulated compressed page-image bytes.</summary>
    public long MaxTotalImageBytes { get; set; } = 512L * 1024 * 1024;

    /// <summary>Maximum time allowed for an external pdftoppm attempt.</summary>
    public TimeSpan ExternalRasterizationTimeout { get; set; } = TimeSpan.FromMinutes(2);

    /// <summary>
    /// Gets or sets the target DPI used when rasterizing PDF pages.
    /// Values outside 72–300 are clamped.
    /// </summary>
    public int Dpi { get; set; } = 144;

    /// <summary>
    /// Gets or sets whether conversion must succeed via page rasterization.
    /// When <c>true</c> (default), text extraction is not used as a visual fallback
    /// because it cannot preserve table/grid layout. It may still be used for the
    /// optional machine-readable text layer.
    /// </summary>
    public bool RequireRasterization { get; set; } = true;

    /// <summary>
    /// Gets or sets whether <c>pdftoppm</c> should be tried before the built-in
    /// Docnet/Pdfium rasterizer. Default is <c>false</c> (Docnet first).
    /// </summary>
    public bool PreferExternalPdfToPpm { get; set; }

    /// <summary>
    /// Gets or sets how extractable PDF text is retained in the OFD output.
    /// The default creates a dual-layer document: a rendered page image for visual
    /// fidelity and transparent, positioned OFD text objects for machine use.
    /// </summary>
    public PdfTextLayerMode TextLayerMode { get; set; } = PdfTextLayerMode.Invisible;

    /// <summary>
    /// Gets or sets the maximum number of semantic text objects emitted for one page.
    /// This bounds memory and output growth for untrusted PDFs.
    /// </summary>
    public int MaxTextObjectsPerPage { get; set; } = 50_000;
}
