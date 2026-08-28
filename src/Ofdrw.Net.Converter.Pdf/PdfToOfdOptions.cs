namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Options that control PDF to OFD conversion fidelity.
/// </summary>
public sealed class PdfToOfdOptions
{
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
