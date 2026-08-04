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
    /// When <c>true</c> (default), PdfPig text extraction is not used as a fallback
    /// because it destroys table/grid layout.
    /// </summary>
    public bool RequireRasterization { get; set; } = true;

    /// <summary>
    /// Gets or sets whether <c>pdftoppm</c> should be tried before the built-in
    /// Docnet/Pdfium rasterizer. Default is <c>false</c> (Docnet first).
    /// </summary>
    public bool PreferExternalPdfToPpm { get; set; }
}
