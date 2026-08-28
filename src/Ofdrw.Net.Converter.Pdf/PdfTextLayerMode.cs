namespace Ofdrw.Net.Converter.Pdf;

/// <summary>
/// Controls whether PDF text is retained as machine-readable OFD text objects.
/// </summary>
public enum PdfTextLayerMode
{
    /// <summary>
    /// Writes only the rendered page image.
    /// </summary>
    None,

    /// <summary>
    /// Writes positioned, transparent text objects above the rendered page image.
    /// The image remains the visual source of truth while the text remains searchable
    /// and extractable.
    /// </summary>
    Invisible
}
