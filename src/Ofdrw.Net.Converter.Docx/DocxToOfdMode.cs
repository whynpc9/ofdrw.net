namespace Ofdrw.Net.Converter.Docx;

/// <summary>
/// Selects how DOCX content is materialized as OFD.
/// </summary>
public enum DocxToOfdMode
{
    /// <summary>
    /// Converts DOCX/OpenXML text directly to native OFD text objects without a PDF stage.
    /// </summary>
    Native,

    /// <summary>
    /// Uses a rendered PDF page as the visual layer and DOCX/OpenXML as the text layer.
    /// </summary>
    DualLayer
}
