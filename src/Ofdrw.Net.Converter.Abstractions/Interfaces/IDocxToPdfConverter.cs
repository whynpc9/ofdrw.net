using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Abstractions.Interfaces;

/// <summary>
/// Converts a DOCX document to PDF.
/// </summary>
public interface IDocxToPdfConverter
{
    /// <summary>
    /// Converts the DOCX data in <paramref name="docxInput"/> to PDF.
    /// </summary>
    Task ConvertAsync(
        Stream docxInput,
        Stream pdfOutput,
        CancellationToken cancellationToken = default);
}
