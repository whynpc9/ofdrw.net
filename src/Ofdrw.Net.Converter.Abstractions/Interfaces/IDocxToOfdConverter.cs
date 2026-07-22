using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Abstractions.Interfaces;

/// <summary>
/// Converts a DOCX document to OFD through a PDF rendering stage.
/// </summary>
public interface IDocxToOfdConverter
{
    /// <summary>
    /// Converts the DOCX data in <paramref name="docxInput"/> to OFD.
    /// </summary>
    Task ConvertAsync(
        Stream docxInput,
        Stream ofdOutput,
        IReadOnlyList<int>? pages = null,
        CancellationToken cancellationToken = default);
}
