using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Abstractions.Interfaces;

public interface IPdfToOfdConverter
{
    Task ConvertAsync(Stream pdfInput, Stream ofdOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default);
}
