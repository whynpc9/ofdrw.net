using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Converter.Abstractions.Interfaces;

public interface IOfdToPdfConverter
{
    Task ConvertAsync(Stream ofdInput, Stream pdfOutput, IReadOnlyList<int>? pages = null, CancellationToken cancellationToken = default);
}
