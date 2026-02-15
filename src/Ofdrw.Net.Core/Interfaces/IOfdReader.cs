using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Interfaces;

public interface IOfdReader
{
    Task<OfdDocumentPackage> ReadAsync(Stream ofdStream, CancellationToken cancellationToken = default);
}
