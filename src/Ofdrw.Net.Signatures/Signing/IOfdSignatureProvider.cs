using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Signatures.Signing;

public interface IOfdSignatureProvider
{
    string ProviderName { get; }
    string Company { get; }
    string Version { get; }
    string SignatureMethod { get; }
    string SignatureType { get; }
    byte[]? SealData { get; }

    Task<byte[]> SignAsync(
        byte[] signatureXml,
        string propertyInformation,
        CancellationToken cancellationToken = default);
}
