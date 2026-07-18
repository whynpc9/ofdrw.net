using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Signatures.Verification;

public interface IOfdSignedValueVerifier
{
    string SignatureMethod { get; }

    Task<bool> VerifyAsync(
        byte[] signatureXml,
        byte[] signedValue,
        string propertyInformation,
        CancellationToken cancellationToken = default);
}
