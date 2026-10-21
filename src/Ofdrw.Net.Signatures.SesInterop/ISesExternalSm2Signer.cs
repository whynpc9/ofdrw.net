using System.Threading;
using System.Threading.Tasks;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Caller-owned SM2/SM3 signing operation. The key remains outside this package.</summary>
/// <remarks>Sign the exact supplied bytes with SM2 ID 1234567812345678 and return canonical DER SEQUENCE { r, s }.</remarks>
public interface ISesExternalSm2Signer
{
    /// <summary>Sign the supplied immutable snapshot. No key material is passed to this package.</summary>
    Task<byte[]> SignAsync(byte[] data, CancellationToken cancellationToken = default);
}
