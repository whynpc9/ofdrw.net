using Ofdrw.Net.Signatures.Crypto;

namespace Ofdrw.Net.Signatures.Signing;

public sealed class OfdSignatureOptions
{
    public string CheckMethod { get; set; } = OfdDigestAlgorithms.Sm3Oid;
    public bool ProtectSignatureList { get; set; }
}
