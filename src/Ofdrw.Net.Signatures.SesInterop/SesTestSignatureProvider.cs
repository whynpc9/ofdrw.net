using System;
using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1;
using Ofdrw.Net.Signatures.Signing;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicit SES V1/V4 SM2 self-signing provider for dev/interop test packages only.</summary>
public sealed class SesTestSignatureProvider : IOfdSignatureProvider
{
    private readonly SesTestIdentity _identity;
    private readonly SesVersion _version;
    private readonly byte[] _seal;
    /// <summary>Create a provider with an ephemeral test identity and an explicit SES layout.</summary>
    public SesTestSignatureProvider(SesTestIdentity identity, SesVersion version)
    {
        _identity = identity ?? throw new ArgumentNullException(nameof(identity));
        SesTestProfile.Version(version); _version = version; _seal = SesTestProfile.Seal(version, identity);
    }
    /// <summary>Clearly marked test provider.</summary>
    public string ProviderName => "Ofdrw.Net SES DEV INTEROP TEST";
    /// <summary>Test-only issuer label.</summary>
    public string Company => "dev/interop; no legal validity";
    /// <summary>Test profile revision.</summary>
    public string Version => "1";
    /// <summary>SM2 with SM3 OID; no default verifier is registered for it.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;
    /// <summary>OFD seal signature record type.</summary>
    public string SignatureType => "Seal";
    /// <summary>A copy of the test seal DER, without a rendered signature appearance.</summary>
    public byte[] SealData => (byte[])_seal.Clone();

    /// <summary>Sign exact XML bytes, UTC time and package path. Cancellation is checked before and after crypto.</summary>
    public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var time = SesTestProfile.ValidateInput(signatureXml, propertyInformation);
        var items = new Asn1EncodableVector(new DerInteger((int)_version), SesDer.Read(_seal),
            SesTestProfile.Time(_version, time), new DerBitString(SesTestProfile.Hash(signatureXml)),
            new DerIA5String(propertyInformation));
        var cert = new DerOctetString(_identity.CertificateDer); var algorithm = new DerObjectIdentifier(SignatureMethod);
        if (_version == SesVersion.V1) items.Add(cert, algorithm);
        var tbs = new DerSequence(items);
        var signature = new DerBitString(SesTestCrypto.Sign(tbs.GetEncoded("DER"), _identity.PrivateKey));
        var result = (_version == SesVersion.V1 ? new DerSequence(tbs, signature)
            : new DerSequence(tbs, cert, algorithm, signature)).GetEncoded("DER");
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(result);
    }
}
