using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Signatures.Signing;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicitly selected payload encodings for the narrow GB/T 35275 container profile.</summary>
public enum Gbt35275PayloadProfile
{
    /// <summary>Sign and attach the raw 32-byte SM3 digest of Signature.xml.</summary>
    RawSm3Digest,
    /// <summary>Compatibility with producers that sign ASCII Base64 of the SM3 digest.</summary>
    Base64Sm3DigestCompatibility
}

/// <summary>Single-signer GB/T 35275 ContentInfo/SignedData provider using a caller-owned SM2 signer.</summary>
/// <remarks>No certificate-chain, revocation or trust service is provided.</remarks>
public sealed class Gbt35275SignatureProvider : IOfdSignatureProvider
{
    private readonly byte[] _certificate;
    private readonly ISesExternalSm2Signer _signer;
    private readonly Gbt35275PayloadProfile _profile;

    /// <summary>Create a provider with an out-of-band signer and an explicit content profile.</summary>
    public Gbt35275SignatureProvider(byte[] signerCertificateDer, ISesExternalSm2Signer signer,
        Gbt35275PayloadProfile profile = Gbt35275PayloadProfile.RawSm3Digest)
    {
        if (signerCertificateDer is null) throw new ArgumentNullException(nameof(signerCertificateDer));
        _signer = signer ?? throw new ArgumentNullException(nameof(signer));
        if (profile != Gbt35275PayloadProfile.RawSm3Digest &&
            profile != Gbt35275PayloadProfile.Base64Sm3DigestCompatibility)
            throw new ArgumentOutOfRangeException(nameof(profile));
        _ = SesTestCrypto.ReadSm2PublicKey(signerCertificateDer);
        _certificate = (byte[])signerCertificateDer.Clone();
        _profile = profile;
    }

    /// <summary>Provider name, without a trust assertion.</summary>
    public string ProviderName => "Ofdrw.Net GB/T 35275 external signer";
    /// <summary>Identity comes from caller-owned certificate and policy.</summary>
    public string Company => "caller-provided identity";
    /// <summary>Provider profile revision.</summary>
    public string Version => "1";
    /// <summary>SM2-with-SM3 OID used in Signature.xml.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;
    /// <summary>OFD digital-signature record type.</summary>
    public string SignatureType => "Sign";
    /// <summary>A digital signature has no electronic seal.</summary>
    public byte[]? SealData => null;

    /// <summary>Sign the selected SM3 digest representation and wrap it in attached SignedData.</summary>
    public async Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _ = SesSignatureInput.Validate(signatureXml, propertyInformation, allowJavaLocalTime: true);
        var digest = SesTestProfile.Hash(signatureXml);
        var content = _profile == Gbt35275PayloadProfile.RawSm3Digest ? digest
            : Encoding.ASCII.GetBytes(Convert.ToBase64String(digest));
        var signature = await _signer.SignAsync((byte[])content.Clone(), cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (signature is null || signature.Length == 0 || signature.Length > 80 ||
            !SesTestCrypto.Verify(content, signature, SesTestCrypto.ReadSm2PublicKey(_certificate)))
            throw new InvalidDataException("External signer did not return a valid SM2/SM3 signature for its certificate.");
        return Gbt35275SignedData.Encode(content, signature, _certificate);
    }
}
