using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1;
using Ofdrw.Net.Signatures.Signing;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicit SES V5 provider using a caller-owned SM2 signer, certificate and prebuilt seal.</summary>
/// <remarks>The seal must use the full-certificate authorization list. The package does not make or trust seals.</remarks>
public sealed class SesV5ExternalSignatureProvider : IOfdSignatureProvider
{
    private readonly byte[] _signerCertificate;
    private readonly byte[] _sealDer;
    private readonly DateTime _validFrom;
    private readonly DateTime _validUntil;
    private readonly ISesExternalSm2Signer _signer;
    private readonly Func<byte[], CancellationToken, Task<byte[]?>>? _timestampHook;

    /// <summary>Create an explicitly selected V5 provider with caller-owned identity and seal.</summary>
    /// <param name="signerCertificateDer">DER certificate whose public key matches the signer.</param>
    /// <param name="sealDer">Externally issued V5 seal DER; its maker signature is checked here and by the verifier.</param>
    /// <param name="signer">Caller-owned SM2/SM3 signing operation.</param>
    /// <param name="timestampHook">Optional callback receiving raw signature DER and returning a token; no TSA check is implied.</param>
    public SesV5ExternalSignatureProvider(byte[] signerCertificateDer, byte[] sealDer,
        ISesExternalSm2Signer signer,
        Func<byte[], CancellationToken, Task<byte[]?>>? timestampHook = null)
    {
        if (signerCertificateDer is null) throw new ArgumentNullException(nameof(signerCertificateDer));
        if (sealDer is null) throw new ArgumentNullException(nameof(sealDer));
        _signer = signer ?? throw new ArgumentNullException(nameof(signer));
        _ = SesTestCrypto.ReadSm2PublicKey(signerCertificateDer);
        var seal = SesV5Seal.Read(sealDer);
        if (!seal.Authorizes(signerCertificateDer))
            throw new InvalidDataException("V5 seal does not authorize the signer certificate.");
        if (!SesTestCrypto.Verify(seal.SealInfoDer, seal.MakerSignature,
                SesTestCrypto.ReadSm2PublicKey(seal.MakerCertificate)))
            throw new InvalidDataException("V5 seal maker signature is invalid.");
        _signerCertificate = (byte[])signerCertificateDer.Clone();
        _sealDer = (byte[])sealDer.Clone();
        _validFrom = seal.ValidFrom;
        _validUntil = seal.ValidUntil;
        _timestampHook = timestampHook;
    }

    /// <summary>Container provider identifier; this does not assert issuer trust.</summary>
    public string ProviderName => "Ofdrw.Net SES V5 external signer";
    /// <summary>Identity is supplied and assessed by the caller.</summary>
    public string Company => "caller-provided identity";
    /// <summary>Provider profile revision.</summary>
    public string Version => "1";
    /// <summary>SM2-with-SM3 OID.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;
    /// <summary>OFD seal signature record.</summary>
    public string SignatureType => "Seal";
    /// <summary>A defensive copy of the external seal DER.</summary>
    public byte[] SealData => (byte[])_sealDer.Clone();

    /// <summary>Build and sign the V5 TBS over exact Signature.xml bytes and package path.</summary>
    public async Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var time = SesSignatureInput.Validate(signatureXml, propertyInformation, allowJavaLocalTime: true);
        if (time < _validFrom || time > _validUntil)
            throw new InvalidDataException("V5 signature time is outside the seal validity interval.");
        var tbs = new DerSequence(new DerInteger(5), SesDer.Read(_sealDer),
            new DerGeneralizedTime(time), new DerBitString(SesTestProfile.Hash(signatureXml)),
            new DerIA5String(propertyInformation));
        var tbsDer = tbs.GetEncoded("DER");
        var signature = await _signer.SignAsync((byte[])tbsDer.Clone(), cancellationToken).ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        if (signature is null || signature.Length == 0 || signature.Length > 80 ||
            !SesTestCrypto.Verify(tbsDer, signature, SesTestCrypto.ReadSm2PublicKey(_signerCertificate)))
            throw new InvalidDataException("External signer did not return a valid SM2/SM3 signature for its certificate.");
        var elements = new Asn1EncodableVector(tbs, new DerOctetString(_signerCertificate),
            new DerObjectIdentifier(SignatureMethod), new DerBitString(signature));
        if (_timestampHook is not null)
        {
            var token = await _timestampHook((byte[])signature.Clone(), cancellationToken).ConfigureAwait(false);
            cancellationToken.ThrowIfCancellationRequested();
            if (token is { Length: > 0 })
            {
                if (token.Length > 64 * 1024) throw new InvalidDataException("V5 timestamp token exceeds the limit.");
                elements.Add(new DerTaggedObject(true, 0, new DerBitString((byte[])token.Clone())));
            }
        }
        return new DerSequence(elements).GetEncoded("DER");
    }
}
