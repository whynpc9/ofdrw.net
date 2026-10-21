using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Security;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicit V5 container verifier with signer and seal-maker certificates pinned out of band.</summary>
/// <remarks>Checks both SM2 signatures and seal authorization, not PKI chain, revocation, TSA or legal effect.</remarks>
public sealed class SesV5PinnedVerifier : IOfdSignedValueVerifier
{
    private readonly byte[] _signerCertificate;
    private readonly byte[] _makerCertificate;
    private readonly ECPublicKeyParameters _signerKey;
    private readonly ECPublicKeyParameters _makerKey;

    /// <summary>Pin both complete DER certificates supplied independently of the document.</summary>
    public SesV5PinnedVerifier(byte[] expectedSignerCertificateDer, byte[] expectedSealMakerCertificateDer)
    {
        if (expectedSignerCertificateDer is null) throw new ArgumentNullException(nameof(expectedSignerCertificateDer));
        if (expectedSealMakerCertificateDer is null) throw new ArgumentNullException(nameof(expectedSealMakerCertificateDer));
        _signerKey = SesTestCrypto.ReadSm2PublicKey(expectedSignerCertificateDer);
        _makerKey = SesTestCrypto.ReadSm2PublicKey(expectedSealMakerCertificateDer);
        _signerCertificate = (byte[])expectedSignerCertificateDer.Clone();
        _makerCertificate = (byte[])expectedSealMakerCertificateDer.Clone();
    }

    /// <summary>SM2-with-SM3 OID. Register only one format-aware verifier for this OID.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;

    /// <summary>Check the narrow V5 full-certificate-list profile against both certificate pins.</summary>
    public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (signatureXml is null) throw new ArgumentNullException(nameof(signatureXml));
        if (signedValue is null) throw new ArgumentNullException(nameof(signedValue));
        if (propertyInformation is null) throw new ArgumentNullException(nameof(propertyInformation));
        bool valid;
        try { valid = Verify(signatureXml, signedValue, propertyInformation); }
        catch (Exception exception) when (exception is InvalidDataException || exception is IOException ||
            exception is ArgumentException || exception is FormatException || exception is ArithmeticException ||
            exception is System.Xml.XmlException || exception is GeneralSecurityException)
        { valid = false; }
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(valid);
    }

    private bool Verify(byte[] xml, byte[] signedValue, string property)
    {
        var xmlTime = SesSignatureInput.Validate(xml, property, allowJavaLocalTime: true,
            out var hasUtcTime);
        var parsed = SesSignedValueReader.Parse(signedValue);
        if (parsed.Version != SesVersion.V5 || parsed.Timestamp is not null ||
            parsed.SignatureAlgorithm != SignatureMethod || parsed.PropertyInformation != property ||
            !SesTestCrypto.Equal(parsed.CertificateDer, _signerCertificate) ||
            !SesTestCrypto.Equal(parsed.DataHash, SesTestProfile.Hash(xml))) return false;
        var tbs = SesDer.Sequence(SesDer.Read(parsed.ToBeSignedDer), 5);
        if (tbs[2].ToAsn1Object() is not Asn1GeneralizedTime signedTimeField) return false;
        var signedTime = signedTimeField.ToDateTime().ToUniversalTime();
        if (hasUtcTime && signedTime != xmlTime) return false;
        var seal = SesV5Seal.Read(parsed.SealDer);
        if (!SesTestCrypto.Equal(seal.MakerCertificate, _makerCertificate) ||
            !seal.Authorizes(_signerCertificate) || signedTime < seal.ValidFrom || signedTime > seal.ValidUntil)
            return false;
        return SesTestCrypto.Verify(seal.SealInfoDer, seal.MakerSignature, _makerKey) &&
            SesTestCrypto.Verify(parsed.ToBeSignedDer, parsed.Signature, _signerKey);
    }
}
