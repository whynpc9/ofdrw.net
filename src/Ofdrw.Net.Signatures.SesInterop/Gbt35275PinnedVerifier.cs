using System;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Security;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicit single-signer GB/T 35275 container verifier with a certificate pinned out of band.</summary>
/// <remarks>Checks the selected profile and SM2 signature; does not establish PKI or legal trust.</remarks>
public sealed class Gbt35275PinnedVerifier : IOfdSignedValueVerifier
{
    private readonly byte[] _certificate;
    private readonly ECPublicKeyParameters _key;
    private readonly Gbt35275PayloadProfile _profile;

    /// <summary>Pin a complete certificate independently of the signed document.</summary>
    public Gbt35275PinnedVerifier(byte[] expectedSignerCertificateDer,
        Gbt35275PayloadProfile profile = Gbt35275PayloadProfile.RawSm3Digest)
    {
        if (expectedSignerCertificateDer is null) throw new ArgumentNullException(nameof(expectedSignerCertificateDer));
        if (profile != Gbt35275PayloadProfile.RawSm3Digest &&
            profile != Gbt35275PayloadProfile.Base64Sm3DigestCompatibility)
            throw new ArgumentOutOfRangeException(nameof(profile));
        _key = SesTestCrypto.ReadSm2PublicKey(expectedSignerCertificateDer);
        _certificate = (byte[])expectedSignerCertificateDer.Clone();
        _profile = profile;
    }

    /// <summary>SM2-with-SM3 OID. Use a format-aware composite to verify multiple containers under this OID.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;

    /// <summary>Verify attached digest, signer identity, algorithm IDs and signature.</summary>
    public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (signatureXml is null) throw new ArgumentNullException(nameof(signatureXml));
        if (signedValue is null) throw new ArgumentNullException(nameof(signedValue));
        if (propertyInformation is null) throw new ArgumentNullException(nameof(propertyInformation));
        bool valid;
        try
        {
            _ = SesSignatureInput.Validate(signatureXml, propertyInformation, allowJavaLocalTime: true);
            var parsed = Gbt35275SignedData.Read(signedValue);
            var digest = SesTestProfile.Hash(signatureXml);
            var content = _profile == Gbt35275PayloadProfile.RawSm3Digest ? digest
                : Encoding.ASCII.GetBytes(Convert.ToBase64String(digest));
            valid = SesTestCrypto.Equal(parsed.CertificateDer, _certificate) &&
                SesTestCrypto.Equal(parsed.Content, content) &&
                SesTestCrypto.Verify(content, parsed.Signature, _key);
        }
        catch (Exception exception) when (exception is InvalidDataException || exception is IOException ||
            exception is ArgumentException || exception is FormatException || exception is ArithmeticException ||
            exception is System.Xml.XmlException || exception is GeneralSecurityException)
        { valid = false; }
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(valid);
    }
}
