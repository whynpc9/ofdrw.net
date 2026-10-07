using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Security;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Explicit, exact-certificate-pinned verification of this package's dev/interop self-signing profile.</summary>
/// <remarks>A successful result means this test profile verified, with no reader recognition, legal effect,
/// certificate-chain/revocation validation, signature appearance validation or qualified TSA claim.</remarks>
public sealed class SesTestSignedValueVerifier : IOfdSignedValueVerifier
{
    private readonly byte[] _certificate;
    private readonly ECPublicKeyParameters _key;
    /// <summary>Pin the complete test certificate out of band. The container cannot choose its trust anchor.</summary>
    public SesTestSignedValueVerifier(byte[] expectedTestCertificate)
    {
        if (expectedTestCertificate is null) throw new ArgumentNullException(nameof(expectedTestCertificate));
        _certificate = (byte[])expectedTestCertificate.Clone();
        _key = SesTestCrypto.ReadPinnedCertificate(_certificate);
    }
    /// <summary>SM2 with SM3 OID, matched only after explicit registration in OfdSignatureVerifier.</summary>
    public string SignatureMethod => SesTestCrypto.Algorithm;

    /// <summary>Verify both seal and document signatures plus exact XML digest, path, time, algorithm and certificate bindings.</summary>
    /// <remarks>Malformed/unsupported inputs return false. Null programmer inputs throw. Cancellation propagates.</remarks>
    public Task<bool> VerifyAsync(byte[] signatureXml, byte[] signedValue, string propertyInformation,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (signatureXml is null) throw new ArgumentNullException(nameof(signatureXml));
        if (signedValue is null) throw new ArgumentNullException(nameof(signedValue));
        if (propertyInformation is null) throw new ArgumentNullException(nameof(propertyInformation));
        bool valid;
        try { valid = Verify(signatureXml, signedValue, propertyInformation); }
        catch (Exception exception) when (exception is InvalidDataException || exception is IOException || exception is ArgumentException ||
            exception is FormatException || exception is ArithmeticException || exception is XmlException ||
            exception is GeneralSecurityException)
        { valid = false; }
        cancellationToken.ThrowIfCancellationRequested();
        return Task.FromResult(valid);
    }

    private bool Verify(byte[] xml, byte[] signedValue, string property)
    {
        var time = SesTestProfile.ValidateInput(xml, property);
        var parsed = SesSignedValueReader.Parse(signedValue);
        if (parsed.Timestamp is not null || parsed.SignatureAlgorithm != SignatureMethod ||
            !SesTestCrypto.Equal(parsed.CertificateDer, _certificate) || parsed.PropertyInformation != property ||
            !SesTestCrypto.Equal(parsed.DataHash, SesTestProfile.Hash(xml))) return false;
        var tbs = SesDer.Sequence(SesDer.Read(parsed.ToBeSignedDer), parsed.Version == SesVersion.V1 ? 7 : 5);
        if (!SesTestCrypto.Equal(tbs[2].GetEncoded("DER"), SesTestProfile.Time(parsed.Version, time).GetEncoded("DER"))) return false;
        var seal = SesDer.Sequence(SesDer.Read(parsed.SealDer), parsed.Version == SesVersion.V1 ? 2 : 4);
        if (!SesTestCrypto.Equal(seal[0].GetEncoded("DER"), SesTestProfile.SealInfo(parsed.Version, _certificate).GetEncoded("DER"))) return false;
        var info = parsed.Version == SesVersion.V1 ? SesDer.Sequence(seal[1], 3) : seal;
        var offset = parsed.Version == SesVersion.V1 ? 0 : 1;
        if (!SesTestCrypto.Equal(SesDer.Octets(info[offset], 64 * 1024), _certificate) ||
            SesDer.Oid(info[offset + 1]) != SignatureMethod) return false;
        var makerData = parsed.Version == SesVersion.V1
            ? new DerSequence(seal[0], info[0], info[1]).GetEncoded("DER") : seal[0].GetEncoded("DER");
        if (!SesTestCrypto.Verify(makerData, SesDer.Bits(info[offset + 2], 80), _key)) return false;
        return SesTestCrypto.Verify(parsed.ToBeSignedDer, parsed.Signature, _key);
    }
}
