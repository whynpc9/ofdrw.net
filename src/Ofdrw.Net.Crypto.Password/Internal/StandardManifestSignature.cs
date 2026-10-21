using System.Security.Cryptography;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto.Signers;
using Org.BouncyCastle.X509;

namespace Ofdrw.Net.Crypto.Password.Internal;

// Narrow single-signer GB/T 35275 shape used by the pinned Java GMProtectSigner.
internal static class StandardManifestSignature
{
    private const string DataOid = "1.2.156.10197.6.1.4.2.1";
    private const string SignedDataOid = "1.2.156.10197.6.1.4.2.2";
    private const string Sm3Oid = "1.2.156.10197.1.401";
    private const string Sm2Oid = "1.2.156.10197.1.301.1";

    internal static async Task<byte[]> SignAsync(byte[] manifest, OfdStandardManifestSigner signer, CancellationToken token)
    {
        var digest = PasswordPrimitives.Sm3(manifest);
        try
        {
            var signature = await signer.SignDigestAsync((byte[])digest.Clone(), token).ConfigureAwait(false);
            token.ThrowIfCancellationRequested();
            if (signature is null || signature.Length == 0 || signature.Length > 80 ||
                !VerifySm2(digest, signature, signer.CertificateDer))
                throw new InvalidDataException("Manifest signer returned an invalid SM2 signature.");
            var certificate = Certificate(signer.CertificateDer);
            var digestAlgorithm = new DerSequence(new DerObjectIdentifier(Sm3Oid));
            var signerInfo = new DerSequence(new DerInteger(1),
                new DerSequence(certificate.Issuer, certificate.SerialNumber), digestAlgorithm,
                new DerSequence(new DerObjectIdentifier(Sm2Oid)), new DerOctetString(signature));
            var data = new DerSequence(new DerObjectIdentifier(DataOid),
                new DerTaggedObject(true, 0, new DerOctetString(digest)));
            var signedData = new DerSequence(new DerInteger(1), new DerSet((Asn1Encodable)digestAlgorithm),
                data, new DerTaggedObject(false, 0, new DerSet((Asn1Encodable)certificate)),
                new DerSet((Asn1Encodable)signerInfo));
            return new DerSequence(new DerObjectIdentifier(SignedDataOid),
                new DerTaggedObject(true, 0, signedData)).GetEncoded("DER");
        }
        finally { CryptographicOperations.ZeroMemory(digest); }
    }

    internal static bool Verify(byte[] manifest, byte[] signedValue, byte[] expectedCertificateDer)
    {
        if (signedValue.Length > 128 * 1024 || expectedCertificateDer.Length > 64 * 1024) return false;
        try
        {
            var outer = Sequence(Canonical(signedValue), 2);
            if (Oid(outer[0]) != SignedDataOid) return false;
            var signedData = Sequence(Explicit(outer[1]), 5);
            if (Number(signedData[0]) != 1 || signedData[1].ToAsn1Object() is not Asn1Set digestSet ||
                digestSet.Count != 1 || Algorithm(digestSet[0]) != Sm3Oid) return false;
            var content = Sequence(signedData[2], 2);
            if (Oid(content[0]) != DataOid) return false;
            var digest = PasswordPrimitives.Sm3(manifest);
            try
            {
                if (!CryptographicOperations.FixedTimeEquals(Asn1OctetString.GetInstance(Explicit(content[1])).GetOctets(), digest))
                    return false;
                if (signedData[3].ToAsn1Object() is not Asn1TaggedObject certTag || certTag.TagNo != 0)
                    return false;
                var certificates = Asn1Set.GetInstance(certTag, false);
                if (certificates.Count != 1 || !CryptographicOperations.FixedTimeEquals(
                        certificates[0].ToAsn1Object().GetEncoded("DER"), expectedCertificateDer)) return false;
                if (signedData[4].ToAsn1Object() is not Asn1Set signers || signers.Count != 1) return false;
                var signer = Sequence(signers[0], 5);
                var cert = Certificate(expectedCertificateDer);
                var identity = Sequence(signer[1], 2);
                if (Number(signer[0]) != 1 || Algorithm(signer[2]) != Sm3Oid || Algorithm(signer[3]) != Sm2Oid ||
                    !identity[0].ToAsn1Object().Equals(cert.Issuer.ToAsn1Object()) ||
                    !identity[1].ToAsn1Object().Equals(cert.SerialNumber.ToAsn1Object())) return false;
                var signature = Asn1OctetString.GetInstance(signer[4]).GetOctets();
                return signature.Length <= 80 && VerifySm2(digest, signature, expectedCertificateDer);
            }
            finally { CryptographicOperations.ZeroMemory(digest); }
        }
        catch (Exception ex) when (ex is not OutOfMemoryException and not StackOverflowException)
        { return false; }
    }

    private static bool VerifySm2(byte[] digest, byte[] signature, byte[] certificateDer)
    {
        var cert = new X509CertificateParser().ReadCertificate(certificateDer);
        if (cert is null || cert.GetPublicKey() is not Org.BouncyCastle.Crypto.Parameters.ECPublicKeyParameters key ||
            !GMObjectIdentifiers.sm2p256v1.Equals(key.PublicKeyParamSet) || key.Q.IsInfinity || !key.Q.IsValid())
            return false;
        var verifier = new SM2Signer(); verifier.Init(false, key);
        verifier.BlockUpdate(digest, 0, digest.Length);
        return verifier.VerifySignature(signature);
    }
    private static X509CertificateStructure Certificate(byte[] der) =>
        X509CertificateStructure.GetInstance(Canonical(der));
    private static Asn1Object Canonical(byte[] der)
    {
        var value = Asn1Object.FromByteArray(der);
        if (!der.AsSpan().SequenceEqual(value.GetEncoded("DER"))) throw new InvalidDataException("Noncanonical DER.");
        return value;
    }
    private static Asn1Sequence Sequence(Asn1Encodable value, int count)
    {
        var seq = Asn1Sequence.GetInstance(value.ToAsn1Object());
        if (seq.Count != count) throw new InvalidDataException("Unexpected manifest signature shape.");
        return seq;
    }
    private static Asn1Object Explicit(Asn1Encodable value)
    {
        if (value.ToAsn1Object() is not Asn1TaggedObject tag || tag.TagNo != 0 || !tag.IsExplicit())
            throw new InvalidDataException("Expected explicit GB/T 35275 content.");
        return tag.GetExplicitBaseObject().ToAsn1Object();
    }
    private static string Oid(Asn1Encodable value) => DerObjectIdentifier.GetInstance(value).Id;
    private static int Number(Asn1Encodable value) => DerInteger.GetInstance(value).Value.IntValue;
    private static string Algorithm(Asn1Encodable value) => Oid(Sequence(value, 1)[0]);
}
