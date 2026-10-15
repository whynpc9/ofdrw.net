using System;
using System.IO;
using System.Text;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Crypto.Signers;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;

namespace Ofdrw.Net.Signatures.SesInterop;

internal static class SesTestCrypto
{
    internal const string Algorithm = "1.2.156.10197.1.501";
    internal static readonly ECNamedDomainParameters Domain = new(
        GMObjectIdentifiers.sm2p256v1, GMNamedCurves.GetByName("sm2p256v1"));
    internal static readonly byte[] Identity = Encoding.ASCII.GetBytes("1234567812345678");

    internal static byte[] Sign(byte[] data, ECPrivateKeyParameters key)
    {
        var signer = new SM2Signer();
        signer.Init(true, new ParametersWithID(new ParametersWithRandom(key, new SecureRandom()), Identity));
        signer.BlockUpdate(data, 0, data.Length);
        return signer.GenerateSignature();
    }
    internal static bool Verify(byte[] data, byte[] signature, ECPublicKeyParameters key, byte[]? identity = null)
    {
        // Avoid BER/zero/out-of-range r/s encodings even if a backend accepts them.
        var pair = SesDer.Sequence(SesDer.Read(signature, 80), 2);
        foreach (var part in pair)
        {
            if (part.ToAsn1Object() is not DerInteger integer || integer.Value.SignValue <= 0 ||
                integer.Value.CompareTo(Domain.N) >= 0) return false;
        }
        var signer = new SM2Signer();
        signer.Init(false, new ParametersWithID(key, identity ?? Identity));
        signer.BlockUpdate(data, 0, data.Length);
        return signer.VerifySignature(signature);
    }
    internal static ECPublicKeyParameters ReadPinnedCertificate(byte[] certificate)
    {
        SesDer.Read(certificate, 64 * 1024);
        var parsed = new X509CertificateParser().ReadCertificate(certificate);
        if (parsed.GetPublicKey() is not ECPublicKeyParameters key ||
            key.PublicKeyParamSet is null || !key.PublicKeyParamSet.Equals(GMObjectIdentifiers.sm2p256v1) ||
            !key.Parameters.Equals(Domain) || key.Q.IsInfinity || !key.Q.IsValid())
            throw new InvalidDataException("Test certificate must use the named SM2 curve.");
        if (parsed.SubjectDN.ToString() != "CN=OFD SES DEV INTEROP TEST" ||
            !parsed.SubjectDN.Equivalent(parsed.IssuerDN) || parsed.SigAlgOid != Algorithm ||
            parsed.NotBefore.ToUniversalTime() != SesTestProfile.Begin || parsed.NotAfter.ToUniversalTime() != SesTestProfile.End)
            throw new InvalidDataException("Expected a self-signed dev/interop test certificate.");
        parsed.Verify(key);
        return key;
    }
    internal static bool Equal(byte[] left, byte[] right)
    {
        int difference = left.Length ^ right.Length;
        for (int i = 0; i < Math.Min(left.Length, right.Length); i++) difference |= left[i] ^ right[i];
        return difference == 0;
    }
}
