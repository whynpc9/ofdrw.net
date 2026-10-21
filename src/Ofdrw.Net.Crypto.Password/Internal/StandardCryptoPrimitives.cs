using System.Security.Cryptography;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Crypto.Engines;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;

namespace Ofdrw.Net.Crypto.Password.Internal;

internal static class StandardCryptoPrimitives
{
    internal static byte[] PasswordKey(string password)
    {
        var bytes = PasswordPrimitives.PasswordBytes(password);
        try
        {
            var input = new byte[bytes.Length + 4];
            try
            {
                bytes.CopyTo(input, 0); input[^1] = 1;
                var digest = PasswordPrimitives.Sm3(input);
                try { return digest[..16]; }
                finally { CryptographicOperations.ZeroMemory(digest); }
            }
            finally { CryptographicOperations.ZeroMemory(input); }
        }
        finally { CryptographicOperations.ZeroMemory(bytes); }
    }

    internal static byte[] EncryptSm2(byte[] plain, byte[] certificateDer)
    {
        if (certificateDer.Length is < 64 or > 64 * 1024)
            throw new InvalidDataException("Invalid SM2 recipient certificate size.");
        var cert = new X509CertificateParser().ReadCertificate(certificateDer)
            ?? throw new InvalidDataException("Invalid SM2 recipient certificate.");
        if (cert.GetPublicKey() is not ECPublicKeyParameters publicKey ||
            !GMObjectIdentifiers.sm2p256v1.Equals(publicKey.PublicKeyParamSet) || publicKey.Q.IsInfinity || !publicKey.Q.IsValid())
            throw new InvalidDataException("Recipient certificate must contain a valid named SM2 key.");
        var engine = new SM2Engine(SM2Engine.Mode.C1C3C2);
        engine.Init(true, new ParametersWithRandom(publicKey, new SecureRandom()));
        var raw = engine.ProcessBlock(plain, 0, plain.Length);
        try { return ToJavaDer(raw); }
        finally { CryptographicOperations.ZeroMemory(raw); }
    }

    internal static byte[] DecryptSm2(byte[] wrapped, byte[] privateKeyPkcs8Der)
    {
        ECPrivateKeyParameters privateKey;
        try
        {
            if (PrivateKeyFactory.CreateKey(privateKeyPkcs8Der) is not ECPrivateKeyParameters key ||
                !GMObjectIdentifiers.sm2p256v1.Equals(key.PublicKeyParamSet))
                throw new InvalidDataException("Recipient private key must use the named SM2 curve.");
            privateKey = key;
        }
        catch (Exception ex) when (ex is ArgumentException or IOException)
        { throw new InvalidDataException("Invalid recipient private key.", ex); }
        var raw = FromJavaDer(wrapped);
        try
        {
            var engine = new SM2Engine(SM2Engine.Mode.C1C3C2);
            engine.Init(false, privateKey);
            var plain = engine.ProcessBlock(raw, 0, raw.Length);
            if (plain.Length != 16) { CryptographicOperations.ZeroMemory(plain); throw new InvalidDataException("Invalid recovered file key."); }
            return plain;
        }
        catch (Exception ex) when (ex is not InvalidDataException)
        { throw new InvalidDataException("Certificate or wrapped file key is invalid.", ex); }
        finally { CryptographicOperations.ZeroMemory(raw); }
    }

    private static byte[] ToJavaDer(byte[] raw)
    {
        if (raw.Length != 113 || raw[0] != 4) throw new InvalidDataException("Unexpected SM2 FEK ciphertext.");
        var x = new Org.BouncyCastle.Math.BigInteger(1, raw.AsSpan(1, 32).ToArray());
        var y = new Org.BouncyCastle.Math.BigInteger(1, raw.AsSpan(33, 32).ToArray());
        return new DerSequence(new DerInteger(x), new DerInteger(y),
            new DerOctetString(raw.AsSpan(65, 32).ToArray()), new DerOctetString(raw.AsSpan(97, 16).ToArray())).GetEncoded("DER");
    }

    private static byte[] FromJavaDer(byte[] encoded)
    {
        if (encoded.Length > 256) throw new InvalidDataException("SM2 wrapper is oversized.");
        var parsed = Asn1Object.FromByteArray(encoded);
        if (!encoded.AsSpan().SequenceEqual(parsed.GetEncoded("DER"))) throw new InvalidDataException("Noncanonical SM2 wrapper.");
        var sequence = Asn1Sequence.GetInstance(parsed);
        if (sequence.Count != 4) throw new InvalidDataException("Invalid SM2 wrapper.");
        var x = DerInteger.GetInstance(sequence[0]).PositiveValue.ToByteArrayUnsigned();
        var y = DerInteger.GetInstance(sequence[1]).PositiveValue.ToByteArrayUnsigned();
        var hash = Asn1OctetString.GetInstance(sequence[2]).GetOctets();
        var cipher = Asn1OctetString.GetInstance(sequence[3]).GetOctets();
        if (x.Length > 32 || y.Length > 32 || hash.Length != 32 || cipher.Length != 16)
            throw new InvalidDataException("Invalid SM2 FEK wrapper fields.");
        var raw = new byte[113]; raw[0] = 4;
        x.CopyTo(raw, 33 - x.Length); y.CopyTo(raw, 65 - y.Length);
        hash.CopyTo(raw, 65); cipher.CopyTo(raw, 97);
        return raw;
    }
}
