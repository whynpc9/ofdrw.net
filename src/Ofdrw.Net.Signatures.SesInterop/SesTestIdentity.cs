using System;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Operators;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>An ephemeral SM2 key and self-signed certificate for dev/interop tests only.</summary>
/// <remarks>No private-key export, production identity, certificate trust store or device integration is provided.</remarks>
public sealed class SesTestIdentity
{
    private readonly byte[] _certificate;
    internal ECPrivateKeyParameters PrivateKey { get; }
    private SesTestIdentity(ECPrivateKeyParameters privateKey, byte[] certificate)
    {
        PrivateKey = privateKey; _certificate = certificate;
    }
    /// <summary>Copy of the generated self-signed test certificate; pin this exact value in the test verifier.</summary>
    public byte[] CertificateDer => (byte[])_certificate.Clone();

    /// <summary>Generate a fresh SM2 test key with cryptographic randomness and a visibly marked test certificate.</summary>
    public static SesTestIdentity Generate()
    {
        var random = new SecureRandom();
        var generator = new ECKeyPairGenerator();
        generator.Init(new ECKeyGenerationParameters(GMObjectIdentifiers.sm2p256v1, random));
        Org.BouncyCastle.Crypto.AsymmetricCipherKeyPair pair;
        // SM2 excludes d=n-1 because (1+d) must have a multiplicative inverse.
        do { pair = generator.GenerateKeyPair(); }
        while (((ECPrivateKeyParameters)pair.Private).D.Equals(SesTestCrypto.Domain.N.Subtract(BigInteger.One)));
        var certificate = new X509V3CertificateGenerator();
        var name = new X509Name("CN=OFD SES DEV INTEROP TEST");
        certificate.SetSerialNumber(new BigInteger(128, random).Add(BigInteger.One));
        certificate.SetIssuerDN(name); certificate.SetSubjectDN(name);
        certificate.SetNotBefore(SesTestProfile.Begin);
        certificate.SetNotAfter(SesTestProfile.End);
        certificate.SetPublicKey(pair.Public);
        certificate.AddExtension(X509Extensions.BasicConstraints, true, new BasicConstraints(false));
        certificate.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.DigitalSignature));
        var encoded = certificate.Generate(new Asn1SignatureFactory("SM3WITHSM2", pair.Private, random)).GetEncoded();
        return new SesTestIdentity((ECPrivateKeyParameters)pair.Private, encoded);
    }
}
