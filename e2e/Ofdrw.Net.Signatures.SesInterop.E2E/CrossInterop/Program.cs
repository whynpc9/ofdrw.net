using System.Security.Cryptography;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Operators;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Crypto.Signers;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using Ofdrw.Net.Signatures.SesInterop;

const string property = "/Doc_0/Signs/Sign_0/Signature.xml";
if (args.Length != 3 || (args[0] != "produce" && args[0] != "verify"))
    throw new ArgumentException("Expected produce|verify xml-path output-directory");
var xml = File.ReadAllBytes(args[1]); var output = Path.GetFullPath(args[2]);
Directory.CreateDirectory(output);
string FileAt(string name) => Path.Combine(output, name);
if (args[0] == "produce")
{
    var maker = TestIdentity.Create(); var signer = TestIdentity.Create();
    var seal = MakeSeal(maker, signer.Certificate);
    var v5 = await new SesV5ExternalSignatureProvider(signer.Certificate, seal, signer).SignAsync(xml, property);
    var gbt = await new Gbt35275SignatureProvider(signer.Certificate, signer).SignAsync(xml, property);
    File.WriteAllBytes(FileAt("net-v5.dat"), v5);
    File.WriteAllBytes(FileAt("net-gbt35275.dat"), gbt);
    File.WriteAllBytes(FileAt("net-signer.cer"), signer.Certificate);
    File.WriteAllBytes(FileAt("net-maker.cer"), maker.Certificate);
    File.WriteAllBytes(FileAt("net-seal.esl"), seal);
    if (!await new SesV5PinnedVerifier(signer.Certificate, maker.Certificate).VerifyAsync(xml, v5, property) ||
        !await new Gbt35275PinnedVerifier(signer.Certificate).VerifyAsync(xml, gbt, property))
        throw new Exception(".NET outbound sample failed local verification.");
    Console.WriteLine(".NET produced V5 and GB/T 35275 samples; private keys were not persisted.");
}
else
{
    var signer = File.ReadAllBytes(FileAt("java-signer.cer"));
    var maker = File.ReadAllBytes(FileAt("java-maker.cer"));
    var v5 = File.ReadAllBytes(FileAt("java-v5.dat"));
    var gbt = File.ReadAllBytes(FileAt("java-gbt35275.dat"));
    if (SesSignedValueReader.Parse(v5).Version != SesVersion.V5 ||
        !await new SesV5PinnedVerifier(signer, maker).VerifyAsync(xml, v5, property) ||
        !await new Gbt35275PinnedVerifier(signer).VerifyAsync(xml, gbt, property))
        throw new Exception(".NET rejected Java-generated signed-value samples.");
    Console.WriteLine(".NET verified Java V5 and GB/T 35275 signed-value samples.");
}
foreach (var path in Directory.GetFiles(output).Where(path =>
    Path.GetFileName(path).StartsWith(args[0] == "produce" ? "net-" : "java-", StringComparison.Ordinal)))
    Console.WriteLine($"{Path.GetFileName(path)} {new FileInfo(path).Length} {Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)))}");

static byte[] MakeSeal(TestIdentity maker, byte[] signerCertificate)
{
    var start = new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    var end = new DateTime(2049, 12, 31, 23, 59, 59, DateTimeKind.Utc);
    var info = new DerSequence(
        new DerSequence(new DerIA5String("ES"), new DerInteger(5), new DerIA5String("Ofdrw.Net.R12.TEST")),
        new DerIA5String("r12-cross-interop"),
        new DerSequence(new DerInteger(0), new DerUtf8String("OFD R12 INTEROP TEST"),
            new DerInteger(1), new DerSequence(new DerOctetString(signerCertificate)),
            new DerGeneralizedTime(start), new DerGeneralizedTime(start), new DerGeneralizedTime(end)),
        new DerSequence(new DerIA5String("PNG"), new DerOctetString(Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=")),
            new DerInteger(1), new DerInteger(1)));
    return new DerSequence(info, new DerOctetString(maker.Certificate),
        new DerObjectIdentifier("1.2.156.10197.1.501"),
        new DerBitString(maker.Sign(info.GetEncoded("DER")))).GetEncoded("DER");
}

sealed class TestIdentity : ISesExternalSm2Signer
{
    private static readonly byte[] Id = System.Text.Encoding.ASCII.GetBytes("1234567812345678");
    private readonly ECPrivateKeyParameters _private;
    public byte[] Certificate { get; }
    private TestIdentity(ECPrivateKeyParameters key, byte[] certificate)
    { _private = key; Certificate = certificate; }
    public static TestIdentity Create()
    {
        var random = new SecureRandom(); var generator = new ECKeyPairGenerator();
        generator.Init(new ECKeyGenerationParameters(GMObjectIdentifiers.sm2p256v1, random));
        Org.BouncyCastle.Crypto.AsymmetricCipherKeyPair pair;
        var domain = new ECNamedDomainParameters(GMObjectIdentifiers.sm2p256v1, GMNamedCurves.GetByName("sm2p256v1"));
        do { pair = generator.GenerateKeyPair(); }
        while (((ECPrivateKeyParameters)pair.Private).D.Equals(domain.N.Subtract(BigInteger.One)));
        var certificate = new X509V3CertificateGenerator();
        var name = new X509Name("CN=OFD R12 INTEROP TEST");
        certificate.SetSerialNumber(new BigInteger(128, random).Add(BigInteger.One));
        certificate.SetIssuerDN(name); certificate.SetSubjectDN(name);
        certificate.SetNotBefore(new DateTime(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc));
        certificate.SetNotAfter(new DateTime(2049, 12, 31, 23, 59, 59, DateTimeKind.Utc));
        certificate.SetPublicKey(pair.Public);
        certificate.AddExtension(X509Extensions.BasicConstraints, true, new BasicConstraints(false));
        certificate.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.DigitalSignature));
        var bytes = certificate.Generate(new Asn1SignatureFactory("SM3WITHSM2", pair.Private, random)).GetEncoded();
        return new TestIdentity((ECPrivateKeyParameters)pair.Private, bytes);
    }
    public byte[] Sign(byte[] data)
    {
        var signer = new SM2Signer();
        signer.Init(true, new ParametersWithID(new ParametersWithRandom(_private, new SecureRandom()), Id));
        signer.BlockUpdate(data, 0, data.Length);
        return signer.GenerateSignature();
    }
    public Task<byte[]> SignAsync(byte[] data, CancellationToken cancellationToken = default)
    { cancellationToken.ThrowIfCancellationRequested(); return Task.FromResult(Sign(data)); }
}
