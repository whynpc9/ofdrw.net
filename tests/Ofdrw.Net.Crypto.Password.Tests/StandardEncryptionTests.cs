using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Asn1.Pkcs;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Operators;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Crypto.Signers;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.OpenSsl;
using Org.BouncyCastle.Pkcs;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;

namespace Ofdrw.Net.Crypto.Password.Tests;

public sealed class StandardEncryptionTests
{
    [Fact]
    public async Task JavaPasswordFixtureRestoresExactOriginalEntriesWhenExplicitlyAllowingUnverifiedContents()
    {
        var dir = FixtureDirectory();
        var cert = (Org.BouncyCastle.X509.X509Certificate)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_cert.pem")))).ReadObject()!;
        var options = new OfdStandardDecryptionOptions
        { ExpectedManifestSignerCertificateDer = cert.GetEncoded(), AllowUnverifiedContents = true };
        var encrypted = File.ReadAllBytes(Path.Combine(dir, "java-pass.ofd"));
        var result = await OfdStandardEncryption.DecryptAsync(new MemoryStream(encrypted),
            OfdStandardCredential.ForPassword("bob", "second-pass"), options);
        Assert.True(result.ManifestSignatureVerified && result.ExactEntrySetVerified);
        Assert.False(result.FileContentsVerified);
        AssertEqual(Unzip(File.ReadAllBytes(Path.Combine(dir, "hello.ofd"))), Unzip(result.OfdBytes));
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(new MemoryStream(encrypted),
            OfdStandardCredential.ForPassword("bob", "second-pass"),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = cert.GetEncoded() }));
    }

    [Fact]
    public async Task JavaCaseOnlyDatCipherPathsRestoreOriginalPathCaseAndBytes()
    {
        var dir = FixtureDirectory();
        var cert = (Org.BouncyCastle.X509.X509Certificate)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_cert.pem")))).ReadObject()!;
        var result = await OfdStandardEncryption.DecryptAsync(
            new MemoryStream(File.ReadAllBytes(Path.Combine(dir, "java-case-dat.ofd"))),
            OfdStandardCredential.ForPassword("bob", "second-pass"),
            new OfdStandardDecryptionOptions
            { ExpectedManifestSignerCertificateDer = cert.GetEncoded(), AllowUnverifiedContents = true });
        Assert.True(result.ManifestSignatureVerified && result.ExactEntrySetVerified);
        Assert.False(result.FileContentsVerified);
        AssertEqual(Unzip(File.ReadAllBytes(Path.Combine(dir, "case-source.ofd"))), Unzip(result.OfdBytes));
    }

    [Fact]
    public async Task JavaCertificateFixtureRestoresExactOriginalEntriesWhenExplicitlyAllowingUnverifiedContents()
    {
        var dir = FixtureDirectory();
        var cert = (Org.BouncyCastle.X509.X509Certificate)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_cert.pem")))).ReadObject()!;
        var pair = (Org.BouncyCastle.Crypto.AsymmetricCipherKeyPair)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_key.pem")))).ReadObject()!;
        var result = await OfdStandardEncryption.DecryptAsync(
            new MemoryStream(File.ReadAllBytes(Path.Combine(dir, "java-cert.ofd"))),
            OfdStandardCredential.ForCertificate("bob", PrivateKeyInfoFactory.CreatePrivateKeyInfo(pair.Private).GetEncoded(),
                cert.GetEncoded()),
            new OfdStandardDecryptionOptions
            { ExpectedManifestSignerCertificateDer = cert.GetEncoded(), AllowUnverifiedContents = true });
        Assert.True(result.ManifestSignatureVerified && result.ExactEntrySetVerified);
        Assert.False(result.FileContentsVerified);
        AssertEqual(Unzip(File.ReadAllBytes(Path.Combine(dir, "hello.ofd"))), Unzip(result.OfdBytes));
    }

    [Fact]
    public async Task MalformedJavaCertificateWrapperIsReportedAsInvalidPackageData()
    {
        var dir = FixtureDirectory();
        var cert = (Org.BouncyCastle.X509.X509Certificate)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_cert.pem")))).ReadObject()!;
        var pair = (Org.BouncyCastle.Crypto.AsymmetricCipherKeyPair)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_key.pem")))).ReadObject()!;
        var entries = Unzip(File.ReadAllBytes(Path.Combine(dir, "java-cert.ofd")));
        var seed = XDocument.Parse(Encoding.UTF8.GetString(entries["decryptseed.dat"]));
        seed.Descendants().Single(x => x.Name.LocalName == "UserInfo" && (string?)x.Attribute("UserName") == "bob")
            .Elements().Single(x => x.Name.LocalName == "EncryptedWK").Value = Convert.ToBase64String(new byte[64]);
        entries["decryptseed.dat"] = Encoding.UTF8.GetBytes(seed.ToString());
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(
            new MemoryStream(Zip(entries)),
            OfdStandardCredential.ForCertificate("bob", PrivateKeyInfoFactory.CreatePrivateKeyInfo(pair.Private).GetEncoded(),
                cert.GetEncoded()),
            new OfdStandardDecryptionOptions
            { ExpectedManifestSignerCertificateDer = cert.GetEncoded(), AllowUnverifiedContents = true }));
    }

    [Fact]
    public async Task NetPackagesForSeparateJavaDecryptionGate()
    {
        var dir = FixtureDirectory();
        var cert = (Org.BouncyCastle.X509.X509Certificate)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_cert.pem")))).ReadObject()!;
        var pair = (Org.BouncyCastle.Crypto.AsymmetricCipherKeyPair)new PemReader(
            new StringReader(File.ReadAllText(Path.Combine(dir, "sign_key.pem")))).ReadObject()!;
        var certificate = cert.GetEncoded();
        var privateKey = (ECPrivateKeyParameters)pair.Private;
        var signer = new OfdStandardManifestSigner(certificate, (digest, token) =>
        {
            token.ThrowIfCancellationRequested();
            var sm2 = new SM2Signer(); sm2.Init(true, new ParametersWithRandom(privateKey, new SecureRandom()));
            sm2.BlockUpdate(digest, 0, digest.Length);
            return Task.FromResult(sm2.GenerateSignature());
        });
        var input = File.ReadAllBytes(Path.Combine(dir, "hello.ofd"));
        var password = new OfdStandardEncryptionOptions { ManifestSigner = signer };
        password.Recipients.Add(OfdStandardRecipient.ForPassword("alice", "first-pass"));
        password.Recipients.Add(OfdStandardRecipient.ForPassword("bob", "second-pass"));
        var netPassword = await OfdStandardEncryption.EncryptAsync(new MemoryStream(input), password);
        var netPasswordResult = await OfdStandardEncryption.DecryptAsync(new MemoryStream(netPassword),
            OfdStandardCredential.ForPassword("bob", "second-pass"),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = certificate });
        Assert.True(netPasswordResult.FileContentsVerified);
        AssertEqual(Unzip(input), Unzip(netPasswordResult.OfdBytes));
        var certificateOptions = new OfdStandardEncryptionOptions { ManifestSigner = signer };
        certificateOptions.Recipients.Add(OfdStandardRecipient.ForCertificate("alice", certificate));
        certificateOptions.Recipients.Add(OfdStandardRecipient.ForCertificate("bob", certificate));
        var netCertificate = await OfdStandardEncryption.EncryptAsync(new MemoryStream(input), certificateOptions);
        var netCertificateResult = await OfdStandardEncryption.DecryptAsync(new MemoryStream(netCertificate),
            OfdStandardCredential.ForCertificate("bob", PrivateKeyInfoFactory.CreatePrivateKeyInfo(pair.Private).GetEncoded(), certificate),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = certificate });
        Assert.True(netCertificateResult.FileContentsVerified);
        AssertEqual(Unzip(input), Unzip(netCertificateResult.OfdBytes));
        var output = Environment.GetEnvironmentVariable("R14_JAVA_OUTPUT_DIR");
        if (!string.IsNullOrEmpty(output))
        {
            Directory.CreateDirectory(output);
            File.WriteAllBytes(Path.Combine(output, "net-pass.ofd"), netPassword);
            File.WriteAllBytes(Path.Combine(output, "net-cert.ofd"), netCertificate);
        }
    }

    [Fact]
    public async Task PasswordRecipientsRestoreExactEntriesAndRejectTampering()
    {
        var identity = TestIdentity.Create();
        var input = Fixture();
        var options = new OfdStandardEncryptionOptions { ManifestSigner = identity.Signer() };
        options.Recipients.Add(OfdStandardRecipient.ForPassword("alice", "first 🔒"));
        options.Recipients.Add(OfdStandardRecipient.ForPassword("bob", "second"));
        var encrypted = await OfdStandardEncryption.EncryptAsync(new MemoryStream(Zip(input)), options);
        foreach (var credential in new[] { OfdStandardCredential.ForPassword("alice", "first 🔒"),
                     OfdStandardCredential.ForPassword("bob", "second") })
        {
            var result = await OfdStandardEncryption.DecryptAsync(new MemoryStream(encrypted), credential,
                new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = identity.Certificate });
            Assert.True(result.ManifestSignatureVerified && result.ExactEntrySetVerified && result.FileContentsVerified);
            AssertEqual(input, Unzip(result.OfdBytes));
        }
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(new MemoryStream(encrypted),
            OfdStandardCredential.ForPassword("alice", "wrong"),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = identity.Certificate }));
        var stored = Unzip(encrypted); stored["Cipher/entry-00001.dat"][0] ^= 1;
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(new MemoryStream(Zip(stored)),
            OfdStandardCredential.ForPassword("alice", "first 🔒"),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = identity.Certificate }));
        stored = Unzip(encrypted); stored.Add("unexpected.txt", [1]);
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(new MemoryStream(Zip(stored)),
            OfdStandardCredential.ForPassword("alice", "first 🔒"),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = identity.Certificate }));
    }

    [Fact]
    public async Task Sm2CertificateRecipientsRestoreExactEntriesAndRejectWrongIdentity()
    {
        var signer = TestIdentity.Create(); var alice = TestIdentity.Create(); var bob = TestIdentity.Create();
        var input = Fixture();
        var options = new OfdStandardEncryptionOptions { ManifestSigner = signer.Signer() };
        options.Recipients.Add(OfdStandardRecipient.ForCertificate("alice", alice.Certificate));
        options.Recipients.Add(OfdStandardRecipient.ForCertificate("bob", bob.Certificate));
        var encrypted = await OfdStandardEncryption.EncryptAsync(new MemoryStream(Zip(input)), options);
        foreach (var (name, identity) in new[] { ("alice", alice), ("bob", bob) })
        {
            var result = await OfdStandardEncryption.DecryptAsync(new MemoryStream(encrypted),
                OfdStandardCredential.ForCertificate(name, identity.PrivatePkcs8, identity.Certificate),
                new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = signer.Certificate });
            Assert.True(result.FileContentsVerified); AssertEqual(input, Unzip(result.OfdBytes));
        }
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(new MemoryStream(encrypted),
            OfdStandardCredential.ForCertificate("alice", bob.PrivatePkcs8, alice.Certificate),
            new OfdStandardDecryptionOptions { ExpectedManifestSignerCertificateDer = signer.Certificate }));
    }

    [Fact]
    public async Task SignedInventoryRejectsManifestTamperingAndMissingCiphertext()
    {
        var identity = TestIdentity.Create();
        var options = new OfdStandardEncryptionOptions { ManifestSigner = identity.Signer() };
        options.Recipients.Add(OfdStandardRecipient.ForPassword("alice", "correct"));
        var encrypted = await OfdStandardEncryption.EncryptAsync(new MemoryStream(Zip(Fixture())), options);
        var credential = OfdStandardCredential.ForPassword("alice", "correct");
        var verifier = new OfdStandardDecryptionOptions
        { ExpectedManifestSignerCertificateDer = identity.Certificate, AllowUnverifiedContents = true };
        var entries = Unzip(encrypted);
        entries["OFDEntries.xml"][0] ^= 1;
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(
            new MemoryStream(Zip(entries)), credential, verifier));
        entries = Unzip(encrypted);
        entries.Remove("Cipher/entry-00001.dat");
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(
            new MemoryStream(Zip(entries)), credential, verifier));
        entries = Unzip(encrypted);
        var renamedPayload = entries["Cipher/entry-00001.dat"];
        entries.Remove("Cipher/entry-00001.dat");
        entries.Add("Cipher/ENTRY-00001.dat", renamedPayload);
        await Assert.ThrowsAsync<InvalidDataException>(() => OfdStandardEncryption.DecryptAsync(
            new MemoryStream(Zip(entries)), credential, verifier));
    }

    private static Dictionary<string, byte[]> Fixture() => new(StringComparer.OrdinalIgnoreCase)
    {
        ["OFD.xml"] = Encoding.UTF8.GetBytes("<OFD xmlns='http://www.ofdspec.org'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
        ["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes("<Document xmlns='http://www.ofdspec.org'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages></Document>"),
        ["Doc_0/Pages/P0/Content.xml"] = Encoding.UTF8.GetBytes("中英文 ABC"),
        ["custom/data.bin"] = [0, 1, 2, 255]
    };
    private static string FixtureDirectory() => Environment.GetEnvironmentVariable("R14_JAVA_FIXTURE_DIR")
        ?? Path.Combine(AppContext.BaseDirectory, "Fixtures", "standard-java");
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var item in entries)
            { using var stream = zip.CreateEntry(item.Key, CompressionLevel.NoCompression).Open(); stream.Write(item.Value); }
        return output.ToArray();
    }
    private static Dictionary<string, byte[]> Unzip(byte[] data)
    {
        using var archive = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
        return archive.Entries.Where(x => !x.FullName.EndsWith("/", StringComparison.Ordinal)).ToDictionary(x => x.FullName, x =>
        { using var stream = x.Open(); using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray(); },
            StringComparer.OrdinalIgnoreCase);
    }
    private static void AssertEqual(Dictionary<string, byte[]> expected, Dictionary<string, byte[]> actual)
    {
        Assert.Equal(expected.Keys.Order(), actual.Keys.Order());
        foreach (var item in expected) Assert.Equal(item.Value, actual[item.Key]);
    }

    private sealed class TestIdentity
    {
        private readonly ECPrivateKeyParameters _private;
        internal byte[] Certificate { get; }
        internal byte[] PrivatePkcs8 => PrivateKeyInfoFactory.CreatePrivateKeyInfo(_private).GetEncoded();
        private TestIdentity(ECPrivateKeyParameters privateKey, byte[] certificate)
        { _private = privateKey; Certificate = certificate; }
        internal static TestIdentity Create()
        {
            var random = new SecureRandom(); var generator = new ECKeyPairGenerator();
            generator.Init(new ECKeyGenerationParameters(GMObjectIdentifiers.sm2p256v1, random));
            var pair = generator.GenerateKeyPair();
            var cert = new X509V3CertificateGenerator(); var name = new X509Name("CN=OFD R14 TEST");
            cert.SetSerialNumber(new BigInteger(128, random).Add(BigInteger.One));
            cert.SetIssuerDN(name); cert.SetSubjectDN(name);
            cert.SetNotBefore(DateTime.UtcNow.AddDays(-1)); cert.SetNotAfter(DateTime.UtcNow.AddDays(30));
            cert.SetPublicKey(pair.Public);
            cert.AddExtension(X509Extensions.BasicConstraints, true, new BasicConstraints(false));
            cert.AddExtension(X509Extensions.KeyUsage, true, new KeyUsage(KeyUsage.DigitalSignature | KeyUsage.KeyEncipherment));
            return new TestIdentity((ECPrivateKeyParameters)pair.Private,
                cert.Generate(new Asn1SignatureFactory("SM3WITHSM2", pair.Private, random)).GetEncoded());
        }
        internal OfdStandardManifestSigner Signer() => new(Certificate, (digest, token) =>
        {
            token.ThrowIfCancellationRequested();
            var signer = new SM2Signer(); signer.Init(true, new ParametersWithRandom(_private, new SecureRandom()));
            signer.BlockUpdate(digest, 0, digest.Length);
            return Task.FromResult(signer.GenerateSignature());
        });
    }
}
