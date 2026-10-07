using System.IO.Compression;
using System.Text;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.GM;
using Org.BouncyCastle.Asn1.Sec;
using Org.BouncyCastle.Asn1.X509;
using Org.BouncyCastle.Crypto.Generators;
using Org.BouncyCastle.Crypto.Operators;
using Org.BouncyCastle.Crypto.Parameters;
using Org.BouncyCastle.Math;
using Org.BouncyCastle.Security;
using Org.BouncyCastle.X509;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.SesInterop.Tests;

public sealed class SesTests
{
    private const string Property = "/Doc_0/Signs/Sign_0/Signature.xml";
    private static byte[] Xml(string extra = "") => Encoding.UTF8.GetBytes(
        "<Signature xmlns=\"http://www.ofdspec.org/2016\"><SignedInfo>" +
        "<SignatureMethod>1.2.156.10197.1.501</SignatureMethod>" +
        "<SignatureDateTime>2026-10-07T12:00:00Z</SignatureDateTime>" + extra + "</SignedInfo></Signature>");

    [Theory]
    [InlineData(SesVersion.V1)]
    [InlineData(SesVersion.V4)]
    public async Task ExplicitRegistrationVerifiesWhileDefaultOnlyVerifiesReferences(SesVersion version)
    {
        var identity = SesTestIdentity.Generate();
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdTextElement { Text = "SES test 测试", FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 10 } } });
        using var source = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, source); source.Position = 0;
        using var signed = new MemoryStream();
        await new OfdSignatureService().SignAsync(source, signed, new SesTestSignatureProvider(identity, version));
        signed.Position = 0; var ordinary = await new OfdSignatureVerifier().VerifyAsync(signed);
        Assert.True(ordinary.ReferenceIntegrityValid); Assert.False(ordinary.FullyValid);
        Assert.Equal(OfdCryptographicVerificationStatus.Unsupported, Assert.Single(ordinary.Signatures).CryptographicStatus);
        signed.Position = 0; var explicitReport = await new OfdSignatureVerifier(
            new[] { new SesTestSignedValueVerifier(identity.CertificateDer) }).VerifyAsync(signed);
        Assert.True(explicitReport.FullyValid);
        Assert.False(OfdCryptographicCapabilities.SupportsBuiltInSesSm2Verification);
        Assert.True(SesInteropCapabilities.SupportsExplicitTestSelfSigning);
        signed.Position = 0; using var zip = new ZipArchive(signed, ZipArchiveMode.Read, true);
        var value = zip.Entries.Single(e => e.FullName.EndsWith("SignedValue.dat"));
        using var entry = value.Open(); using var bytes = new MemoryStream(); await entry.CopyToAsync(bytes);
        var parsed = SesSignedValueReader.Parse(bytes.ToArray());
        Assert.Equal(version, parsed.Version); Assert.Equal("Ofdrw.Net.DEV", parsed.Seal.VendorId);
        Assert.Equal(Property, parsed.PropertyInformation); Assert.Equal(identity.CertificateDer, parsed.CertificateDer);
        Assert.Null(parsed.Timestamp);
        Assert.Equal(32, parsed.DataHash.Length);
        // Preserve the original page bytes when signing, independently of reference checks.
        source.Position = 0; using var original = new ZipArchive(source, ZipArchiveMode.Read, true);
        using var originalPage = original.GetEntry("Doc_0/Pages/Page_0/Content.xml")!.Open();
        using var signedPage = zip.GetEntry("Doc_0/Pages/Page_0/Content.xml")!.Open();
        using var before = new MemoryStream(); using var after = new MemoryStream();
        originalPage.CopyTo(before); signedPage.CopyTo(after); Assert.Equal(before.ToArray(), after.ToArray());
    }

    [Fact]
    public void ParsesExistingUpstreamV4WithoutTrustingIt()
    {
        var root = RepositoryRoot();
        using var zip = ZipFile.OpenRead(Path.Combine(root, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/upstream-ofdrw/999.ofd"));
        using var source = zip.GetEntry("Doc_0/Signs/Sign_0/SignedValue.dat")!.Open();
        using var target = new MemoryStream(); source.CopyTo(target); var bytes = target.ToArray();
        Assert.Equal("DC10A090EA5F40F7D7F8FF6B057EB4BBD185DFA386D7842506A339671F8E82D3",
            Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)));
        var value = SesSignedValueReader.Parse(bytes);
        Assert.Equal(SesVersion.V4, value.Version); Assert.Equal("GOMAIN", value.Seal.VendorId);
        Assert.Equal("50011200000323", value.Seal.SealId); Assert.Equal("ofd", value.Seal.PictureType);
        Assert.Equal(30, value.Seal.WidthMillimeters); Assert.Equal(20, value.Seal.HeightMillimeters);
        Assert.Equal("20200817111329Z", value.TimeText); Assert.Equal(Property, value.PropertyInformation);
        Assert.Equal("1.2.156.10197.1.501", value.SignatureAlgorithm);
        Assert.ThrowsAny<Exception>(() => new SesTestSignedValueVerifier(value.CertificateDer));
        var copy = value.DataHash; copy[0] ^= 1; Assert.NotEqual(copy, value.DataHash);
    }

    [Fact]
    public void PublishedSm2VectorAndIdentityBinding()
    {
        // BC release-2.6.2 SM2SignerTest.cs DoSignerTestFpStandardSM3, independent r/s known answer.
        var d = new BigInteger("110E7973206F68C19EE5F7328C036F26911C8C73B4E4F36AE3291097F8984FFC", 16);
        var key = new ECPublicKeyParameters(SesTestCrypto.Domain.G.Multiply(d), SesTestCrypto.Domain);
        var signature = new DerSequence(
            new DerInteger(new BigInteger("05890B9077B92E47B17A1FF42A814280E556AFD92B4A98B9670BF8B1A274C2FA", 16)),
            new DerInteger(new BigInteger("E3ABBB8DB2B6ECD9B24ECCEA7F679FB9A4B1DB52F4AA985E443AD73237FA1993", 16))).GetEncoded("DER");
        var id = Encoding.ASCII.GetBytes("sm2test@example.com"); var message = Encoding.ASCII.GetBytes("hi chappy");
        Assert.True(SesTestCrypto.Verify(message, signature, key, id));
        Assert.False(SesTestCrypto.Verify(message, signature, key));
        Assert.False(SesTestCrypto.Verify(Encoding.ASCII.GetBytes("hi chappY"), signature, key, id));
    }

    [Theory]
    [InlineData(SesVersion.V1)]
    [InlineData(SesVersion.V4)]
    public async Task RejectsWrongIdentityXmlAndPath(SesVersion version)
    {
        var identity = SesTestIdentity.Generate(); var value = await new SesTestSignatureProvider(identity, version).SignAsync(Xml(), Property);
        var verifier = new SesTestSignedValueVerifier(identity.CertificateDer);
        Assert.False(await new SesTestSignedValueVerifier(SesTestIdentity.Generate().CertificateDer).VerifyAsync(Xml(), value, Property));
        Assert.False(await verifier.VerifyAsync(Xml("<Extra/>"), value, Property));
        Assert.False(await verifier.VerifyAsync(Xml(), value, "/Doc_0/Signs/Sign_1/Signature.xml"));
        var clone = identity.CertificateDer; clone[0] = 0; Assert.True(await verifier.VerifyAsync(Xml(), value, Property));
    }

    [Theory]
    [InlineData(SesVersion.V1, "time")]
    [InlineData(SesVersion.V4, "time")]
    [InlineData(SesVersion.V1, "path")]
    [InlineData(SesVersion.V4, "path")]
    [InlineData(SesVersion.V1, "hash")]
    [InlineData(SesVersion.V4, "hash")]
    [InlineData(SesVersion.V1, "seal-signature")]
    [InlineData(SesVersion.V4, "seal-signature")]
    [InlineData(SesVersion.V1, "seal-name")]
    [InlineData(SesVersion.V4, "seal-name")]
    [InlineData(SesVersion.V1, "certificate")]
    [InlineData(SesVersion.V4, "certificate")]
    [InlineData(SesVersion.V1, "algorithm")]
    [InlineData(SesVersion.V4, "algorithm")]
    public async Task RejectsResignedIllegalMetadata(SesVersion version, string mutation)
    {
        var identity = SesTestIdentity.Generate(); var original = await new SesTestSignatureProvider(identity, version).SignAsync(Xml(), Property);
        var outer = Asn1Sequence.GetInstance(Asn1Object.FromByteArray(original)); var tbs = Asn1Sequence.GetInstance(outer[0]);
        switch (mutation)
        {
            case "time": tbs = Replace(tbs, 2, SesTestProfile.Time(version, new DateTime(2026, 10, 8, 12, 0, 0, DateTimeKind.Utc))); break;
            case "path": tbs = Replace(tbs, 4, new DerIA5String("/Doc_0/Signs/Sign_1/Signature.xml")); break;
            case "hash": tbs = Replace(tbs, 3, new DerBitString(new byte[32])); break;
            case "certificate":
                // Same public key, same subject, different certificate serial: pin complete DER, not just Q.
                var cert = new DerOctetString(AlternateCertificate(identity));
                if (version == SesVersion.V1) tbs = Replace(tbs, 5, cert); else outer = Replace(outer, 1, cert);
                break;
            case "algorithm":
                var oid = new DerObjectIdentifier("1.2.840.10045.4.3.2");
                if (version == SesVersion.V1) tbs = Replace(tbs, 6, oid); else outer = Replace(outer, 2, oid);
                break;
            default:
                var seal = Asn1Sequence.GetInstance(tbs[1]);
                if (mutation == "seal-signature")
                {
                    if (version == SesVersion.V1)
                        seal = Replace(seal, 1, Replace(Asn1Sequence.GetInstance(seal[1]), 2, new DerBitString(new byte[] { 1 })));
                    else seal = Replace(seal, 3, new DerBitString(new byte[] { 1 }));
                }
                else
                {
                    var info = Asn1Sequence.GetInstance(seal[0]); var properties = Asn1Sequence.GetInstance(info[2]);
                    info = Replace(info, 2, Replace(properties, 1, new DerUtf8String("unapproved seal name")));
                    seal = Replace(seal, 0, info);
                    var maker = version == SesVersion.V1 ? Asn1Sequence.GetInstance(seal[1]) : seal;
                    var makerData = version == SesVersion.V1 ? new DerSequence(info, maker[0], maker[1]).GetEncoded("DER") : info.GetEncoded("DER");
                    var signature = new DerBitString(SesTestCrypto.Sign(makerData, identity.PrivateKey));
                    if (version == SesVersion.V1) seal = Replace(seal, 1, Replace(maker, 2, signature)); else seal = Replace(seal, 3, signature);
                }
                tbs = Replace(tbs, 1, seal); break;
        }
        var signed = new DerBitString(SesTestCrypto.Sign(tbs.GetEncoded("DER"), identity.PrivateKey));
        outer = Replace(Replace(outer, 0, tbs), version == SesVersion.V1 ? 1 : 3, signed);
        Assert.False(await new SesTestSignedValueVerifier(identity.CertificateDer).VerifyAsync(Xml(), outer.GetEncoded("DER"), Property));
    }

    [Fact]
    public async Task OptionalExtensionsAndTimestampAreInspectionOnly()
    {
        var identity = SesTestIdentity.Generate(); var bytes = await new SesTestSignatureProvider(identity, SesVersion.V4).SignAsync(Xml(), Property);
        var outer = Asn1Sequence.GetInstance(Asn1Object.FromByteArray(bytes)); var verifier = new SesTestSignedValueVerifier(identity.CertificateDer);
        var timestamp = new DerSequence(outer.Cast<Asn1Encodable>().Concat(new[] { new DerTaggedObject(true, 0, new DerBitString(new byte[] { 1, 2, 3 })) }).ToArray()).GetEncoded("DER");
        Assert.Equal(new byte[] { 1, 2, 3 }, SesSignedValueReader.Parse(timestamp).Timestamp);
        Assert.False(await verifier.VerifyAsync(Xml(), timestamp, Property));
        var tbs = Asn1Sequence.GetInstance(outer[0]);
        var extended = new DerSequence(tbs.Cast<Asn1Encodable>().Concat(new[] { new DerTaggedObject(true, 0, new DerSequence()) }).ToArray());
        var withExtensions = Replace(Replace(outer, 0, extended), 3, new DerBitString(SesTestCrypto.Sign(extended.GetEncoded("DER"), identity.PrivateKey))).GetEncoded("DER");
        Assert.Equal(SesVersion.V4, SesSignedValueReader.Parse(withExtensions).Version);
        Assert.False(await verifier.VerifyAsync(Xml(), withExtensions, Property));
    }

    [Theory]
    [InlineData("<SignatureDateTime>2026-10-07T12:00:00Z</SignatureDateTime>")]
    [InlineData("<x:SignatureDateTime xmlns:x='urn:wrong'>anything</x:SignatureDateTime>")]
    [InlineData("<Wrapper><SignatureMethod>1.2.156.10197.1.501</SignatureMethod></Wrapper>")]
    public async Task RejectsAmbiguousXml(string extra)
    {
        var identity = SesTestIdentity.Generate(); var provider = new SesTestSignatureProvider(identity, SesVersion.V4);
        var value = await provider.SignAsync(Xml(), Property);
        await Assert.ThrowsAsync<InvalidDataException>(() => provider.SignAsync(Xml(extra), Property));
        Assert.False(await new SesTestSignedValueVerifier(identity.CertificateDer).VerifyAsync(Xml(extra), value, Property));
    }

    [Fact]
    public async Task RejectsDtdWrongNamespaceTimeAndUnsafePath()
    {
        var identity = SesTestIdentity.Generate(); var provider = new SesTestSignatureProvider(identity, SesVersion.V1);
        var original = await provider.SignAsync(Xml(), Property); var verifier = new SesTestSignedValueVerifier(identity.CertificateDer);
        foreach (var xml in new[] {
            Encoding.UTF8.GetBytes("<!DOCTYPE Signature [<!ENTITY x 'x'>]>" + Encoding.UTF8.GetString(Xml())),
            Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(Xml()).Replace("http://www.ofdspec.org/2016", "urn:wrong")),
            Encoding.UTF8.GetBytes(Encoding.UTF8.GetString(Xml()).Replace("2026-10-07T12:00:00Z", "2026-10-07T12:00:00+08:00")),
            new byte[1024 * 1024 + 1] }) Assert.False(await verifier.VerifyAsync(xml, original, Property));
        foreach (var path in new[] { "Doc_0/Signature.xml", "/Doc_0/../Signature.xml", "/Doc_0//Signature.xml", "/中文/Signature.xml", "/Doc_0/signature.xml" })
            await Assert.ThrowsAsync<InvalidDataException>(() => provider.SignAsync(Xml(), path));
    }

    [Fact]
    public async Task CancellationAndNullProgrammerInputsPropagate()
    {
        var identity = SesTestIdentity.Generate(); var provider = new SesTestSignatureProvider(identity, SesVersion.V4);
        var verifier = new SesTestSignedValueVerifier(identity.CertificateDer); var token = new CancellationToken(true);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => provider.SignAsync(Xml(), Property, token));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => verifier.VerifyAsync(Xml(), new byte[1], Property, token));
        Assert.Throws<ArgumentNullException>(() => SesSignedValueReader.Parse(null!));
        Assert.Throws<ArgumentNullException>(() => new SesTestSignedValueVerifier(null!));
        Assert.Throws<ArgumentOutOfRangeException>(() => new SesTestSignatureProvider(identity, (SesVersion)2));
        await Assert.ThrowsAsync<ArgumentNullException>(() => provider.SignAsync(null!, Property));
        await Assert.ThrowsAsync<ArgumentNullException>(() => verifier.VerifyAsync(Xml(), null!, Property));
    }

    [Fact]
    public void RejectsMalformedDerBeforeRecursiveLibraryParsing()
    {
        foreach (var bytes in new[] { Array.Empty<byte>(), new byte[] { 0x30, 0x80, 0, 0 },
            new byte[] { 0x30, 0x81, 0 }, new byte[] { 0x30, 0x82, 0, 128 },
            new byte[] { 0x30, 0, 0x30, 0 }, new byte[] { 0x30, 0x83, 255, 255, 255 },
            new byte[SesSignedValueReader.MaximumSignedValueBytes + 1] })
            Assert.Throws<InvalidDataException>(() => SesSignedValueReader.Parse(bytes));
        Asn1Encodable deep = new DerInteger(1);
        for (var i = 0; i < 40; i++) deep = new DerSequence(deep);
        Assert.Throws<InvalidDataException>(() => SesSignedValueReader.Parse(deep.GetEncoded("DER")));
        Assert.Throws<InvalidDataException>(() => SesSignedValueReader.Parse(
            new DerSequence(Enumerable.Repeat<Asn1Encodable>(new DerInteger(1), 4100).ToArray()).GetEncoded("DER")));
    }

    [Fact]
    public void RejectsNonSm2CertificateEvenIfItIsASelfSignedEcCertificate()
    {
        var generator = new ECKeyPairGenerator(); generator.Init(new ECKeyGenerationParameters(SecObjectIdentifiers.SecP256r1, new SecureRandom()));
        var pair = generator.GenerateKeyPair(); var certificate = new X509V3CertificateGenerator();
        var name = new X509Name("CN=OFD SES DEV INTEROP TEST");
        certificate.SetSerialNumber(BigInteger.One); certificate.SetIssuerDN(name); certificate.SetSubjectDN(name);
        certificate.SetNotBefore(SesTestProfile.Begin); certificate.SetNotAfter(SesTestProfile.End); certificate.SetPublicKey(pair.Public);
        var bytes = certificate.Generate(new Asn1SignatureFactory("SHA256WITHECDSA", pair.Private)).GetEncoded();
        Assert.Throws<InvalidDataException>(() => new SesTestSignedValueVerifier(bytes));
    }

    [Fact]
    public void DerLimitsAcceptBoundaryAndRejectNextNodeDepthOrByte()
    {
        Asn1Encodable nested = new DerInteger(1);
        for (var i = 0; i < 31; i++) nested = new DerSequence(nested);
        _ = SesDer.Read(nested.GetEncoded("DER"));
        Assert.Throws<InvalidDataException>(() => SesDer.Read(new DerSequence(nested).GetEncoded("DER")));
        var nodes = new DerSequence(Enumerable.Repeat<Asn1Encodable>(new DerInteger(1), 4095).ToArray());
        _ = SesDer.Read(nodes.GetEncoded("DER"));
        Assert.Throws<InvalidDataException>(() => SesDer.Read(new DerSequence(
            Enumerable.Repeat<Asn1Encodable>(new DerInteger(1), 4096).ToArray()).GetEncoded("DER")));
        var bytes = new DerOctetString(new byte[SesDer.MaximumBytes - 5]).GetEncoded("DER");
        Assert.Equal(SesDer.MaximumBytes, bytes.Length); _ = SesDer.Read(bytes);
        Assert.Throws<InvalidDataException>(() => SesDer.Read(new DerOctetString(new byte[SesDer.MaximumBytes - 4]).GetEncoded("DER")));
    }

    [Fact]
    public async Task XmlAndPropertyLimitsAcceptBoundaryAndRejectNextByte()
    {
        var identity = SesTestIdentity.Generate(); var provider = new SesTestSignatureProvider(identity, SesVersion.V4);
        var prefix = "/Doc_0/"; var suffix = "/Signature.xml";
        var property = prefix + new string('a', 1024 - prefix.Length - suffix.Length) + suffix;
        var xml = Xml().Concat(Enumerable.Repeat((byte)' ', 1024 * 1024 - Xml().Length)).ToArray();
        var value = await provider.SignAsync(xml, property);
        Assert.True(await new SesTestSignedValueVerifier(identity.CertificateDer).VerifyAsync(xml, value, property));
        await Assert.ThrowsAsync<InvalidDataException>(() => provider.SignAsync(xml.Concat(new[] { (byte)' ' }).ToArray(), property));
        await Assert.ThrowsAsync<InvalidDataException>(() => provider.SignAsync(xml, prefix + "a" + property.Substring(prefix.Length)));
    }

    [Theory]
    [InlineData(SesVersion.V1)]
    [InlineData(SesVersion.V4)]
    public async Task RejectsWrongSm2IdAndZeroOrOutOfRangeSignatures(SesVersion version)
    {
        var identity = SesTestIdentity.Generate(); var bytes = await new SesTestSignatureProvider(identity, version).SignAsync(Xml(), Property);
        var outer = Asn1Sequence.GetInstance(Asn1Object.FromByteArray(bytes)); var tbs = outer[0].GetEncoded("DER");
        var signer = new Org.BouncyCastle.Crypto.Signers.SM2Signer();
        signer.Init(true, new ParametersWithID(new ParametersWithRandom(identity.PrivateKey, new SecureRandom()), Encoding.ASCII.GetBytes("wrong-test-id")));
        signer.BlockUpdate(tbs, 0, tbs.Length);
        var verifier = new SesTestSignedValueVerifier(identity.CertificateDer);
        foreach (var signature in new[] { signer.GenerateSignature(),
            new DerSequence(new DerInteger(0), new DerInteger(1)).GetEncoded("DER"),
            new DerSequence(new DerInteger(SesTestCrypto.Domain.N), new DerInteger(1)).GetEncoded("DER"), new byte[64] })
            Assert.False(await verifier.VerifyAsync(Xml(), Replace(outer, version == SesVersion.V1 ? 1 : 3, new DerBitString(signature)).GetEncoded("DER"), Property));
    }

    [Theory]
    [InlineData(SesVersion.V1)]
    [InlineData(SesVersion.V4)]
    public async Task RejectsWrongHeaderDimensionsPaddingAndIntegerOverflow(SesVersion version)
    {
        var identity = SesTestIdentity.Generate(); var bytes = await new SesTestSignatureProvider(identity, version).SignAsync(Xml(), Property);
        var outer = Asn1Sequence.GetInstance(Asn1Object.FromByteArray(bytes)); var tbs = Asn1Sequence.GetInstance(outer[0]);
        var seal = Asn1Sequence.GetInstance(tbs[1]); var info = Asn1Sequence.GetInstance(seal[0]);
        var header = Asn1Sequence.GetInstance(info[0]); var picture = Asn1Sequence.GetInstance(info[3]);
        var verifier = new SesTestSignedValueVerifier(identity.CertificateDer);
        foreach (var altered in new Asn1Encodable[] {
            Replace(info, 0, Replace(header, 1, new DerInteger(version == SesVersion.V1 ? 4 : 1))),
            Replace(info, 3, Replace(picture, 2, new DerInteger(0))),
            Replace(info, 3, Replace(picture, 3, new DerInteger(-1))),
            Replace(info, 3, Replace(picture, 2, new DerInteger(BigInteger.One.ShiftLeft(80)))) })
        {
            var malformed = Replace(outer, 0, Replace(tbs, 1, Replace(seal, 0, altered))).GetEncoded("DER");
            Assert.Throws<InvalidDataException>(() => SesSignedValueReader.Parse(malformed));
            Assert.False(await verifier.VerifyAsync(Xml(), malformed, Property));
        }
        var padding = Replace(outer, 0, Replace(tbs, 3, new DerBitString(new byte[32], 1))).GetEncoded("DER");
        Assert.Throws<InvalidDataException>(() => SesSignedValueReader.Parse(padding));
        Assert.False(await verifier.VerifyAsync(Xml(), padding, Property));
    }

    [Fact]
    public async Task V4CertificateDigestListCanBeInspectedButCannotPassTestProfile()
    {
        var identity = SesTestIdentity.Generate(); var bytes = await new SesTestSignatureProvider(identity, SesVersion.V4).SignAsync(Xml(), Property);
        var outer = Asn1Sequence.GetInstance(Asn1Object.FromByteArray(bytes)); var tbs = Asn1Sequence.GetInstance(outer[0]);
        var seal = Asn1Sequence.GetInstance(tbs[1]); var info = Asn1Sequence.GetInstance(seal[0]); var properties = Asn1Sequence.GetInstance(info[2]);
        var digestList = new DerSequence((Asn1Encodable)new DerSequence(new DerPrintableString("SM3"), new DerOctetString(new byte[32])));
        properties = Replace(Replace(properties, 2, new DerInteger(2)), 3, digestList);
        var value = Replace(outer, 0, Replace(tbs, 1, Replace(seal, 0, Replace(info, 2, properties)))).GetEncoded("DER");
        Assert.Equal(SesVersion.V4, SesSignedValueReader.Parse(value).Version);
        Assert.False(await new SesTestSignedValueVerifier(identity.CertificateDer).VerifyAsync(Xml(), value, Property));
    }

    private static byte[] AlternateCertificate(SesTestIdentity identity)
    {
        var certificate = new X509V3CertificateGenerator(); var name = new X509Name("CN=OFD SES DEV INTEROP TEST");
        certificate.SetSerialNumber(BigInteger.One); certificate.SetIssuerDN(name); certificate.SetSubjectDN(name);
        certificate.SetNotBefore(SesTestProfile.Begin); certificate.SetNotAfter(SesTestProfile.End);
        certificate.SetPublicKey(new ECPublicKeyParameters(SesTestCrypto.Domain.G.Multiply(identity.PrivateKey.D), SesTestCrypto.Domain));
        return certificate.Generate(new Asn1SignatureFactory("SM3WITHSM2", identity.PrivateKey)).GetEncoded();
    }
    private static DerSequence Replace(Asn1Sequence sequence, int index, Asn1Encodable value) =>
        new(sequence.Cast<Asn1Encodable>().Select((v, i) => i == index ? value : v).ToArray());
    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException();
    }
}
