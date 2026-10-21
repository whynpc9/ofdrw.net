using System.Text;
using Org.BouncyCastle.Asn1;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Signatures.SesInterop.Tests;

public sealed class SesV5AndGbtTests
{
    private const string Property = "/Doc_0/Signs/Sign_0/Signature.xml";
    private static readonly byte[] Xml = Encoding.UTF8.GetBytes(
        "<Signature xmlns=\"http://www.ofdspec.org/2016\"><SignedInfo>" +
        "<SignatureMethod>1.2.156.10197.1.501</SignatureMethod>" +
        "<SignatureDateTime>2026-10-21T12:00:00Z</SignatureDateTime>" +
        "</SignedInfo></Signature>");
    private static byte[] JavaLocalXml => Encoding.UTF8.GetBytes(
        Encoding.UTF8.GetString(Xml).Replace("2026-10-21T12:00:00Z", "2026-10-21 20:00:00"));

    [Fact]
    public async Task V5ExternalIdentitySealAndTimestampBoundaries()
    {
        var maker = SesTestIdentity.Generate(); var signer = SesTestIdentity.Generate();
        var seal = Seal(maker, signer.CertificateDer);
        var provider = new SesV5ExternalSignatureProvider(signer.CertificateDer, seal, new Signer(signer));
        var value = await provider.SignAsync(Xml, Property);
        Assert.Equal(SesVersion.V5, SesSignedValueReader.Parse(value).Version);
        var verifier = new SesV5PinnedVerifier(signer.CertificateDer, maker.CertificateDer);
        Assert.True(await verifier.VerifyAsync(Xml, value, Property));
        var outer = SesDer.Sequence(SesDer.Read(value), 4);
        var originalTbs = SesDer.Sequence(outer[0], 5);
        var differentTimeTbs = new DerSequence(originalTbs[0], originalTbs[1],
            new DerGeneralizedTime(new DateTime(2026, 10, 21, 12, 0, 1, DateTimeKind.Utc)),
            originalTbs[3], originalTbs[4]);
        var differentTimeValue = new DerSequence(differentTimeTbs, outer[1], outer[2],
            new DerBitString(SesTestCrypto.Sign(differentTimeTbs.GetEncoded("DER"), signer.PrivateKey)))
            .GetEncoded("DER");
        Assert.False(await verifier.VerifyAsync(Xml, differentTimeValue, Property));
        Assert.False(await new SesV5PinnedVerifier(SesTestIdentity.Generate().CertificateDer, maker.CertificateDer)
            .VerifyAsync(Xml, value, Property));
        Assert.False(await new SesV5PinnedVerifier(signer.CertificateDer, SesTestIdentity.Generate().CertificateDer)
            .VerifyAsync(Xml, value, Property));
        Assert.False(await verifier.VerifyAsync(Xml.Concat(new byte[] { 32 }).ToArray(), value, Property));
        Assert.False(await verifier.VerifyAsync(Xml, value, "/Doc_0/Signs/Sign_1/Signature.xml"));
        var localTimeValue = await provider.SignAsync(JavaLocalXml, Property);
        Assert.True(await verifier.VerifyAsync(JavaLocalXml, localTimeValue, Property));
        var tampered = (byte[])value.Clone(); tampered[tampered.Length - 3] ^= 1;
        Assert.False(await verifier.VerifyAsync(Xml, tampered, Property));
        Assert.Throws<InvalidDataException>(() => new SesV5ExternalSignatureProvider(
            SesTestIdentity.Generate().CertificateDer, seal, new Signer(signer)));
        var corruptedSeal = (byte[])seal.Clone(); corruptedSeal[corruptedSeal.Length - 3] ^= 1;
        Assert.Throws<InvalidDataException>(() => new SesV5ExternalSignatureProvider(
            signer.CertificateDer, corruptedSeal, new Signer(signer)));
        await Assert.ThrowsAsync<InvalidDataException>(() =>
            new SesV5ExternalSignatureProvider(signer.CertificateDer, seal, new Signer(maker)).SignAsync(Xml, Property));
        var expiredSigner = new Signer(signer);
        var expiredProvider = new SesV5ExternalSignatureProvider(signer.CertificateDer,
            Seal(maker, signer.CertificateDer, validUntil: new DateTime(2025, 1, 1, 0, 0, 0,
                DateTimeKind.Utc)), expiredSigner);
        await Assert.ThrowsAsync<InvalidDataException>(() => expiredProvider.SignAsync(Xml, Property));
        Assert.Equal(0, expiredSigner.Calls);
        var futureSigner = new Signer(signer);
        var futureProvider = new SesV5ExternalSignatureProvider(signer.CertificateDer,
            Seal(maker, signer.CertificateDer, validFrom: new DateTime(2027, 1, 1, 0, 0, 0,
                DateTimeKind.Utc)), futureSigner);
        await Assert.ThrowsAsync<InvalidDataException>(() => futureProvider.SignAsync(Xml, Property));
        Assert.Equal(0, futureSigner.Calls);

        byte[]? seenSignature = null;
        var withTimestamp = new SesV5ExternalSignatureProvider(signer.CertificateDer, seal, new Signer(signer),
            (signature, _) => { seenSignature = signature; return Task.FromResult<byte[]?>(new byte[] { 1, 2, 3 }); });
        var stamped = await withTimestamp.SignAsync(Xml, Property);
        Assert.NotNull(seenSignature);
        Assert.Equal(new byte[] { 1, 2, 3 }, SesSignedValueReader.Parse(stamped).Timestamp);
        Assert.False(await verifier.VerifyAsync(Xml, stamped, Property)); // No TSA verification is installed.
        await Assert.ThrowsAsync<InvalidOperationException>(() =>
            new SesV5ExternalSignatureProvider(signer.CertificateDer, seal, new Signer(signer),
                (_, _) => throw new InvalidOperationException("timestamp service failed")).SignAsync(Xml, Property));
    }

    [Theory]
    [InlineData(Gbt35275PayloadProfile.RawSm3Digest)]
    [InlineData(Gbt35275PayloadProfile.Base64Sm3DigestCompatibility)]
    public async Task Gbt35275ExplicitProfileAndContainerDispatch(Gbt35275PayloadProfile profile)
    {
        var identity = SesTestIdentity.Generate();
        var value = await new Gbt35275SignatureProvider(identity.CertificateDer, new Signer(identity), profile)
            .SignAsync(Xml, Property);
        var parsed = Gbt35275SignedData.Read(value);
        Assert.Equal(identity.CertificateDer, parsed.CertificateDer);
        var verifier = new Gbt35275PinnedVerifier(identity.CertificateDer, profile);
        Assert.True(await verifier.VerifyAsync(Xml, value, Property));
        Assert.False(await new Gbt35275PinnedVerifier(identity.CertificateDer,
            profile == Gbt35275PayloadProfile.RawSm3Digest
                ? Gbt35275PayloadProfile.Base64Sm3DigestCompatibility : Gbt35275PayloadProfile.RawSm3Digest)
            .VerifyAsync(Xml, value, Property));
        Assert.False(await new Gbt35275PinnedVerifier(SesTestIdentity.Generate().CertificateDer, profile)
            .VerifyAsync(Xml, value, Property));
        Assert.False(await verifier.VerifyAsync(Xml.Concat(new byte[] { 32 }).ToArray(), value, Property));
        var javaTimeValue = await new Gbt35275SignatureProvider(identity.CertificateDer, new Signer(identity), profile)
            .SignAsync(JavaLocalXml, Property);
        Assert.True(await verifier.VerifyAsync(JavaLocalXml, javaTimeValue, Property));
        var wrongType = (byte[])value.Clone(); wrongType[12] ^= 1;
        Assert.False(await verifier.VerifyAsync(Xml, wrongType, Property));
        Assert.True(await new SesSm2ContainerVerifier(signedData: verifier).VerifyAsync(Xml, value, Property));
        Assert.False(await new SesSm2ContainerVerifier(v5: new SesV5PinnedVerifier(
            identity.CertificateDer, identity.CertificateDer)).VerifyAsync(Xml, value, Property));
    }

    [Fact]
    public async Task BothFormatsRemainOptInAtOfdPackageBoundary()
    {
        var maker = SesTestIdentity.Generate(); var signer = SesTestIdentity.Generate();
        var v5 = new SesV5ExternalSignatureProvider(signer.CertificateDer,
            Seal(maker, signer.CertificateDer), new Signer(signer));
        var gbt = new Gbt35275SignatureProvider(signer.CertificateDer, new Signer(signer));
        var composite = new SesSm2ContainerVerifier(
            new SesV5PinnedVerifier(signer.CertificateDer, maker.CertificateDer),
            new Gbt35275PinnedVerifier(signer.CertificateDer));
        var document = new OfdDocumentPackage();
        document.Pages.Add(new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdTextElement { Text = "SES V5 测试", FontName = "Arial",
                FontSizeMillimeters = 4, XMillimeters = 10, YMillimeters = 10,
                WidthMillimeters = 80, HeightMillimeters = 10 } } });
        using var unsigned = new MemoryStream(); await new OfdPackageWriter().WriteAsync(document, unsigned);
        foreach (var provider in new IOfdSignatureProvider[] { v5, gbt })
        {
            unsigned.Position = 0; using var signed = new MemoryStream();
            await new OfdSignatureService().SignAsync(unsigned, signed, provider);
            signed.Position = 0; var ordinary = await new OfdSignatureVerifier().VerifyAsync(signed);
            Assert.True(ordinary.ReferenceIntegrityValid);
            Assert.False(ordinary.FullyValid);
            Assert.Equal(OfdCryptographicVerificationStatus.Unsupported,
                Assert.Single(ordinary.Signatures).CryptographicStatus);
            signed.Position = 0; var explicitReport = await new OfdSignatureVerifier(new[] { composite }).VerifyAsync(signed);
            Assert.True(explicitReport.FullyValid);
        }
        Assert.False(OfdCryptographicCapabilities.SupportsBuiltInSesSm2Verification);
        Assert.True(SesInteropCapabilities.SupportsExplicitExternalV5);
        Assert.True(SesInteropCapabilities.SupportsExplicitGbt35275SingleSigner);
    }

    private sealed class Signer(SesTestIdentity identity) : ISesExternalSm2Signer
    {
        public int Calls { get; private set; }
        public Task<byte[]> SignAsync(byte[] data, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Calls++;
            return Task.FromResult(SesTestCrypto.Sign(data, identity.PrivateKey));
        }
    }

    private static byte[] Seal(SesTestIdentity maker, byte[] authorizedSignerCertificate,
        DateTime? validFrom = null, DateTime? validUntil = null)
    {
        var v4Info = SesDer.Sequence(SesTestProfile.SealInfo(SesVersion.V4, authorizedSignerCertificate), 4);
        var originalProperties = SesDer.Sequence(v4Info[2], 7);
        var properties = new DerSequence(originalProperties[0], originalProperties[1],
            originalProperties[2], originalProperties[3], originalProperties[4],
            new DerGeneralizedTime(validFrom ?? SesTestProfile.Begin),
            new DerGeneralizedTime(validUntil ?? SesTestProfile.End));
        var info = new DerSequence(
            new DerSequence(new DerIA5String("ES"), new DerInteger(5), new DerIA5String("Ofdrw.Net.external-test")),
            v4Info[1], properties, v4Info[3]);
        return new DerSequence(info, new DerOctetString(maker.CertificateDer),
            new DerObjectIdentifier(SesTestCrypto.Algorithm),
            new DerBitString(SesTestCrypto.Sign(info.GetEncoded("DER"), maker.PrivateKey))).GetEncoded("DER");
    }
}
