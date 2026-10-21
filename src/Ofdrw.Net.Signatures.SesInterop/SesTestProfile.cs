using System;
using System.IO;
using System.Text;
using System.Globalization;
using Org.BouncyCastle.Asn1;
using Ofdrw.Net.Signatures.Crypto;

namespace Ofdrw.Net.Signatures.SesInterop;

internal static class SesTestProfile
{
    internal static readonly DateTime Begin = new(2020, 1, 1, 0, 0, 0, DateTimeKind.Utc);
    internal static readonly DateTime End = new(2049, 12, 31, 23, 59, 59, DateTimeKind.Utc);
    private static readonly byte[] Picture = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");

    internal static void Version(SesVersion version)
    {
        if (version != SesVersion.V1 && version != SesVersion.V4) throw new ArgumentOutOfRangeException(nameof(version));
    }
    internal static DateTime ValidateInput(byte[] xml, string property)
    {
        var time = SesSignatureInput.Validate(xml, property);
        if (time < Begin || time >= End)
            throw new InvalidDataException("Expected UTC second-resolution test signature time within the test validity interval.");
        return time;
    }
    internal static Asn1Encodable Time(SesVersion version, DateTime time) => version == SesVersion.V1
        ? new DerBitString(Encoding.UTF8.GetBytes(time.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)))
        : new DerGeneralizedTime(time);

    internal static DerSequence SealInfo(SesVersion version, byte[] certificate)
    {
        var properties = new Asn1EncodableVector(new DerInteger(0), new DerUtf8String("OFD SES DEV INTEROP TEST"));
        if (version == SesVersion.V4) properties.Add(new DerInteger(1));
        Asn1Encodable SealTime(DateTime time) => version == SesVersion.V1
            ? new DerUtcTime(time.ToString("yyMMddHHmmss'Z'", CultureInfo.InvariantCulture)) : new DerGeneralizedTime(time);
        properties.Add(new DerSequence(new DerOctetString(certificate)),
            SealTime(Begin), SealTime(Begin), SealTime(End));
        return new DerSequence(
            new DerSequence(new DerIA5String("ES"), new DerInteger((int)version), new DerIA5String("Ofdrw.Net.DEV")),
            new DerIA5String("dev-interop-test"), new DerSequence(properties),
            new DerSequence(new DerIA5String("PNG"), new DerOctetString(Picture), new DerInteger(1), new DerInteger(1)));
    }
    internal static byte[] Seal(SesVersion version, SesTestIdentity identity)
    {
        var info = SealInfo(version, identity.CertificateDer);
        var cert = new DerOctetString(identity.CertificateDer);
        var oid = new DerObjectIdentifier(SesTestCrypto.Algorithm);
        var data = version == SesVersion.V1 ? new DerSequence(info, cert, oid).GetEncoded("DER") : info.GetEncoded("DER");
        var signature = new DerBitString(SesTestCrypto.Sign(data, identity.PrivateKey));
        return (version == SesVersion.V1
            ? new DerSequence(info, new DerSequence(cert, oid, signature))
            : new DerSequence(info, cert, oid, signature)).GetEncoded("DER");
    }
    internal static byte[] Hash(byte[] xml) => OfdDigestAlgorithms.Compute(OfdDigestAlgorithms.Sm3Oid, xml);
}
