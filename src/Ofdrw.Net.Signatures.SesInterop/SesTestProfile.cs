using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;
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
        if (xml is null) throw new ArgumentNullException(nameof(xml));
        if (property is null) throw new ArgumentNullException(nameof(property));
        if (xml.Length == 0 || xml.Length > 1024 * 1024) throw new InvalidDataException("Signature XML size limit.");
        if (property.Length == 0 || property.Length > 1024 || property[0] != '/' ||
            property.Any(c => c < 33 || c > 126 || c == '\\') ||
            property.Split('/').Skip(1).Any(s => s.Length == 0 || s == "." || s == "..") ||
            !property.EndsWith("/Signature.xml", StringComparison.Ordinal))
            throw new InvalidDataException("Expected a bounded absolute Signature.xml package path.");
        using var stream = new MemoryStream(xml, writable: false);
        using var reader = XmlReader.Create(stream, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = 1024 * 1024
        });
        var document = XDocument.Load(reader);
        XNamespace ns = "http://www.ofdspec.org/2016";
        if (document.Root?.Name != ns + "Signature") throw new InvalidDataException("Unexpected signature namespace/root.");
        XElement Unique(string name, XElement parent)
        {
            var all = document.Descendants().Where(e => e.Name.LocalName == name).ToArray();
            if (all.Length != 1 || all[0].Name != ns + name || all[0].Parent != parent || all[0].HasElements)
                throw new InvalidDataException("Ambiguous signature field: " + name);
            return all[0];
        }
        var infos = document.Descendants().Where(e => e.Name.LocalName == "SignedInfo").ToArray();
        if (infos.Length != 1 || infos[0].Name != ns + "SignedInfo" || infos[0].Parent != document.Root)
            throw new InvalidDataException("Ambiguous SignedInfo.");
        if (Unique("SignatureMethod", infos[0]).Value != SesTestCrypto.Algorithm)
            throw new InvalidDataException("Expected SM2-with-SM3 method.");
        var text = Unique("SignatureDateTime", infos[0]).Value;
        if (!DateTime.TryParseExact(text, "yyyy-MM-dd'T'HH:mm:ss'Z'", CultureInfo.InvariantCulture,
            DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal, out var time) || time < Begin || time >= End)
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
