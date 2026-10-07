using System;
using System.IO;
using System.Text;
using Org.BouncyCastle.Asn1;

namespace Ofdrw.Net.Signatures.SesInterop;

/// <summary>Bounded, structural SES V1/V4 inspection. It does not establish trust or signature validity.</summary>
public static class SesSignedValueReader
{
    /// <summary>Maximum DER input size in bytes.</summary>
    public const int MaximumSignedValueBytes = 1024 * 1024;

    /// <summary>Parse common fields from one canonical DER container.</summary>
    /// <exception cref="ArgumentNullException">Input is null.</exception>
    /// <exception cref="InvalidDataException">Input is malformed, unsupported or over a documented limit.</exception>
    public static SesSignedValue Parse(byte[] signedValue)
    {
        if (signedValue is null) throw new ArgumentNullException(nameof(signedValue));
        try
        {
            var outer = SesDer.Sequence(SesDer.Read(signedValue), 2, 4, 5);
            var tbs = SesDer.Sequence(outer[0], 5, 6, 7);
            var version = (SesVersion)SesDer.Integer(tbs[0]);
            if (version != SesVersion.V1 && version != SesVersion.V4) throw new InvalidDataException("Unsupported SES version.");
            if (version == SesVersion.V1 && (outer.Count != 2 || tbs.Count != 7) ||
                version == SesVersion.V4 && (outer.Count < 4 || tbs.Count > 6))
                throw new InvalidDataException("SES version/layout mismatch.");
            var seal = ReadSeal(tbs[1], version);
            byte[] time; string timeText;
            if (version == SesVersion.V1)
            {
                time = SesDer.Bits(tbs[2], 128);
                timeText = new UTF8Encoding(false, true).GetString(time);
            }
            else
            {
                if (tbs[2].ToAsn1Object() is not Asn1GeneralizedTime generalized)
                    throw new InvalidDataException("Expected V4 GeneralizedTime.");
                time = generalized.GetEncoded("DER"); timeText = generalized.TimeString;
                _ = generalized.ToDateTime();
                if (tbs.Count == 6) Tagged(tbs[5]);
            }
            var hash = SesDer.Bits(tbs[3], 64);
            var property = SesDer.Ia5(tbs[4]);
            var cert = SesDer.Octets(version == SesVersion.V1 ? tbs[5] : outer[1], 64 * 1024);
            var algorithm = SesDer.Oid(version == SesVersion.V1 ? tbs[6] : outer[2]);
            var signature = SesDer.Bits(version == SesVersion.V1 ? outer[1] : outer[3], 4096);
            byte[]? timestamp = null;
            if (version == SesVersion.V4 && outer.Count == 5)
                timestamp = SesDer.Bits(Tagged(outer[4]), 64 * 1024);
            return new SesSignedValue(version, seal, time, timeText, hash, property, cert, algorithm,
                signature, tbs.GetEncoded("DER"), tbs[1].GetEncoded("DER"), timestamp);
        }
        catch (Exception exception) when (exception is ArgumentException || exception is ArithmeticException ||
            exception is FormatException || exception is DecoderFallbackException)
        {
            throw new InvalidDataException("Malformed SES fields.", exception);
        }
    }

    private static Asn1Object Tagged(Asn1Encodable value)
    {
        if (value.ToAsn1Object() is not Asn1TaggedObject tagged || tagged.TagNo != 0 || !tagged.IsExplicit())
            throw new InvalidDataException("Expected SES [0] EXPLICIT field.");
        return tagged.GetExplicitBaseObject().ToAsn1Object();
    }

    private static SesSealInfo ReadSeal(Asn1Encodable value, SesVersion version)
    {
        var seal = SesDer.Sequence(value, version == SesVersion.V1 ? 2 : 4);
        var info = SesDer.Sequence(seal[0], 4, 5);
        var header = SesDer.Sequence(info[0], 3);
        if (SesDer.Ia5(header[0]) != "ES" || SesDer.Integer(header[1]) != (int)version)
            throw new InvalidDataException("SES seal header mismatch.");
        var properties = SesDer.Sequence(info[2], version == SesVersion.V1 ? 6 : 7);
        if (properties[1].ToAsn1Object() is not DerUtf8String name || name.GetString().Length > 1024)
            throw new InvalidDataException("Expected bounded SES UTF8 name.");
        var certIndex = version == SesVersion.V1 ? 2 : 3;
        var listType = version == SesVersion.V1 ? 1 : SesDer.Integer(properties[2]);
        if (listType != 1 && listType != 2) throw new InvalidDataException("Unsupported SES certificate list type.");
        if (properties[certIndex].ToAsn1Object() is not Asn1Sequence certList || certList.Count > 64)
            throw new InvalidDataException("Invalid SES certificate list.");
        foreach (var certificate in certList)
        {
            if (listType == 1) _ = SesDer.Octets(certificate, 64 * 1024);
            else
            {
                var digest = SesDer.Sequence(certificate, 2);
                if (digest[0].ToAsn1Object() is not DerPrintableString text || text.GetString().Length > 128)
                    throw new InvalidDataException("Expected SES certificate digest type.");
                _ = SesDer.Octets(digest[1], 64);
            }
        }
        for (var i = certIndex + 1; i < properties.Count; i++)
        {
            if (version == SesVersion.V1 && properties[i].ToAsn1Object() is Asn1UtcTime utc) _ = utc.ToDateTime();
            else if (version == SesVersion.V4 && properties[i].ToAsn1Object() is Asn1GeneralizedTime date) _ = date.ToDateTime();
            else throw new InvalidDataException("Expected version-specific seal time.");
        }
        var picture = SesDer.Sequence(info[3], 4);
        var width = SesDer.Integer(picture[2]); var height = SesDer.Integer(picture[3]);
        if (width <= 0 || height <= 0) throw new InvalidDataException("Invalid seal picture dimensions.");
        var signInfo = version == SesVersion.V1 ? SesDer.Sequence(seal[1], 3) : seal;
        var offset = version == SesVersion.V1 ? 0 : 1;
        _ = SesDer.Octets(signInfo[offset], 64 * 1024);
        _ = SesDer.Oid(signInfo[offset + 1]); _ = SesDer.Bits(signInfo[offset + 2], 4096);
        return new SesSealInfo(SesDer.Ia5(info[1]), SesDer.Ia5(header[2]), (int)version,
            SesDer.Integer(properties[0]), name.GetString(), SesDer.Ia5(picture[0]),
            SesDer.Octets(picture[1], 512 * 1024), width, height);
    }
}
