using System;
using System.Collections.Generic;
using System.IO;
using Org.BouncyCastle.Asn1;

namespace Ofdrw.Net.Signatures.SesInterop;

// This is the narrow V5 full-certificate-list profile used by the external provider.
// Parsing the wider structure remains the responsibility of SesSignedValueReader.
internal sealed class SesV5Seal
{
    internal byte[] SealInfoDer { get; private set; } = Array.Empty<byte>();
    internal byte[] MakerCertificate { get; private set; } = Array.Empty<byte>();
    internal byte[] MakerSignature { get; private set; } = Array.Empty<byte>();
    internal List<byte[]> AuthorizedCertificates { get; } = new();
    internal DateTime ValidFrom { get; private set; }
    internal DateTime ValidUntil { get; private set; }

    internal static SesV5Seal Read(byte[] sealDer)
    {
        var seal = SesDer.Sequence(SesDer.Read(sealDer), 4);
        var info = SesDer.Sequence(seal[0], 4);
        var header = SesDer.Sequence(info[0], 3);
        if (SesDer.Ia5(header[0]) != "ES" || SesDer.Integer(header[1]) != 5)
            throw new InvalidDataException("Expected SES V5 seal header.");
        _ = SesDer.Ia5(header[2]);
        _ = SesDer.Ia5(info[1]);
        var properties = SesDer.Sequence(info[2], 7);
        _ = SesDer.Integer(properties[0]);
        if (properties[1].ToAsn1Object() is not DerUtf8String name || name.GetString().Length > 1024)
            throw new InvalidDataException("Invalid V5 seal name.");
        if (SesDer.Integer(properties[2]) != 1)
            throw new InvalidDataException("Only the full-certificate V5 authorization list is supported.");
        if (properties[3].ToAsn1Object() is not Asn1Sequence list || list.Count == 0 || list.Count > 64)
            throw new InvalidDataException("Invalid V5 authorization list.");
        var result = new SesV5Seal { SealInfoDer = info.GetEncoded("DER") };
        foreach (var item in list) result.AuthorizedCertificates.Add(SesDer.Octets(item, 64 * 1024));
        for (var i = 4; i <= 6; i++)
            if (properties[i].ToAsn1Object() is not Asn1GeneralizedTime)
                throw new InvalidDataException("Expected V5 seal GeneralizedTime.");
        result.ValidFrom = ((Asn1GeneralizedTime)properties[5].ToAsn1Object()).ToDateTime().ToUniversalTime();
        result.ValidUntil = ((Asn1GeneralizedTime)properties[6].ToAsn1Object()).ToDateTime().ToUniversalTime();
        if (result.ValidFrom > result.ValidUntil)
            throw new InvalidDataException("Invalid V5 seal validity interval.");
        var picture = SesDer.Sequence(info[3], 4);
        _ = SesDer.Ia5(picture[0]);
        _ = SesDer.Octets(picture[1], 512 * 1024);
        if (SesDer.Integer(picture[2]) <= 0 || SesDer.Integer(picture[3]) <= 0)
            throw new InvalidDataException("Invalid V5 seal picture dimensions.");
        result.MakerCertificate = SesDer.Octets(seal[1], 64 * 1024);
        if (SesDer.Oid(seal[2]) != SesTestCrypto.Algorithm)
            throw new InvalidDataException("Unsupported V5 seal signature algorithm.");
        result.MakerSignature = SesDer.Bits(seal[3], 4096);
        return result;
    }

    internal bool Authorizes(byte[] signerCertificate)
    {
        foreach (var certificate in AuthorizedCertificates)
            if (SesTestCrypto.Equal(certificate, signerCertificate)) return true;
        return false;
    }
}
