using System;
using System.IO;
using Org.BouncyCastle.Asn1;
using Org.BouncyCastle.Asn1.X509;

namespace Ofdrw.Net.Signatures.SesInterop;

// Narrow single-signer, attached-content profile matching the independent Java container.
internal sealed class Gbt35275SignedData
{
    internal const string DataOid = "1.2.156.10197.6.1.4.2.1";
    internal const string SignedDataOid = "1.2.156.10197.6.1.4.2.2";
    internal const string Sm3Oid = "1.2.156.10197.1.401";
    internal const string Sm2SignOid = "1.2.156.10197.1.301.1";

    internal byte[] CertificateDer { get; private set; } = Array.Empty<byte>();
    internal byte[] Content { get; private set; } = Array.Empty<byte>();
    internal byte[] Signature { get; private set; } = Array.Empty<byte>();

    internal static byte[] Encode(byte[] content, byte[] signature, byte[] certificateDer)
    {
        var certificate = X509CertificateStructure.GetInstance(SesDer.Read(certificateDer, 64 * 1024));
        var digestAlgorithm = new DerSequence(new DerObjectIdentifier(Sm3Oid));
        var signerInfo = new DerSequence(new DerInteger(1),
            new DerSequence(certificate.Issuer, certificate.SerialNumber),
            digestAlgorithm,
            new DerSequence(new DerObjectIdentifier(Sm2SignOid)),
            new DerOctetString(signature));
        var data = new DerSequence(new DerObjectIdentifier(DataOid),
            new DerTaggedObject(true, 0, new DerOctetString(content)));
        var signedData = new DerSequence(new DerInteger(1), new DerSet((Asn1Encodable)digestAlgorithm), data,
            new DerTaggedObject(false, 0, new DerSet((Asn1Encodable)certificate)),
            new DerSet((Asn1Encodable)signerInfo));
        return new DerSequence(new DerObjectIdentifier(SignedDataOid),
            new DerTaggedObject(true, 0, signedData)).GetEncoded("DER");
    }

    internal static Gbt35275SignedData Read(byte[] bytes)
    {
        var outer = SesDer.Sequence(SesDer.Read(bytes), 2);
        if (SesDer.Oid(outer[0]) != SignedDataOid) throw new InvalidDataException("Expected GB/T 35275 signedData.");
        var signedData = SesDer.Sequence(Explicit(outer[1]), 5);
        if (SesDer.Integer(signedData[0]) != 1 ||
            signedData[1].ToAsn1Object() is not Asn1Set digestAlgorithms || digestAlgorithms.Count != 1 ||
            Algorithm(digestAlgorithms[0]) != Sm3Oid)
            throw new InvalidDataException("Unsupported GB/T 35275 digest profile.");
        var contentInfo = SesDer.Sequence(signedData[2], 2);
        if (SesDer.Oid(contentInfo[0]) != DataOid)
            throw new InvalidDataException("Expected GB/T 35275 attached data.");
        var content = SesDer.Octets(Explicit(contentInfo[1]), 128);
        if (signedData[3].ToAsn1Object() is not Asn1TaggedObject tagged || tagged.TagNo != 0)
            throw new InvalidDataException("Expected implicit GB/T 35275 certificates.");
        var certificates = Asn1Set.GetInstance(tagged, false);
        if (!SesTestCrypto.Equal(tagged.GetEncoded("DER"),
            new DerTaggedObject(false, 0, certificates).GetEncoded("DER")))
            throw new InvalidDataException("Non-canonical implicit certificate set.");
        if (certificates.Count != 1) throw new InvalidDataException("Expected one signer certificate.");
        var certificateDer = certificates[0].ToAsn1Object().GetEncoded("DER");
        if (certificateDer.Length > 64 * 1024) throw new InvalidDataException("Certificate exceeds limit.");
        var certificate = X509CertificateStructure.GetInstance(certificates[0]);
        if (signedData[4].ToAsn1Object() is not Asn1Set signers || signers.Count != 1)
            throw new InvalidDataException("Expected one GB/T 35275 signer.");
        var signer = SesDer.Sequence(signers[0], 5);
        if (SesDer.Integer(signer[0]) != 1 || Algorithm(signer[2]) != Sm3Oid ||
            Algorithm(signer[3]) != Sm2SignOid)
            throw new InvalidDataException("Unsupported GB/T 35275 signer profile.");
        var issuerSerial = SesDer.Sequence(signer[1], 2);
        if (!issuerSerial[0].ToAsn1Object().Equals(certificate.Issuer.ToAsn1Object()) ||
            !issuerSerial[1].ToAsn1Object().Equals(certificate.SerialNumber.ToAsn1Object()))
            throw new InvalidDataException("Signer issuer and serial do not match the certificate.");
        return new Gbt35275SignedData
        {
            CertificateDer = certificateDer,
            Content = content,
            Signature = SesDer.Octets(signer[4], 80)
        };
    }

    private static Asn1Object Explicit(Asn1Encodable value)
    {
        if (value.ToAsn1Object() is not Asn1TaggedObject tagged || tagged.TagNo != 0 || !tagged.IsExplicit())
            throw new InvalidDataException("Expected [0] EXPLICIT content.");
        return tagged.GetExplicitBaseObject().ToAsn1Object();
    }

    private static string Algorithm(Asn1Encodable value)
    {
        var sequence = SesDer.Sequence(value, 1);
        return SesDer.Oid(sequence[0]);
    }
}
