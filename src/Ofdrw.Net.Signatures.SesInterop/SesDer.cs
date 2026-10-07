using System;
using System.IO;
using System.Linq;
using Org.BouncyCastle.Asn1;

namespace Ofdrw.Net.Signatures.SesInterop;

internal static class SesDer
{
    internal const int MaximumBytes = 1024 * 1024;

    internal static Asn1Object Read(byte[] bytes, int maximumBytes = MaximumBytes)
    {
        if (bytes is null) throw new ArgumentNullException(nameof(bytes));
        if (bytes.Length == 0 || bytes.Length > maximumBytes)
            throw new InvalidDataException("DER input exceeds its size limit or is empty.");
        try
        {
            int offset = 0, nodes = 0;
            Scan(bytes, ref offset, bytes.Length, 0, ref nodes);
            if (offset != bytes.Length) throw new InvalidDataException("Trailing DER data.");
            var value = Asn1Object.FromByteArray(bytes);
            if (!bytes.SequenceEqual(value.GetEncoded("DER"))) throw new InvalidDataException("Non-canonical DER.");
            return value;
        }
        catch (Exception exception) when (exception is ArgumentException || exception is IOException)
        {
            throw new InvalidDataException("Malformed or unsupported DER.", exception);
        }
    }

    // Bound recursion before the ASN.1 library sees constructed input (including certificates).
    private static void Scan(byte[] bytes, ref int offset, int end, int depth, ref int nodes)
    {
        if (depth >= 32 || ++nodes > 4096 || offset >= end) throw new InvalidDataException("DER structural limit.");
        var tag = bytes[offset++];
        if ((tag & 31) == 31 || tag == 0 || offset >= end) throw new InvalidDataException("Unsupported DER tag.");
        int length = bytes[offset++];
        if (length >= 128)
        {
            int count = length & 127;
            if (count == 0 || count > 3 || count > end - offset || bytes[offset] == 0)
                throw new InvalidDataException("Invalid DER length.");
            length = 0;
            for (int index = 0; index < count; index++) length = (length << 8) | bytes[offset++];
            if (length < 128) throw new InvalidDataException("Non-minimal DER length.");
        }
        if (length > end - offset) throw new InvalidDataException("Truncated DER.");
        int next = offset + length;
        if ((tag & 32) != 0)
        {
            while (offset < next) Scan(bytes, ref offset, next, depth + 1, ref nodes);
        }
        else offset = next;
    }

    internal static Asn1Sequence Sequence(Asn1Encodable value, params int[] counts)
    {
        if (value.ToAsn1Object() is not Asn1Sequence sequence || !counts.Contains(sequence.Count))
            throw new InvalidDataException("Unexpected SES sequence shape.");
        return sequence;
    }
    internal static int Integer(Asn1Encodable value)
    {
        if (value.ToAsn1Object() is not DerInteger integer) throw new InvalidDataException("Expected SES integer.");
        return integer.IntValueExact;
    }
    internal static string Ia5(Asn1Encodable value)
    {
        if (value.ToAsn1Object() is not DerIA5String text) throw new InvalidDataException("Expected SES IA5 string.");
        var result = text.GetString();
        if (result.Length > 1024 || result.Any(c => c > 127 || c < 32)) throw new InvalidDataException("Invalid SES IA5 text.");
        return result;
    }
    internal static byte[] Octets(Asn1Encodable value, int limit)
    {
        if (value.ToAsn1Object() is not Asn1OctetString octets) throw new InvalidDataException("Expected SES octets.");
        var result = octets.GetOctets();
        if (result.Length > limit) throw new InvalidDataException("SES octet limit.");
        return result;
    }
    internal static byte[] Bits(Asn1Encodable value, int limit)
    {
        if (value.ToAsn1Object() is not DerBitString bits || bits.PadBits != 0) throw new InvalidDataException("Expected octet-aligned SES bits.");
        var result = bits.GetOctets();
        if (result.Length > limit) throw new InvalidDataException("SES bit-string limit.");
        return result;
    }
    internal static string Oid(Asn1Encodable value)
    {
        if (value.ToAsn1Object() is not DerObjectIdentifier oid) throw new InvalidDataException("Expected SES OID.");
        return oid.Id;
    }
}
