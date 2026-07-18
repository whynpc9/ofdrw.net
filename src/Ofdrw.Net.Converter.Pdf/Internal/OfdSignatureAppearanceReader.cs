using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class OfdSignatureAppearanceReader
{
    public static IReadOnlyList<OfdSignatureAppearance> Read(
        OfdDocumentPackage package)
    {
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var ofdBytes))
        {
            return Array.Empty<OfdSignatureAppearance>();
        }

        try
        {
            var ofd = ParseXml(ofdBytes);
            var signatureLists = ofd
                .Descendants()
                .Where(element => element.Name.LocalName == "Signatures")
                .Select(element => NormalizePath(element.Value))
                .Where(path => path.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var result = new List<OfdSignatureAppearance>();
            foreach (var listPath in signatureLists)
            {
                ReadSignatureList(package.PreservedEntries, listPath, result);
            }

            return result;
        }
        catch
        {
            // A malformed or unsupported signature must not prevent the document
            // body from being converted.
            return Array.Empty<OfdSignatureAppearance>();
        }
    }

    private static void ReadSignatureList(
        IReadOnlyDictionary<string, byte[]> entries,
        string listPath,
        ICollection<OfdSignatureAppearance> destination)
    {
        if (!entries.TryGetValue(listPath, out var listBytes))
        {
            return;
        }

        var list = ParseXml(listBytes);
        foreach (var record in list
            .Descendants()
            .Where(element => element.Name.LocalName == "Signature"))
        {
            var baseLocation = record.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(baseLocation))
            {
                continue;
            }

            var signaturePath = ResolvePath(listPath, baseLocation!);
            ReadSignature(entries, signaturePath, destination);
        }
    }

    private static void ReadSignature(
        IReadOnlyDictionary<string, byte[]> entries,
        string signaturePath,
        ICollection<OfdSignatureAppearance> destination)
    {
        if (!entries.TryGetValue(signaturePath, out var signatureBytes))
        {
            return;
        }

        var signature = ParseXml(signatureBytes);
        var appearanceData = ReadAppearanceData(
            entries,
            signature,
            signaturePath);
        if (appearanceData.Length == 0)
        {
            return;
        }

        foreach (var stamp in signature
            .Descendants()
            .Where(element => element.Name.LocalName == "StampAnnot"))
        {
            var pageId = stamp.Attribute("PageRef")?.Value;
            if (string.IsNullOrWhiteSpace(pageId) ||
                !TryParseBox(stamp.Attribute("Boundary")?.Value, out var box) ||
                box.Width <= 0 ||
                box.Height <= 0)
            {
                continue;
            }

            destination.Add(new OfdSignatureAppearance(
                pageId!,
                box.X,
                box.Y,
                box.Width,
                box.Height,
                appearanceData));
        }
    }

    private static byte[] ReadAppearanceData(
        IReadOnlyDictionary<string, byte[]> entries,
        XDocument signature,
        string signaturePath)
    {
        var sealLocation = signature
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "Seal")
            ?.Attribute("BaseLoc")?.Value;
        if (!string.IsNullOrWhiteSpace(sealLocation))
        {
            var sealPath = ResolvePath(signaturePath, sealLocation!);
            if (entries.TryGetValue(sealPath, out var sealBytes) &&
                TryFindAppearancePayload(sealBytes, out var sealAppearance))
            {
                return sealAppearance;
            }
        }

        var signedValueLocation = signature
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "SignedValue")
            ?.Value;
        if (string.IsNullOrWhiteSpace(signedValueLocation))
        {
            return Array.Empty<byte>();
        }

        var signedValuePath = ResolvePath(signaturePath, signedValueLocation!);
        return entries.TryGetValue(signedValuePath, out var signedValue) &&
            TryFindAppearancePayload(signedValue, out var appearance)
                ? appearance
                : Array.Empty<byte>();
    }

    private static bool TryFindAppearancePayload(
        byte[] data,
        out byte[] appearance)
    {
        if (IsSupportedAppearance(data))
        {
            appearance = (byte[])data.Clone();
            return true;
        }

        var candidates = new List<ArraySegment<byte>>();
        CollectOctetStrings(data, 0, data.Length, 0, candidates);
        foreach (var candidate in candidates.OrderByDescending(value => value.Count))
        {
            if (!IsSupportedAppearance(
                    candidate.Array!,
                    candidate.Offset,
                    candidate.Count))
            {
                continue;
            }

            appearance = new byte[candidate.Count];
            Buffer.BlockCopy(
                candidate.Array!,
                candidate.Offset,
                appearance,
                0,
                candidate.Count);
            return true;
        }

        appearance = Array.Empty<byte>();
        return false;
    }

    private static void CollectOctetStrings(
        byte[] data,
        int offset,
        int length,
        int depth,
        ICollection<ArraySegment<byte>> destination)
    {
        if (depth > 32 || offset < 0 || length < 0 || offset + length > data.Length)
        {
            return;
        }

        var end = offset + length;
        while (offset < end)
        {
            if (!TryReadTagAndLength(
                    data,
                    offset,
                    end,
                    out var tagClass,
                    out var tagNumber,
                    out var constructed,
                    out var contentOffset,
                    out var contentLength,
                    out var nextOffset))
            {
                return;
            }

            if (tagClass == 0 && tagNumber == 4 && !constructed)
            {
                destination.Add(
                    new ArraySegment<byte>(
                        data,
                        contentOffset,
                        contentLength));
            }

            if (constructed)
            {
                CollectOctetStrings(
                    data,
                    contentOffset,
                    contentLength,
                    depth + 1,
                    destination);
            }

            offset = nextOffset;
        }
    }

    private static bool TryReadTagAndLength(
        byte[] data,
        int offset,
        int end,
        out int tagClass,
        out int tagNumber,
        out bool constructed,
        out int contentOffset,
        out int contentLength,
        out int nextOffset)
    {
        tagClass = 0;
        tagNumber = 0;
        constructed = false;
        contentOffset = 0;
        contentLength = 0;
        nextOffset = 0;
        if (offset >= end)
        {
            return false;
        }

        var first = data[offset++];
        tagClass = first >> 6;
        constructed = (first & 0x20) != 0;
        tagNumber = first & 0x1f;
        if (tagNumber == 0x1f)
        {
            tagNumber = 0;
            var tagOctets = 0;
            do
            {
                if (offset >= end || tagOctets++ >= 5)
                {
                    return false;
                }

                var value = data[offset++];
                if (tagNumber > (int.MaxValue >> 7))
                {
                    return false;
                }

                tagNumber = (tagNumber << 7) | (value & 0x7f);
                if ((value & 0x80) == 0)
                {
                    break;
                }
            }
            while (true);
        }

        if (offset >= end)
        {
            return false;
        }

        var firstLength = data[offset++];
        if ((firstLength & 0x80) == 0)
        {
            contentLength = firstLength;
        }
        else
        {
            var lengthOctets = firstLength & 0x7f;
            if (lengthOctets == 0 || lengthOctets > 4 || offset + lengthOctets > end)
            {
                return false;
            }

            contentLength = 0;
            for (var index = 0; index < lengthOctets; index++)
            {
                if (contentLength > (int.MaxValue >> 8))
                {
                    return false;
                }

                contentLength = (contentLength << 8) | data[offset++];
            }
        }

        contentOffset = offset;
        if (contentLength < 0 || contentOffset + contentLength > end)
        {
            return false;
        }

        nextOffset = contentOffset + contentLength;
        return true;
    }

    private static bool IsSupportedAppearance(byte[] data)
    {
        return IsSupportedAppearance(data, 0, data.Length);
    }

    private static bool IsSupportedAppearance(
        byte[] data,
        int offset,
        int length)
    {
        return IsZip(data, offset, length) ||
            IsPng(data, offset, length) ||
            IsJpeg(data, offset, length) ||
            IsGif(data, offset, length) ||
            IsBmp(data, offset, length) ||
            IsTiff(data, offset, length);
    }

    private static bool IsZip(byte[] data, int offset, int length)
    {
        return length >= 4 &&
            data[offset] == 0x50 &&
            data[offset + 1] == 0x4b &&
            data[offset + 2] == 0x03 &&
            data[offset + 3] == 0x04;
    }

    private static bool IsPng(byte[] data, int offset, int length)
    {
        return length >= 8 &&
            data[offset] == 0x89 &&
            data[offset + 1] == 0x50 &&
            data[offset + 2] == 0x4e &&
            data[offset + 3] == 0x47 &&
            data[offset + 4] == 0x0d &&
            data[offset + 5] == 0x0a &&
            data[offset + 6] == 0x1a &&
            data[offset + 7] == 0x0a;
    }

    private static bool IsJpeg(byte[] data, int offset, int length)
    {
        return length >= 3 &&
            data[offset] == 0xff &&
            data[offset + 1] == 0xd8 &&
            data[offset + 2] == 0xff;
    }

    private static bool IsGif(byte[] data, int offset, int length)
    {
        return length >= 6 &&
            data[offset] == (byte)'G' &&
            data[offset + 1] == (byte)'I' &&
            data[offset + 2] == (byte)'F' &&
            data[offset + 3] == (byte)'8' &&
            (data[offset + 4] == (byte)'7' ||
             data[offset + 4] == (byte)'9') &&
            data[offset + 5] == (byte)'a';
    }

    private static bool IsBmp(byte[] data, int offset, int length)
    {
        return length >= 2 &&
            data[offset] == (byte)'B' &&
            data[offset + 1] == (byte)'M';
    }

    private static bool IsTiff(byte[] data, int offset, int length)
    {
        return length >= 4 &&
            ((data[offset] == (byte)'I' &&
              data[offset + 1] == (byte)'I' &&
              data[offset + 2] == 0x2a &&
              data[offset + 3] == 0x00) ||
             (data[offset] == (byte)'M' &&
              data[offset + 1] == (byte)'M' &&
              data[offset + 2] == 0x00 &&
              data[offset + 3] == 0x2a));
    }

    private static XDocument ParseXml(byte[] data)
    {
        return XDocument.Parse(
            Encoding.UTF8.GetString(data).TrimStart('\uFEFF'),
            LoadOptions.PreserveWhitespace);
    }

    private static string ResolvePath(string containingPath, string reference)
    {
        if (reference.StartsWith("/", StringComparison.Ordinal))
        {
            return NormalizePath(reference);
        }

        var slash = containingPath.LastIndexOf('/');
        var directory = slash < 0
            ? string.Empty
            : containingPath.Substring(0, slash + 1);
        return NormalizePath(directory + reference);
    }

    private static string NormalizePath(string path)
    {
        var segments = new Stack<string>();
        foreach (var segment in path
            .Replace('\\', '/')
            .TrimStart('/')
            .Split('/'))
        {
            if (string.IsNullOrWhiteSpace(segment) || segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.Pop();
                }

                continue;
            }

            segments.Push(segment);
        }

        return string.Join("/", segments.Reverse());
    }

    private static bool TryParseBox(string? value, out OfdBox box)
    {
        box = default;
        if (string.IsNullOrWhiteSpace(value))
        {
            return false;
        }

        var values = value!
            .Split(
                new[] { ' ', '\t', '\r', '\n' },
                StringSplitOptions.RemoveEmptyEntries)
            .Select(part => double.TryParse(
                part,
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out var parsed)
                    ? parsed
                    : double.NaN)
            .ToArray();
        if (values.Length != 4 || values.Any(double.IsNaN))
        {
            return false;
        }

        box = new OfdBox(values[0], values[1], values[2], values[3]);
        return true;
    }

    private readonly struct OfdBox
    {
        public OfdBox(double x, double y, double width, double height)
        {
            X = x;
            Y = y;
            Width = width;
            Height = height;
        }

        public double X { get; }

        public double Y { get; }

        public double Width { get; }

        public double Height { get; }
    }
}

internal sealed class OfdSignatureAppearance
{
    public OfdSignatureAppearance(
        string pageId,
        double xMillimeters,
        double yMillimeters,
        double widthMillimeters,
        double heightMillimeters,
        byte[] data)
    {
        PageId = pageId;
        XMillimeters = xMillimeters;
        YMillimeters = yMillimeters;
        WidthMillimeters = widthMillimeters;
        HeightMillimeters = heightMillimeters;
        Data = data;
    }

    public string PageId { get; }

    public double XMillimeters { get; }

    public double YMillimeters { get; }

    public double WidthMillimeters { get; }

    public double HeightMillimeters { get; }

    public byte[] Data { get; }
}
