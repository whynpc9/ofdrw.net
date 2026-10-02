using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using System.Threading;
using System.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class OfdSignatureAppearanceReader
{
    public static IReadOnlyList<OfdSignatureAppearance> Read(
        OfdDocumentPackage package, HashSet<string>? selectedPageIds = null,
        int? maximumAppearances = null, CancellationToken cancellationToken = default)
    {
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var ofdBytes))
        {
            return Array.Empty<OfdSignatureAppearance>();
        }

        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            var ofd = ParseXml(ofdBytes);
            var signatureLists = ofd
                .Descendants()
                .Where(element => element.Name.LocalName == "Signatures")
                .Select(element => NormalizePath(element.Value))
                .Where(path => path.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var result = new List<OfdSignatureAppearance>();
            var payloadCache = maximumAppearances.HasValue ? new Dictionary<byte[], byte[]>() : null;
            var candidates = 0;
            foreach (var listPath in signatureLists)
            {
                ReadSignatureList(package.PreservedEntries, listPath, result, selectedPageIds,
                    maximumAppearances, ref candidates, payloadCache, cancellationToken);
            }

            return result;
        }
        catch (OperationCanceledException) { throw; }
        catch (InvalidDataException) when (maximumAppearances.HasValue) { throw; }
        catch (OutOfMemoryException) { throw; }
        catch (Exception exception) when (maximumAppearances.HasValue)
        {
            throw new InvalidDataException("Cannot parse selected signature appearance metadata.", exception);
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
        ICollection<OfdSignatureAppearance> destination, HashSet<string>? selectedPageIds,
        int? maximumAppearances, ref int candidates, Dictionary<byte[], byte[]>? payloadCache, CancellationToken token)
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
            token.ThrowIfCancellationRequested();
            ReadSignature(entries, signaturePath, destination, selectedPageIds, maximumAppearances, ref candidates, payloadCache, token);
        }
    }

    private static void ReadSignature(
        IReadOnlyDictionary<string, byte[]> entries,
        string signaturePath,
        ICollection<OfdSignatureAppearance> destination, HashSet<string>? selectedPageIds,
        int? maximumAppearances, ref int candidates, Dictionary<byte[], byte[]>? payloadCache, CancellationToken token)
    {
        if (!entries.TryGetValue(signaturePath, out var signatureBytes))
        {
            return;
        }

        var signature = ParseXml(signatureBytes);
        var stamps = new List<XElement>();
        foreach (var stamp in signature.Descendants().Where(element => element.Name.LocalName == "StampAnnot" &&
            (selectedPageIds is null || selectedPageIds.Contains(element.Attribute("PageRef")?.Value ?? string.Empty))))
        {
            token.ThrowIfCancellationRequested();
            if (maximumAppearances.HasValue && candidates >= maximumAppearances.Value)
                throw new InvalidDataException("Selected signature appearance count exceeds the configured limit.");
            candidates++;
            stamps.Add(stamp);
        }
        if (stamps.Count == 0) return;
        var appearanceData = ReadAppearanceData(
            entries,
            signature,
            signaturePath, payloadCache, token);
        if (appearanceData.Length == 0)
        {
            if (maximumAppearances.HasValue)
                throw new InvalidDataException("Selected signature stamp has no supported appearance payload.");
            return;
        }

        foreach (var stamp in stamps)
        {
            var pageId = stamp.Attribute("PageRef")?.Value;
            if (string.IsNullOrWhiteSpace(pageId) ||
                !TryParseBox(stamp.Attribute("Boundary")?.Value, out var box) ||
                !PdfOperandGeometry.Box(box.X, box.Y, box.Width, box.Height))
            {
                if (maximumAppearances.HasValue)
                    throw new InvalidDataException("Selected signature stamp Boundary must have finite PDF coordinates and positive dimensions representable by the PDF writer.");
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
        string signaturePath, Dictionary<byte[], byte[]>? payloadCache, CancellationToken token)
    {
        var sealLocation = signature
            .Descendants()
            .FirstOrDefault(element => element.Name.LocalName == "Seal")
            ?.Attribute("BaseLoc")?.Value;
        if (!string.IsNullOrWhiteSpace(sealLocation))
        {
            var sealPath = ResolvePath(signaturePath, sealLocation!);
            if (entries.TryGetValue(sealPath, out var sealBytes) &&
                TryFindCachedPayload(sealBytes, payloadCache, token, out var sealAppearance))
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
            TryFindCachedPayload(signedValue, payloadCache, token, out var appearance)
                ? appearance
                : Array.Empty<byte>();
    }

    private static bool TryFindCachedPayload(byte[] data, Dictionary<byte[], byte[]>? cache,
        CancellationToken token, out byte[] appearance)
    {
        if (cache is not null && cache.TryGetValue(data, out appearance!)) return appearance.Length > 0;
        token.ThrowIfCancellationRequested();
        if (IsSupportedAppearance(data)) appearance = cache is null ? (byte[])data.Clone() : data;
        else
        {
            ArraySegment<byte> best = default;
            FindLargestAppearance(data, 0, data.Length, 0, ref best, token, strict: cache is not null);
            appearance = best.Count == 0 ? Array.Empty<byte>() : new byte[best.Count];
            if (best.Count > 0) Buffer.BlockCopy(data, best.Offset, appearance, 0, best.Count);
        }
        if (cache is not null) cache.Add(data, appearance); // Cache misses too; the key is the archive's canonical entry byte array.
        return appearance.Length > 0;
    }

    private static void FindLargestAppearance(byte[] data, int offset, int length, int depth,
        ref ArraySegment<byte> best, CancellationToken token, bool strict)
    {
        if (depth > 32) return;
        ValidateSlice(data, offset, length);
        var end = offset + length;
        while (offset < end)
        {
            token.ThrowIfCancellationRequested();
            if (!TryReadTagAndLength(data, offset, end, out var tagClass, out var tagNumber,
                out var constructed, out var contentOffset, out var contentLength, out var nextOffset))
            {
                if (strict) throw new InvalidDataException("Malformed ASN.1 signature appearance length or tag.");
                return; // Tolerant PDF: retain candidates already found and appearances from other signatures.
            }
            // Retain one candidate, rather than allocating/sorting a list proportional to every ASN.1 octet.
            if (tagClass == 0 && tagNumber == 4 && !constructed && contentLength > best.Count &&
                IsSupportedAppearance(data, contentOffset, contentLength))
                best = new ArraySegment<byte>(data, contentOffset, contentLength);
            if (constructed) FindLargestAppearance(data, contentOffset, contentLength, depth + 1, ref best, token, strict);
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
            if (lengthOctets == 0 || lengthOctets > 4 || lengthOctets > end - offset)
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
        if (contentLength < 0 || contentOffset > end || contentLength > end - contentOffset)
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

    private static void ValidateSlice(byte[] data, int offset, int length)
    {
        if (offset < 0 || length < 0 || offset > data.Length || length > data.Length - offset)
            throw new InvalidDataException("Signature appearance slice exceeds its encoded payload.");
    }

    private static bool IsSupportedAppearance(
        byte[] data,
        int offset,
        int length)
    {
        ValidateSlice(data, offset, length);
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
        if (values.Length != 4 || values.Any(value => double.IsNaN(value) || double.IsInfinity(value)))
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
