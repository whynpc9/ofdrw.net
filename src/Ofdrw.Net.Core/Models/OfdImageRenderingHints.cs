using System;
using System.IO;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// An exact, nonstructural export hint. No other foreign attribute is granted meaning.
internal static class OfdImageRenderingHints
{
    internal static readonly XName PdfInterpolateV1 = XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints");
    internal static bool IsCanonicalValue(string value) => value is "true" or "false";
    internal static bool? ReadPdfInterpolate(OfdImageElement image)
    {
        if (string.IsNullOrWhiteSpace(image.SourceXml)) return null;
        var root = XElement.Parse(image.SourceXml!, LoadOptions.PreserveWhitespace);
        // Only the root ImageObject owns the hint. Never search descendants or local names.
        if (root.Name.LocalName != "ImageObject") return null;
        var hint = root.Attribute(PdfInterpolateV1);
        if (hint is null) return null;
        if (!IsCanonicalValue(hint.Value)) throw new InvalidDataException("Invalid PdfInterpolateV1 hint; canonical true/false required.");
        return string.Equals(hint.Value, "true", StringComparison.Ordinal);
    }
}
