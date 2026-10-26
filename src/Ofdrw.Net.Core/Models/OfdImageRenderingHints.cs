using System;
using System.Globalization;
using System.IO;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// An exact, nonstructural export hint. No other foreign attribute is granted meaning.
internal static class OfdImageRenderingHints
{
    internal static readonly XName PdfInterpolateV1 = XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints");
    internal static readonly XName PdfSourceCropV1 = XName.Get("PdfSourceCropV1", "https://ofdrw.net/image-hints");
    internal static bool IsCanonicalValue(string value) => value is "true" or "false";
    internal static string FormatSourceCrop(double x, double y, double width, double height) =>
        string.Join(" ", OfdNumericFormat.Plain(x), OfdNumericFormat.Plain(y),
            OfdNumericFormat.Plain(width), OfdNumericFormat.Plain(height));
    internal static bool IsCanonicalSourceCrop(string value) => ParseSourceCrop(value) is not null;
    internal static (double X, double Y, double Width, double Height)? ReadPdfSourceCrop(OfdImageElement image)
    {
        if (string.IsNullOrWhiteSpace(image.SourceXml)) return null;
        var root = XElement.Parse(image.SourceXml!, LoadOptions.PreserveWhitespace);
        if (root.Name.LocalName != "ImageObject") return null;
        var hint = root.Attribute(PdfSourceCropV1);
        if (hint is null) return null;
        return ParseSourceCrop(hint.Value) ?? throw new InvalidDataException("Invalid PdfSourceCropV1 hint.");
    }
    private static (double X, double Y, double Width, double Height)? ParseSourceCrop(string value)
    {
        var parts = value.Split(' ');
        if (parts.Length != 4) return null;
        var numbers = new double[4];
        for (var i = 0; i < 4; i++)
        {
            if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out numbers[i]) ||
                double.IsNaN(numbers[i]) || double.IsInfinity(numbers[i]) ||
                parts[i] != OfdNumericFormat.Plain(numbers[i])) return null;
        }
        if (numbers[0] < 0 || numbers[1] < 0 || numbers[2] <= 0 || numbers[3] <= 0) return null;
        return (numbers[0], numbers[1], numbers[2], numbers[3]);
    }
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
