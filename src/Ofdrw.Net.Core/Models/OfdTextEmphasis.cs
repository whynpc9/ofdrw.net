using System;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdTextEmphasis
{
    internal static readonly XName FauxItalicFactor = XName.Get("FauxItalicMatrixV1", "https://ofdrw.net/style-hints");

    internal static (double[]? Matrix, double[]? Factor) DrawingTransform(OfdTextElement text)
    {
        if (text.Transform is not { Length: 6 } matrix || string.IsNullOrWhiteSpace(text.SourceXml)) return (text.Transform, null);
        var value = XElement.Parse(text.SourceXml!).Attribute(FauxItalicFactor)?.Value;
        if (string.IsNullOrWhiteSpace(value)) return (matrix, null);
        var values = value!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        var factor = values.Select(token => double.TryParse(token, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ? number : double.NaN).ToArray();
        if (factor.Length != 6 || factor.Any(number => double.IsNaN(number) || double.IsInfinity(number)) ||
            factor[0] != 1 || factor[1] != 0 || factor[2] != -0.2 || factor[3] != 1 || factor[4] < 0 || factor[5] != 0)
            return (matrix, null);
        var inverse = new[] { 1d, 0, 0.2, 1, -factor[4], 0 };
        return (new[] { matrix[0] * inverse[0] + matrix[2] * inverse[1], matrix[1] * inverse[0] + matrix[3] * inverse[1],
            matrix[0] * inverse[2] + matrix[2] * inverse[3], matrix[1] * inverse[2] + matrix[3] * inverse[3],
            matrix[0] * inverse[4] + matrix[2] * inverse[5] + matrix[4], matrix[1] * inverse[4] + matrix[3] * inverse[5] + matrix[5] }, factor);
    }

    internal static (double X, double Y) Anchor(double x, double y, double[]? factor) => factor is null ? (x, y) :
        (factor[0] * x + factor[2] * y + factor[4], factor[1] * x + factor[3] * y + factor[5]);
}
