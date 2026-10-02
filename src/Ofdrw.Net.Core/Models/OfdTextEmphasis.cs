using System;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdTextEmphasis
{
    internal static readonly XName FauxItalicFactor = XName.Get("FauxItalicMatrixV1", "https://ofdrw.net/style-hints");

    internal static (double[] Matrix, double[] Factor) ComposeNameOnlyItalic(double[] matrix, double writtenSize)
    {
        if (matrix.Length != 6 || matrix.Any(value => double.IsNaN(value) || double.IsInfinity(value)) ||
            double.IsNaN(writtenSize) || double.IsInfinity(writtenSize))
            throw new InvalidDataException("Generated italic CTM requires finite geometry.");
        const double shear = 0.2;
        var offset = shear * writtenSize;
        var combined = new[] { matrix[0], matrix[1], matrix[2] - shear * matrix[0], matrix[3] - shear * matrix[1],
            matrix[4] + matrix[0] * offset, matrix[5] + matrix[1] * offset };
        if (combined.Any(value => double.IsNaN(value) || double.IsInfinity(value)))
            throw new InvalidDataException("Generated italic CTM overflowed.");
        if (!OfdNumericFormat.Nonsingular(combined[0], combined[1], combined[2], combined[3]))
            throw new InvalidDataException("Generated italic CTM became singular.");
        return (combined, new[] { 1d, 0, -shear, 1, offset, 0 });
    }

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
