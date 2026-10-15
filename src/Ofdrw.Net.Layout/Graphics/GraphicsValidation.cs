using System;
using System.Globalization;

namespace Ofdrw.Net.Layout.Graphics;

internal static class GraphicsValidation
{
    internal static void Finite(params double[] values)
    {
        foreach (var value in values)
            if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("Graphics coordinates must be finite.");
    }
    internal static void Positive(double value)
    {
        Finite(value);
        var written = WriterValue(value);
        Finite(written);
        if (value <= 0 || written <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Size must remain positive at OFD writer precision.");
    }
    internal static double WriterValue(double value) => double.Parse(value.ToString("0.###", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
    internal static string Number(double value) => Ofdrw.Net.Core.Models.OfdNumericFormat.Plain(value);
}
