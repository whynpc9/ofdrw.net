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
        if (value <= 0 || Math.Round(value, 3) <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Size must remain positive at OFD writer precision.");
    }
    internal static string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
}
