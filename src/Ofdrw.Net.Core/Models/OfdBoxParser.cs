using System;
using System.Globalization;

namespace Ofdrw.Net.Core.Models;

internal static class OfdBoxParser
{
    internal static bool TryParse(string? value, out (double x, double y, double w, double h) box)
    {
        box = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        var parts = value!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 4) return false;
        var numbers = new double[4];
        for (var i = 0; i < numbers.Length; i++)
            if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out numbers[i]) ||
                double.IsNaN(numbers[i]) || double.IsInfinity(numbers[i])) return false;
        box = (numbers[0], numbers[1], numbers[2], numbers[3]);
        return true;
    }
}
