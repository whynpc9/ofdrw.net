using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;

namespace Ofdrw.Net.Core.Models;

internal static class OfdTextGeometry
{
    internal static IReadOnlyList<string> Glyphs(string text)
    {
        var result = new List<string>();
        var enumerator = StringInfo.GetTextElementEnumerator(text);
        while (enumerator.MoveNext()) result.Add(enumerator.GetTextElement());
        return result;
    }

    internal static IReadOnlyList<double> ExpandDeltas(string? value, int maximumCount)
    {
        if (string.IsNullOrWhiteSpace(value) || maximumCount <= 0) return Array.Empty<double>();
        var tokens = value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var result = new List<double>();
        for (var index = 0; index < tokens.Length && result.Count < maximumCount;)
        {
            var count = 1;
            if (string.Equals(tokens[index], "g", StringComparison.OrdinalIgnoreCase))
            {
                if (index + 2 >= tokens.Length || !int.TryParse(tokens[index + 1], NumberStyles.Integer,
                        CultureInfo.InvariantCulture, out count) || count < 0)
                    throw new InvalidDataException("Invalid OFD repeated glyph advance.");
                index += 2;
            }
            if (!double.TryParse(tokens[index++], NumberStyles.Float, CultureInfo.InvariantCulture, out var delta) ||
                double.IsNaN(delta) || double.IsInfinity(delta))
                throw new InvalidDataException("Invalid OFD glyph advance.");
            count = Math.Min(count, maximumCount - result.Count);
            for (var repeat = 0; repeat < count; repeat++) result.Add(delta);
        }
        return result;
    }
}
