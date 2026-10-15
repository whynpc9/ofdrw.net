using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal static class OfdTextGeometry
{
    internal static IReadOnlyList<string> Glyphs(string text)
    {
        var result = new List<string>();
        foreach (var glyph in EnumerateGlyphs(text, default)) result.Add(glyph);
        return result;
    }

    internal static IEnumerable<string> EnumerateGlyphs(string text, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var enumerator = StringInfo.GetTextElementEnumerator(text);
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!enumerator.MoveNext()) break;
            yield return enumerator.GetTextElement();
        }
        cancellationToken.ThrowIfCancellationRequested();
    }

    internal static bool IsWhiteSpace(string text, CancellationToken cancellationToken)
    {
        for (var index = 0; index < text.Length; index++)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!char.IsWhiteSpace(text[index])) return false;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return true;
    }

    internal static (IReadOnlyList<double> X, IReadOnlyList<double> Y) PositioningDeltas(
        OfdTextRun run, CancellationToken cancellationToken)
    {
        // The common unpositioned case never enumerates or materializes graphemes.
        // A nonempty expression (including g 0) still goes through the existing parser.
        var noX = run.DeltaX is null || IsWhiteSpace(run.DeltaX, cancellationToken);
        var noY = run.DeltaY is null || IsWhiteSpace(run.DeltaY, cancellationToken);
        cancellationToken.ThrowIfCancellationRequested();
        if (noX && noY) return (Array.Empty<double>(), Array.Empty<double>());
        var count = 0;
        var enumerator = StringInfo.GetTextElementEnumerator(run.Text);
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!enumerator.MoveNext()) break;
            count++; // Do not call GetTextElement: counting requires no per-glyph strings.
        }
        var maximumCount = Math.Max(0, count - 1);
        return (ExpandDeltas(run.DeltaX, maximumCount, cancellationToken),
            ExpandDeltas(run.DeltaY, maximumCount, cancellationToken));
    }

    internal static IReadOnlyList<double> ExpandDeltas(string? value, int maximumCount)
        => ExpandDeltas(value, maximumCount, default);

    internal static IReadOnlyList<double> ExpandDeltas(string? value, int maximumCount, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (maximumCount <= 0 || value is null || IsWhiteSpace(value, cancellationToken)) return Array.Empty<double>();
        var tokens = value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var result = new List<double>();
        for (var index = 0; index < tokens.Length && result.Count < maximumCount;)
        {
            cancellationToken.ThrowIfCancellationRequested();
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
            for (var repeat = 0; repeat < count; repeat++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                result.Add(delta);
            }
        }
        return result;
    }
}
