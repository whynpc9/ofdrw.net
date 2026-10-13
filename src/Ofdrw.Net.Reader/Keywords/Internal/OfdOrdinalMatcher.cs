using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;

namespace Ofdrw.Net.Reader.Keywords.Internal;

internal readonly struct OfdKeywordMatch
{
    internal OfdKeywordMatch(int start, int length, int clusterStart, int clusterCount)
    { Utf16Start = start; Utf16Length = length; ClusterStart = clusterStart; ClusterCount = clusterCount; }
    internal int Utf16Start { get; }
    internal int Utf16Length { get; }
    internal int ClusterStart { get; }
    internal int ClusterCount { get; }
}

// KMP with one prefix table per operation. It consumes proven mappings; it never invents glyph clusters.
internal sealed class OfdOrdinalMatcher
{
    private readonly string keyword;
    private readonly int[] prefix;
    private readonly OfdKeywordOperation operation;

    internal OfdOrdinalMatcher(string keyword, OfdKeywordOperation operation)
    {
        this.operation = operation ?? throw new ArgumentNullException(nameof(operation)); operation.Check();
        if (string.IsNullOrEmpty(keyword)) throw new ArgumentException("Keyword must be nonempty.", nameof(keyword));
        if (keyword.Length > operation.Options.Limits.MaxKeywordUtf16Units) throw new InvalidDataException("Keyword length capacity exceeded.");
        ValidateUnicode(keyword, operation); this.keyword = keyword; prefix = new int[keyword.Length];
        var matched = 0;
        for (var i = 1; i < keyword.Length; i++)
        {
            while (matched > 0 && !Equal(keyword[i], keyword[matched])) matched = prefix[matched - 1];
            if (Equal(keyword[i], keyword[matched])) matched++;
            prefix[i] = matched;
        }
        operation.Check();
    }

    internal IReadOnlyList<OfdKeywordMatch> Find(string text, IReadOnlyList<OfdGlyphClusterGeometry> clusters)
    {
        operation.Check();
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (clusters is null) throw new ArgumentNullException(nameof(clusters));
        operation.ChargeText(text.Length, clusters.Count);
        ValidateUnicode(text, operation);
        // Bounded by charged UTF-16 units; each valid source boundary stores its cluster ordinal.
        var boundaries = new int[text.Length + 1];
        for (var i = 0; i < boundaries.Length; i++) { if ((i & 255) == 0) operation.Check(); boundaries[i] = -1; }
        var cursor = 0; boundaries[0] = 0;
        for (var i = 0; i < clusters.Count; i++)
        {
            if ((i & 255) == 0) operation.Check();
            var cluster = clusters[i] ?? throw new InvalidDataException("Missing cluster mapping.");
            if (cluster.Utf16Start != cursor || cluster.Utf16Length > text.Length - cursor)
                throw new InvalidDataException("Cluster mappings must cover source text exactly and in order.");
            cursor += cluster.Utf16Length;
            if (cursor < text.Length && char.IsLowSurrogate(text[cursor])) throw new InvalidDataException("Cluster splits a Unicode scalar.");
            boundaries[cursor] = i + 1;
        }
        if (cursor != text.Length) throw new InvalidDataException("Cluster mappings do not cover source text.");

        var result = new List<OfdKeywordMatch>(); var matched = 0;
        for (var i = 0; i < text.Length; i++)
        {
            while (matched > 0 && !Equal(text[i], keyword[matched])) matched = prefix[matched - 1];
            if (Equal(text[i], keyword[matched])) matched++;
            if (matched != keyword.Length) continue;
            var start = i + 1 - matched;
            if (boundaries[start] >= 0 && boundaries[i + 1] >= 0)
            {
                operation.ChargeHit();
                result.Add(new OfdKeywordMatch(start, matched, boundaries[start], boundaries[i + 1] - boundaries[start]));
            }
            matched = prefix[matched - 1]; // retain overlapping occurrences in source order.
        }
        operation.Check(); return new ReadOnlyCollection<OfdKeywordMatch>(result);
    }

    private bool Equal(char left, char right) { operation.ChargeComparison(); return left == right; }
    internal static void ValidateUnicode(string value, OfdKeywordOperation operation)
    {
        var nextCheck = 0;
        for (var i = 0; i < value.Length; i++)
        {
            if (i >= nextCheck) { operation.Check(); nextCheck = i + 256; }
            if (char.IsHighSurrogate(value[i]))
            {
                if (++i >= value.Length || !char.IsLowSurrogate(value[i])) throw new InvalidDataException("Unpaired UTF-16 surrogate.");
            }
            else if (char.IsLowSurrogate(value[i])) throw new InvalidDataException("Unpaired UTF-16 surrogate.");
        }
        operation.Check();
    }
}
