using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace Ofdrw.Net.Core.Models;

/// <summary>One actual font glyph, its logical UTF-16 cluster and its OFD-local origin.</summary>
public sealed class OfdPositionedGlyph
{
    /// <summary>Creates a glyph positioned in millimeters relative to its TextObject.</summary>
    public OfdPositionedGlyph(uint glyphId, int clusterStart, int clusterLength,
        double originXMillimeters, double originYMillimeters,
        double xAdvanceMillimeters, double yAdvanceMillimeters)
    {
        if (glyphId is 0 or > ushort.MaxValue || clusterStart < 0 || clusterLength < 1 ||
            !Finite(originXMillimeters) || !Finite(originYMillimeters) ||
            !Finite(xAdvanceMillimeters) || !Finite(yAdvanceMillimeters))
            throw new ArgumentException("Invalid positioned glyph or source cluster.");
        GlyphId = glyphId; ClusterStart = clusterStart; ClusterLength = clusterLength;
        OriginXMillimeters = originXMillimeters; OriginYMillimeters = originYMillimeters;
        XAdvanceMillimeters = xAdvanceMillimeters; YAdvanceMillimeters = yAdvanceMillimeters;
    }

    /// <summary>Actual original GID in the bound face.</summary>
    public uint GlyphId { get; }
    /// <summary>Absolute UTF-16 start in the caller's logical source.</summary>
    public int ClusterStart { get; }
    /// <summary>UTF-16 length associated with this glyph; clusters may be shared.</summary>
    public int ClusterLength { get; }
    /// <summary>Glyph origin X relative to TextObject, in millimeters.</summary>
    public double OriginXMillimeters { get; }
    /// <summary>Glyph origin Y relative to TextObject, in millimeters.</summary>
    public double OriginYMillimeters { get; }
    /// <summary>Final horizontal advance in millimeters, distinct from origin deltas.</summary>
    public double XAdvanceMillimeters { get; }
    /// <summary>Final vertical advance in millimeters, distinct from origin deltas.</summary>
    public double YAdvanceMillimeters { get; }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}

/// <summary>How the OFD writer maps a positioned run's logical text to its glyphs.</summary>
public enum OfdPositionedGlyphMapping
{
    /// <summary>The historical bounded word profiles (one CGTransform for a short Latin/Hebrew TextCode).</summary>
    BoundedWord = 0,
    /// <summary>One CGTransform per source cluster, in logical order, covering the complete TextCode,
    /// with an explicit origin for every glyph slot. CodePosition and CodeCount count Unicode scalar
    /// values, so supplementary characters count once. Used by final paragraph layout.</summary>
    ClusterMap = 1,
}

/// <summary>Typed producer-owned glyph data for one logical TextCode and one immutable font face.</summary>
public sealed class OfdPositionedGlyphRun
{
    /// <summary>Freezes glyphs and their logical source interval; output profiles are checked by the writer.</summary>
    public OfdPositionedGlyphRun(string logicalText, int sourceStart, string faceSha256,
        IReadOnlyList<OfdPositionedGlyph> glyphs)
        : this(logicalText, sourceStart, faceSha256, glyphs, OfdPositionedGlyphMapping.BoundedWord)
    {
    }

    /// <summary>Freezes glyphs with an explicit writer mapping. <see cref="OfdPositionedGlyphMapping.ClusterMap"/>
    /// requires text whose glyph clusters partition it, each cluster's glyphs contiguous.</summary>
    public OfdPositionedGlyphRun(string logicalText, int sourceStart, string faceSha256,
        IReadOnlyList<OfdPositionedGlyph> glyphs, OfdPositionedGlyphMapping mapping)
    {
        if (mapping != OfdPositionedGlyphMapping.BoundedWord && mapping != OfdPositionedGlyphMapping.ClusterMap)
            throw new ArgumentOutOfRangeException(nameof(mapping));
        if (logicalText is null) throw new ArgumentNullException(nameof(logicalText));
        if (faceSha256 is null) throw new ArgumentNullException(nameof(faceSha256));
        if (glyphs is null) throw new ArgumentNullException(nameof(glyphs));
        if (sourceStart < 0 || (long)sourceStart + logicalText.Length > int.MaxValue ||
            logicalText.Length is < 1 or > 4096 || glyphs.Count is < 1 or > 16384 ||
            faceSha256.Length != 64)
            throw new ArgumentException("Invalid positioned run source, face or glyph budget.");
        foreach (var value in faceSha256)
            if (!Uri.IsHexDigit(value)) throw new ArgumentException("Face identity must be SHA-256 hex.", nameof(faceSha256));
        var copy = new List<OfdPositionedGlyph>(glyphs.Count);
        var covered = new bool[logicalText.Length];
        foreach (var glyph in glyphs)
        {
            if (glyph is null || glyph.ClusterStart < sourceStart ||
                (long)glyph.ClusterStart + glyph.ClusterLength > (long)sourceStart + logicalText.Length)
                throw new ArgumentException("A glyph cluster lies outside the logical source.", nameof(glyphs));
            var first = glyph.ClusterStart - sourceStart;
            var end = first + glyph.ClusterLength;
            if (first > 0 && char.IsLowSurrogate(logicalText[first]) ||
                end < logicalText.Length && char.IsLowSurrogate(logicalText[end]))
                throw new ArgumentException("A glyph cluster splits a surrogate pair.", nameof(glyphs));
            for (var index = first; index < end; index++) covered[index] = true;
            copy.Add(glyph);
        }
        foreach (var value in covered)
            if (!value) throw new ArgumentException("A logical source interval has no glyph association.", nameof(glyphs));
        if (mapping == OfdPositionedGlyphMapping.ClusterMap) ValidateClusterMap(logicalText, sourceStart, copy);
        LogicalText = logicalText; SourceStart = sourceStart;
        FaceSha256 = faceSha256.ToLowerInvariant();
        Glyphs = new ReadOnlyCollection<OfdPositionedGlyph>(copy);
        Mapping = mapping;
    }

    /// <summary>Writer mapping of this run.</summary>
    public OfdPositionedGlyphMapping Mapping { get; }

    /// <summary>Glyph groups ordered by logical cluster start: (start, length, first glyph index, glyph count).</summary>
    internal IReadOnlyList<(int Start, int Length, int First, int Count)> LogicalClusters()
    {
        var groups = new List<(int Start, int Length, int First, int Count)>();
        for (var index = 0; index < Glyphs.Count;)
        {
            var end = index + 1;
            while (end < Glyphs.Count && Glyphs[end].ClusterStart == Glyphs[index].ClusterStart) end++;
            groups.Add((Glyphs[index].ClusterStart, Glyphs[index].ClusterLength, index, end - index));
            index = end;
        }
        groups.Sort((left, right) => left.Start.CompareTo(right.Start));
        return groups;
    }

    private static void ValidateClusterMap(string text, int sourceStart, List<OfdPositionedGlyph> glyphs)
    {
        // CodePosition/CodeCount count Unicode scalar values (XML characters); the constructor already
        // rejects clusters that split a surrogate pair, so only unpaired surrogates remain to refuse.
        for (var index = 0; index < text.Length; index++)
        {
            if (char.IsHighSurrogate(text[index]) && index + 1 < text.Length && char.IsLowSurrogate(text[index + 1])) index++;
            else if (char.IsSurrogate(text[index]))
                throw new ArgumentException("A cluster-mapped run cannot contain an unpaired surrogate.", nameof(text));
        }
        var seen = new HashSet<int>();
        var covered = 0;
        for (var index = 0; index < glyphs.Count;)
        {
            var start = glyphs[index].ClusterStart;
            var length = glyphs[index].ClusterLength;
            if (!seen.Add(start))
                throw new ArgumentException("A cluster's glyphs must be contiguous.", nameof(glyphs));
            var end = index + 1;
            while (end < glyphs.Count && glyphs[end].ClusterStart == start)
            {
                if (glyphs[end].ClusterLength != length)
                    throw new ArgumentException("Glyphs of one cluster disagree about its length.", nameof(glyphs));
                end++;
            }
            covered += length;
            index = end;
        }
        // Distinct cluster starts whose lengths sum to the text length, all within it and covering it
        // (checked by the caller), cannot overlap.
        if (covered != text.Length)
            throw new ArgumentException("Cluster-mapped glyph clusters must partition the logical text.", nameof(glyphs));
    }

    /// <summary>Unmodified logical Unicode of this TextCode.</summary>
    public string LogicalText { get; }
    /// <summary>Absolute UTF-16 start in the caller's original source.</summary>
    public int SourceStart { get; }
    /// <summary>SHA-256 of the exact standalone face used for shaping.</summary>
    public string FaceSha256 { get; }
    /// <summary>Actual glyph IDs, clusters and origins in final visual order.</summary>
    public IReadOnlyList<OfdPositionedGlyph> Glyphs { get; }
}
