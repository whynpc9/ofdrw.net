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

/// <summary>Typed producer-owned glyph data for one logical TextCode and one immutable font face.</summary>
public sealed class OfdPositionedGlyphRun
{
    /// <summary>Freezes glyphs and their logical source interval; output profiles are checked by the writer.</summary>
    public OfdPositionedGlyphRun(string logicalText, int sourceStart, string faceSha256,
        IReadOnlyList<OfdPositionedGlyph> glyphs)
    {
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
        LogicalText = logicalText; SourceStart = sourceStart;
        FaceSha256 = faceSha256.ToLowerInvariant();
        Glyphs = new ReadOnlyCollection<OfdPositionedGlyph>(copy);
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
