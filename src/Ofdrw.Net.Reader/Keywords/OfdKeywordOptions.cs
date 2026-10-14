using System;

namespace Ofdrw.Net.Reader.Keywords;

/// <summary>Determines how a query handles source content whose geometry is unsupported.</summary>
public enum OfdUnsupportedContentBehavior
{
    /// <summary>Fail rather than return a partial search.</summary>
    Fail,
    /// <summary>Return diagnostics and mark the search incomplete.</summary>
    ReportIncomplete
}

/// <summary>Immutable limits for one keyword operation. Counts are cumulative, not per text object.</summary>
public sealed class OfdKeywordLimits
{
    /// <summary>Creates finite, positive limits for scanning and matching.</summary>
    public OfdKeywordLimits(int maxTextUtf16Units = 2_000_000, int maxKeywordUtf16Units = 4_096,
        int maxClusterVisits = 2_000_000, int maxHits = 10_000, long maxMatchComparisons = 4_100_000,
        int maxFontReadStreams = 128)
    {
        if (maxTextUtf16Units <= 0 || maxTextUtf16Units > 16_000_000) throw new ArgumentOutOfRangeException(nameof(maxTextUtf16Units));
        if (maxKeywordUtf16Units <= 0 || maxKeywordUtf16Units > 65_536) throw new ArgumentOutOfRangeException(nameof(maxKeywordUtf16Units));
        if (maxClusterVisits <= 0 || maxClusterVisits > 16_000_000) throw new ArgumentOutOfRangeException(nameof(maxClusterVisits));
        if (maxHits <= 0 || maxHits > 1_000_000) throw new ArgumentOutOfRangeException(nameof(maxHits));
        if (maxMatchComparisons <= 0) throw new ArgumentOutOfRangeException(nameof(maxMatchComparisons));
        if (maxFontReadStreams <= 0 || maxFontReadStreams > 4_096) throw new ArgumentOutOfRangeException(nameof(maxFontReadStreams));
        MaxTextUtf16Units = maxTextUtf16Units; MaxKeywordUtf16Units = maxKeywordUtf16Units;
        MaxClusterVisits = maxClusterVisits; MaxHits = maxHits;
        MaxMatchComparisons = maxMatchComparisons; MaxFontReadStreams = maxFontReadStreams;
    }

    /// <summary>Maximum total source UTF-16 units visited.</summary>
    public int MaxTextUtf16Units { get; }
    /// <summary>Maximum keyword length in UTF-16 units.</summary>
    public int MaxKeywordUtf16Units { get; }
    /// <summary>Maximum total verified source clusters visited.</summary>
    public int MaxClusterVisits { get; }
    /// <summary>Maximum returned occurrences, including overlapping occurrences.</summary>
    public int MaxHits { get; }
    /// <summary>Maximum ordinal character comparisons, including keyword preprocessing.</summary>
    public long MaxMatchComparisons { get; }
    /// <summary>Maximum total font streams opened during the operation, including reopened streams.</summary>
    public int MaxFontReadStreams { get; }
}

/// <summary>Immutable resource and cooperative work limits passed to a glyph provider session.</summary>
public sealed class OfdGlyphGeometryLimits
{
    /// <summary>Creates finite limits. These limits do not make third-party parser calls interruptible.</summary>
    public OfdGlyphGeometryLimits(int maxFontFaces = 64, int maxFontBytes = 32 * 1024 * 1024,
        long maxTotalFontBytes = 128L * 1024 * 1024, int maxGlyphOperations = 2_000_000,
        int maxOutlinePointVisits = 1_000_000, int maxCompositeDepth = 32)
    {
        if (maxFontFaces <= 0 || maxFontFaces > 4_096) throw new ArgumentOutOfRangeException(nameof(maxFontFaces));
        if (maxFontBytes <= 0 || maxFontBytes > 128 * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxFontBytes));
        if (maxTotalFontBytes <= 0 || maxTotalFontBytes > 512L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxTotalFontBytes));
        if (maxGlyphOperations <= 0 || maxGlyphOperations > 16_000_000) throw new ArgumentOutOfRangeException(nameof(maxGlyphOperations));
        if (maxOutlinePointVisits <= 0 || maxOutlinePointVisits > 32_000_000) throw new ArgumentOutOfRangeException(nameof(maxOutlinePointVisits));
        if (maxCompositeDepth <= 0 || maxCompositeDepth > 64) throw new ArgumentOutOfRangeException(nameof(maxCompositeDepth));
        MaxFontFaces = maxFontFaces; MaxFontBytes = maxFontBytes; MaxTotalFontBytes = maxTotalFontBytes;
        MaxGlyphOperations = maxGlyphOperations; MaxOutlinePointVisits = maxOutlinePointVisits; MaxCompositeDepth = maxCompositeDepth;
    }

    /// <summary>Maximum distinct bound font resources.</summary>
    public int MaxFontFaces { get; }
    /// <summary>Maximum original bytes in one font snapshot.</summary>
    public int MaxFontBytes { get; }
    /// <summary>Maximum cumulative original bytes copied into font snapshots.</summary>
    public long MaxTotalFontBytes { get; }
    /// <summary>Maximum glyph requests charged by a cooperating provider.</summary>
    public int MaxGlyphOperations { get; }
    /// <summary>Maximum outline visits charged by a cooperating provider.</summary>
    public int MaxOutlinePointVisits { get; }
    /// <summary>Maximum supported composite expansion depth.</summary>
    public int MaxCompositeDepth { get; }
}

/// <summary>Immutable query policies. No locator is supplied by this foundation contract.</summary>
public sealed class OfdKeywordSearchOptions
{
    /// <summary>Creates query policies; unsupported content fails by default.</summary>
    public OfdKeywordSearchOptions(OfdKeywordLimits? limits = null, OfdGlyphGeometryLimits? geometryLimits = null,
        OfdUnsupportedContentBehavior unsupportedContent = OfdUnsupportedContentBehavior.Fail)
    {
        if (unsupportedContent != OfdUnsupportedContentBehavior.Fail && unsupportedContent != OfdUnsupportedContentBehavior.ReportIncomplete)
            throw new ArgumentOutOfRangeException(nameof(unsupportedContent));
        Limits = limits ?? new OfdKeywordLimits(); GeometryLimits = geometryLimits ?? new OfdGlyphGeometryLimits();
        UnsupportedContent = unsupportedContent;
    }
    /// <summary>Gets scanning and matching limits.</summary>
    public OfdKeywordLimits Limits { get; }
    /// <summary>Gets font and provider work limits.</summary>
    public OfdGlyphGeometryLimits GeometryLimits { get; }
    /// <summary>Gets the explicit unsupported-content policy.</summary>
    public OfdUnsupportedContentBehavior UnsupportedContent { get; }
}
