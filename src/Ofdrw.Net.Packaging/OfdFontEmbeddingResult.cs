namespace Ofdrw.Net.Packaging;

/// <summary>Observable payload decision for one source content identity.</summary>
public sealed class OfdFontEmbeddingResult
{
    /// <summary>SHA-256 of the supplied source bytes, independent of family name or style alias.</summary>
    public string SourceIdentity { get; internal set; } = string.Empty;
    /// <summary>SHA-256 of the written payload.</summary>
    public string PayloadIdentity { get; internal set; } = string.Empty;
    /// <summary>Source font byte count.</summary>
    public long SourceBytes { get; internal set; }
    /// <summary>Written font byte count.</summary>
    public long PayloadBytes { get; internal set; }
    /// <summary>Number of resource aliases sharing this source.</summary>
    public int ResourceCount { get; internal set; }
    /// <summary>True when actual glyph outlines were removed.</summary>
    public bool IsSubset { get; internal set; }
    /// <summary>Glyph count retained by closure; null for full preservation.</summary>
    public int? RetainedGlyphs { get; internal set; }
    /// <summary>Original glyph count; null for full preservation.</summary>
    public int? OriginalGlyphs { get; internal set; }
    /// <summary>Reason for full preservation, or null for a subset.</summary>
    public string? PreservationReason { get; internal set; }
}
