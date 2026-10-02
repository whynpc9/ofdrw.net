namespace Ofdrw.Net.Core.Models;

/// <summary>Controls embedded font payloads when saving an OFD package.</summary>
public sealed class OfdFontEmbeddingOptions
{
    /// <summary>Subset modeled static TrueType usage; preserve unknown content with diagnostics.</summary>
    public OfdFontEmbeddingMode Mode { get; set; } = OfdFontEmbeddingMode.SubsetWhenSafe;

    /// <summary>Maximum source bytes for each face processed by the subset backend.</summary>
    public long MaximumFontBytes { get; set; } = 64L * 1024 * 1024;

    /// <summary>Maximum distinct Unicode scalars in a face's package-wide usage closure.</summary>
    public int MaximumUsedScalars { get; set; } = 65535;

    /// <summary>Maximum total GSUB/composite closure operations per face, including repeated passes.</summary>
    public long MaximumGlyphClosureOperations { get; set; } = 2_000_000;
}

/// <summary>Font payload policy. Neither mode grants permission to embed a font.</summary>
public enum OfdFontEmbeddingMode
{
    /// <summary>Aggregate package usage, validate coverage, and subset supported faces safely.</summary>
    SubsetWhenSafe,

    /// <summary>Explicitly preserve the previous full embedding behavior without coverage validation.</summary>
    Full
}
