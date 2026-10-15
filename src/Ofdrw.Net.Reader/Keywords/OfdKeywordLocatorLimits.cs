using System;

namespace Ofdrw.Net.Reader.Keywords;

/// <summary>Finite cumulative limits for source admission, validation and returned geometry.</summary>
public sealed class OfdKeywordLocatorLimits
{
    /// <summary>Creates limits. These are cooperative work/allocation budgets, not CPU or process-memory guarantees.</summary>
    public OfdKeywordLocatorLimits(int maxPages = 1_000, int maxNodes = 100_000, int maxReferences = 200_000,
        int maxOccurrences = 100_000, int maxDepth = 64, int maxSingleXmlBytes = 16 * 1024 * 1024,
        long maxXmlBytes = 128L * 1024 * 1024, int maxDeltaValues = 4_000_000,
        int maxDeltaTokens = 4_000_000, int maxFragments = 100_000, int maxDiagnostics = 10_000,
        long maxValidationFields = 10_000_000, long maxValidationCharacters = 64_000_000,
        long maxValidationBytes = 512L * 1024 * 1024)
    {
        if (maxPages <= 0 || maxPages > 100_000 || maxNodes <= 0 || maxNodes > 4_000_000 ||
            maxReferences <= 0 || maxReferences > 8_000_000 || maxOccurrences <= 0 || maxOccurrences > 1_000_000 ||
            maxDepth <= 0 || maxDepth > 128 || maxSingleXmlBytes <= 0 || maxSingleXmlBytes > 32 * 1024 * 1024 ||
            maxXmlBytes <= 0 || maxXmlBytes > 512L * 1024 * 1024 || maxDeltaValues <= 0 || maxDeltaValues > 32_000_000 ||
            maxDeltaTokens <= 0 || maxDeltaTokens > 32_000_000 || maxFragments <= 0 || maxFragments > 1_000_000 ||
            maxDiagnostics <= 0 || maxDiagnostics > 100_000 || maxValidationFields <= 0 ||
            maxValidationCharacters <= 0 || maxValidationBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(maxPages), "Locator limits must be positive and within hard ceilings.");
        MaxPages = maxPages; MaxNodes = maxNodes; MaxReferences = maxReferences; MaxOccurrences = maxOccurrences;
        MaxDepth = maxDepth; MaxSingleXmlBytes = maxSingleXmlBytes; MaxXmlBytes = maxXmlBytes;
        MaxDeltaValues = maxDeltaValues; MaxDeltaTokens = maxDeltaTokens; MaxFragments = maxFragments;
        MaxDiagnostics = maxDiagnostics; MaxValidationFields = maxValidationFields;
        MaxValidationCharacters = maxValidationCharacters; MaxValidationBytes = maxValidationBytes;
    }
    /// <summary>Maximum selected page positions.</summary>
    public int MaxPages { get; }
    /// <summary>Maximum nodes per A0 traversal and cumulative extra admission nodes.</summary>
    public int MaxNodes { get; }
    /// <summary>Maximum references per A0 traversal and cumulative extra admission references.</summary>
    public int MaxReferences { get; }
    /// <summary>Maximum A0 occurrences plus diagnostics.</summary>
    public int MaxOccurrences { get; }
    /// <summary>Maximum XML/traversal depth.</summary>
    public int MaxDepth { get; }
    /// <summary>Maximum one control XML byte length before parsing.</summary>
    public int MaxSingleXmlBytes { get; }
    /// <summary>Maximum cumulative extra admission XML bytes before parsing.</summary>
    public long MaxXmlBytes { get; }
    /// <summary>Maximum cumulative expanded delta values, including missing-axis zero arrays.</summary>
    public int MaxDeltaValues { get; }
    /// <summary>Maximum cumulative delta tokens consumed.</summary>
    public int MaxDeltaTokens { get; }
    /// <summary>Maximum output cluster fragments, including repetitions in overlapping hits.</summary>
    public int MaxFragments { get; }
    /// <summary>Maximum returned diagnostics; excess is fatal rather than truncated.</summary>
    public int MaxDiagnostics { get; }
    /// <summary>Maximum cumulative fields compared during snapshot validation.</summary>
    public long MaxValidationFields { get; }
    /// <summary>Maximum cumulative string units compared during snapshot validation.</summary>
    public long MaxValidationCharacters { get; }
    /// <summary>Maximum cumulative envelope bytes compared during snapshot validation.</summary>
    public long MaxValidationBytes { get; }
}
