using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>A genuine operation-bound keyword hit and its desired logical text.</summary>
public sealed class OfdTextReplacement
{
    public OfdTextReplacement(OfdKeywordHit hit, string replacementText)
    { Hit = hit ?? throw new ArgumentNullException(nameof(hit)); ReplacementText = replacementText ?? throw new ArgumentNullException(nameof(replacementText)); }
    public OfdKeywordHit Hit { get; }
    public string ReplacementText { get; }
}

/// <summary>Cumulative limits for one all-or-nothing replacement preparation.</summary>
public sealed class OfdTextReplacementLimits
{
    public OfdTextReplacementLimits(int maxRequests = 10_000, int maxReplacementUtf16Units = 1_000_000,
        long maxMaterializedBytes = 512L * 1024 * 1024)
    {
        if (maxRequests <= 0 || maxRequests > 10_000) throw new ArgumentOutOfRangeException(nameof(maxRequests));
        if (maxReplacementUtf16Units <= 0 || maxReplacementUtf16Units > 1_000_000) throw new ArgumentOutOfRangeException(nameof(maxReplacementUtf16Units));
        if (maxMaterializedBytes <= 0 || maxMaterializedBytes > 512L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxMaterializedBytes));
        MaxRequests = maxRequests; MaxReplacementUtf16Units = maxReplacementUtf16Units; MaxMaterializedBytes = maxMaterializedBytes;
    }
    public int MaxRequests { get; }
    public int MaxReplacementUtf16Units { get; }
    /// <summary>Separate ceiling for each storage, copy, output and validation-work ledger; not a process memory limit.</summary>
    public long MaxMaterializedBytes { get; }
}

/// <summary>Strict replacement policies. Overflow always rejects the entire batch.</summary>
public sealed class OfdTextReplacementOptions
{
    public OfdTextReplacementOptions(OfdTextReplacementLimits? limits = null, OfdKeywordSearchOptions? searchOptions = null)
    {
        Limits = limits ?? new OfdTextReplacementLimits(); SearchOptions = searchOptions ?? new OfdKeywordSearchOptions();
        if (SearchOptions.UnsupportedContent != OfdUnsupportedContentBehavior.Fail)
            throw new ArgumentException("Replacement requires complete strict source admission.", nameof(searchOptions));
    }
    public OfdTextReplacementLimits Limits { get; }
    public OfdKeywordSearchOptions SearchOptions { get; }
}

/// <summary>Copied observations about one authenticated request; retains no source package or hit.</summary>
public sealed class OfdTextReplacementObservation
{
    internal OfdTextReplacementObservation(OfdTextReplacement request)
    {
        var hit = request.Hit; PageIndex = hit.PageIndex; SourceVersion = hit.SourceVersion;
        OriginalText = hit.MatchedText; ReplacementText = request.ReplacementText;
        SourcePath = hit.Fragments[0].TextCodePath; OriginalScalarCount = hit.Fragments.Count;
        var count = 0;
        for (var i = 0; i < ReplacementText.Length; i++, count++) if (char.IsHighSurrogate(ReplacementText[i])) i++;
        ReplacementScalarCount = count;
    }
    public int PageIndex { get; }
    public string SourceVersion { get; }
    public string SourcePath { get; }
    public string OriginalText { get; }
    public string ReplacementText { get; }
    public int OriginalScalarCount { get; }
    public int ReplacementScalarCount { get; }
    public bool Changed => OriginalText != ReplacementText;
    public string Profile => "LiteralSimpleScalarsTranslationOnly";
}

/// <summary>A bounded replacement notice containing no font bytes.</summary>
public sealed class OfdTextReplacementDiagnostic
{
    internal OfdTextReplacementDiagnostic(string code) => Code = code;
    public string Code { get; }
}

/// <summary>Immutable cumulative counters captured at a preparation phase boundary.</summary>
public sealed class OfdTextReplacementPhase
{
    internal OfdTextReplacementPhase(string name, long storage, long copy, long output, long work, long references = 0, long nodes = 0)
    { Name = name; StorageBytes = storage; CopyBytes = copy; OutputUpperBoundBytes = output; ValidationWorkBytes = work;
        TraversalReferences = references; TraversalNodes = nodes; }
    public string Name { get; }
    public long StorageBytes { get; }
    /// <summary>Cumulative logical copy charges, including conservative preallocation reservations; not measured heap allocation.</summary>
    public long CopyBytes { get; }
    public long OutputUpperBoundBytes { get; }
    /// <summary>Cumulative payload-validation charges plus reserved logical byte-equivalent work; not actual font-read bytes.</summary>
    public long ValidationWorkBytes { get; }
    public long TraversalReferences { get; }
    public long TraversalNodes { get; }
}

/// <summary>Preparation counters only. Later independent writing is a separate operation.</summary>
public sealed class OfdTextReplacementStatistics
{
    internal OfdTextReplacementStatistics(OfdTextReplacementPhase[] phases, OfdKeywordLocatorStatistics[] analyses,
        OfdTextReplacementFitStatistics? fit)
    { Phases = Array.AsReadOnly(phases); Analyses = Array.AsReadOnly(analyses); Fit = fit; }
    public IReadOnlyList<OfdTextReplacementPhase> Phases { get; }
    /// <summary>Separate actual query/geometry passes; all use the same enclosing material ledger.</summary>
    public IReadOnlyList<OfdKeywordLocatorStatistics> Analyses { get; }
    public OfdTextReplacementFitStatistics? Fit { get; }
}

/// <summary>Observed pure planner actions, separately from byte-equivalent work reservations.</summary>
public sealed class OfdTextReplacementFitStatistics
{
    internal OfdTextReplacementFitStatistics(long visits, long comparisons, long outputPositions)
    { Visits = visits; Comparisons = comparisons; OutputPositions = outputPositions; }
    public long Visits { get; }
    public long Comparisons { get; }
    public long OutputPositions { get; }
}
