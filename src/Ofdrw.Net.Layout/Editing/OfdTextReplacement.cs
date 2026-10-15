using Ofdrw.Net.Reader.Keywords;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>A genuine operation-bound keyword hit and its desired logical text.</summary>
public sealed class OfdTextReplacement
{
    /// <summary>Binds a genuine keyword hit to non-null replacement text.</summary>
    public OfdTextReplacement(OfdKeywordHit hit, string replacementText)
    { Hit = hit ?? throw new ArgumentNullException(nameof(hit)); ReplacementText = replacementText ?? throw new ArgumentNullException(nameof(replacementText)); }
    /// <summary>Authenticated operation-bound keyword hit to replace.</summary>
    public OfdKeywordHit Hit { get; }
    /// <summary>Desired logical replacement text; it must satisfy the supported literal-scalar profile.</summary>
    public string ReplacementText { get; }
}

/// <summary>Cumulative limits for one all-or-nothing replacement preparation.</summary>
public sealed class OfdTextReplacementLimits
{
    /// <summary>Creates validated cumulative limits; each byte ledger is capped independently.</summary>
    public OfdTextReplacementLimits(int maxRequests = 10_000, int maxReplacementUtf16Units = 1_000_000,
        long maxMaterializedBytes = 512L * 1024 * 1024)
    {
        if (maxRequests <= 0 || maxRequests > 10_000) throw new ArgumentOutOfRangeException(nameof(maxRequests));
        if (maxReplacementUtf16Units <= 0 || maxReplacementUtf16Units > 1_000_000) throw new ArgumentOutOfRangeException(nameof(maxReplacementUtf16Units));
        if (maxMaterializedBytes <= 0 || maxMaterializedBytes > 512L * 1024 * 1024) throw new ArgumentOutOfRangeException(nameof(maxMaterializedBytes));
        MaxRequests = maxRequests; MaxReplacementUtf16Units = maxReplacementUtf16Units; MaxMaterializedBytes = maxMaterializedBytes;
    }
    /// <summary>Maximum requests admitted by one replacement batch, capped at 10,000.</summary>
    public int MaxRequests { get; }
    /// <summary>Maximum cumulative replacement UTF-16 units, capped at 1,000,000.</summary>
    public int MaxReplacementUtf16Units { get; }
    /// <summary>Separate ceiling for each storage, copy, output and validation-work ledger; not a process memory limit.</summary>
    public long MaxMaterializedBytes { get; }
}

/// <summary>Strict replacement policies. Overflow always rejects the entire batch.</summary>
public sealed class OfdTextReplacementOptions
{
    /// <summary>Creates a strict all-or-nothing replacement policy; unsupported-content skipping is rejected.</summary>
    public OfdTextReplacementOptions(OfdTextReplacementLimits? limits = null, OfdKeywordSearchOptions? searchOptions = null)
    {
        Limits = limits ?? new OfdTextReplacementLimits(); SearchOptions = searchOptions ?? new OfdKeywordSearchOptions();
        if (SearchOptions.UnsupportedContent != OfdUnsupportedContentBehavior.Fail)
            throw new ArgumentException("Replacement requires complete strict source admission.", nameof(searchOptions));
    }
    /// <summary>Immutable cumulative replacement limits.</summary>
    public OfdTextReplacementLimits Limits { get; }
    /// <summary>Strict source-admission and keyword-search policies.</summary>
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
    /// <summary>Zero-based source page index for this request.</summary>
    public int PageIndex { get; }
    /// <summary>Captured source-version identity for the authenticated hit.</summary>
    public string SourceVersion { get; }
    /// <summary>Captured TextCode path identifying the original source slot.</summary>
    public string SourcePath { get; }
    /// <summary>Exact logical text matched before replacement.</summary>
    public string OriginalText { get; }
    /// <summary>Desired logical replacement text; it must satisfy the supported literal-scalar profile.</summary>
    public string ReplacementText { get; }
    /// <summary>Number of captured source scalar fragments in the hit.</summary>
    public int OriginalScalarCount { get; }
    /// <summary>Number of Unicode scalar values in the replacement.</summary>
    public int ReplacementScalarCount { get; }
    /// <summary>Whether at least one requested logical string differs from its original text.</summary>
    public bool Changed => OriginalText != ReplacementText;
    /// <summary>Name of the restricted literal-scalar translation-only editing profile.</summary>
    public string Profile => "LiteralSimpleScalarsTranslationOnly";
}

/// <summary>A bounded replacement notice containing no font bytes.</summary>
public sealed class OfdTextReplacementDiagnostic
{
    internal OfdTextReplacementDiagnostic(string code) => Code = code;
    /// <summary>Stable diagnostic code for this replacement operation.</summary>
    public string Code { get; }
}

/// <summary>Immutable cumulative counters captured at a preparation phase boundary.</summary>
public sealed class OfdTextReplacementPhase
{
    internal OfdTextReplacementPhase(string name, long storage, long copy, long output, long work, long references = 0, long nodes = 0)
    { Name = name; StorageBytes = storage; CopyBytes = copy; OutputUpperBoundBytes = output; ValidationWorkBytes = work;
        TraversalReferences = references; TraversalNodes = nodes; }
    /// <summary>Preparation phase name at which these cumulative counters were captured.</summary>
    public string Name { get; }
    /// <summary>Cumulative logically retained input-storage bytes at this phase.</summary>
    public long StorageBytes { get; }
    /// <summary>Cumulative logical copy charges, including conservative preallocation reservations; not measured heap allocation.</summary>
    public long CopyBytes { get; }
    /// <summary>Conservative uncompressed output-byte reservation at this phase.</summary>
    public long OutputUpperBoundBytes { get; }
    /// <summary>Cumulative payload-validation charges plus reserved logical byte-equivalent work; not actual font-read bytes.</summary>
    public long ValidationWorkBytes { get; }
    /// <summary>Cumulative visited model references.</summary>
    public long TraversalReferences { get; }
    /// <summary>Cumulative traversed model and XML nodes.</summary>
    public long TraversalNodes { get; }
}

/// <summary>Preparation counters only. Later independent writing is a separate operation.</summary>
public sealed class OfdTextReplacementStatistics
{
    internal OfdTextReplacementStatistics(OfdTextReplacementPhase[] phases, OfdKeywordLocatorStatistics[] analyses,
        OfdTextReplacementFitStatistics? fit)
    { Phases = Array.AsReadOnly(phases); Analyses = Array.AsReadOnly(analyses); Fit = fit; }
    /// <summary>Ordered immutable preparation-phase counter snapshots.</summary>
    public IReadOnlyList<OfdTextReplacementPhase> Phases { get; }
    /// <summary>Separate actual query/geometry passes; all use the same enclosing material ledger.</summary>
    public IReadOnlyList<OfdKeywordLocatorStatistics> Analyses { get; }
    /// <summary>Pure fit-planner counters, or null when no fit was required.</summary>
    public OfdTextReplacementFitStatistics? Fit { get; }
}

/// <summary>Observed pure planner actions, separately from byte-equivalent work reservations.</summary>
public sealed class OfdTextReplacementFitStatistics
{
    internal OfdTextReplacementFitStatistics(long visits, long comparisons, long outputPositions)
    { Visits = visits; Comparisons = comparisons; OutputPositions = outputPositions; }
    /// <summary>Number of pure fit-planner visits.</summary>
    public long Visits { get; }
    /// <summary>Number of pure fit-planner comparisons.</summary>
    public long Comparisons { get; }
    /// <summary>Number of positions emitted by the fit planner.</summary>
    public long OutputPositions { get; }
}
