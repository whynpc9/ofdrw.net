using System;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Core.Validation;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>The result of a finite local observation, never a certification decision.</summary>
public enum PreflightOutcome
{
    /// <summary>The declared local condition was observed.</summary>
    Pass,
    /// <summary>A definite violation of the local condition was observed.</summary>
    Fail,
    /// <summary>Complete observation established no applicable objects, or policy disabled the rule.</summary>
    NotApplicable,
    /// <summary>Evidence or coverage is insufficient.</summary>
    Unknown
}

/// <summary>Whether the requested bounded inspection finished.</summary>
public enum PreflightRunState
{
    /// <summary>Inspection finished; individual outcomes can still be unknown.</summary>
    Complete,
    /// <summary>Input rejection or a budget stopped inspection.</summary>
    Incomplete
}

/// <summary>Local policy and resource budgets, snapshotted before inspecting input.</summary>
public sealed class ArchivalPreflightOptions
{
    /// <summary>Require exactly one document with a page.</summary>
    public bool RequireSingleDocument { get; set; } = true;
    /// <summary>Require absence of the known encryption markers.</summary>
    public bool RequireNoEncryption { get; set; } = true;
    /// <summary>ZIP loading budgets. The checker copies these values and does not mutate this instance.</summary>
    public OfdPackageLoadOptions LoadOptions { get; set; } = new OfdPackageLoadOptions();
    /// <summary>Maximum cumulative bytes of distinct XML entries inspected.</summary>
    public long MaxXmlBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum bytes in one XML entry, checked before parsing.</summary>
    public int MaxSingleXmlBytes { get; set; } = 8 * 1024 * 1024;
    /// <summary>Maximum XML reader depth.</summary>
    public int MaxXmlDepth { get; set; } = 64;
    /// <summary>Maximum cumulative XML reader nodes.</summary>
    public int MaxXmlNodes { get; set; } = 200_000;
    /// <summary>Maximum known reference visits, including repeated references.</summary>
    public int MaxReferences { get; set; } = 100_000;
    /// <summary>Maximum results, including six reserved interruption diagnostics; must be at least twelve.</summary>
    public int MaxResults { get; set; } = 10_000;

    internal ArchivalPreflightOptions Snapshot()
    {
        if (LoadOptions == null) throw new ArgumentException("LoadOptions is required.", nameof(LoadOptions));
        var load = LoadOptions;
        if (MaxXmlBytes <= 0 || MaxSingleXmlBytes <= 0 || MaxXmlDepth <= 0 || MaxXmlNodes <= 0 ||
            MaxReferences <= 0 || MaxResults < 12) throw new ArgumentOutOfRangeException(nameof(ArchivalPreflightOptions));
        return new ArchivalPreflightOptions
        {
            RequireSingleDocument = RequireSingleDocument, RequireNoEncryption = RequireNoEncryption,
            MaxXmlBytes = MaxXmlBytes, MaxSingleXmlBytes = MaxSingleXmlBytes, MaxXmlDepth = MaxXmlDepth,
            MaxXmlNodes = MaxXmlNodes, MaxReferences = MaxReferences, MaxResults = MaxResults,
            LoadOptions = new OfdPackageLoadOptions
            {
                MaxInputBytes = load.MaxInputBytes, MaxEntryCount = load.MaxEntryCount,
                MaxEntryUncompressedBytes = load.MaxEntryUncompressedBytes,
                MaxTotalUncompressedBytes = load.MaxTotalUncompressedBytes, MaxCompressionRatio = load.MaxCompressionRatio,
                MaxPageCount = load.MaxPageCount, MaxAnnotationObjectCount = load.MaxAnnotationObjectCount,
                MaxAnnotationXmlBytes = load.MaxAnnotationXmlBytes
            }
        };
    }
}

/// <summary>Namespace-aware location and bounded metadata, with no payload content.</summary>
public sealed class PreflightEvidence
{
    internal PreflightEvidence(string entry, string selector, string observed, string expected)
    { Entry = entry; Selector = selector; Observed = observed; Expected = expected; }
    /// <summary>Report-local evidence identity; not a persistent object identifier.</summary>
    public string EvidenceId { get; internal set; } = string.Empty;
    /// <summary>Report-local observation this evidence supports.</summary>
    public string ResultId { get; internal set; } = string.Empty;
    /// <summary>Entry name or the literal &lt;input&gt; for input-level diagnostics.</summary>
    public string Entry { get; }
    /// <summary>Expanded-name selector with one-based sibling indices, or package-level marker.</summary>
    public string Selector { get; }
    /// <summary>Sanitized observed metadata.</summary>
    public string Observed { get; }
    /// <summary>The finite local expectation.</summary>
    public string Expected { get; }
}

/// <summary>One observation about an object in the declared local scope.</summary>
public sealed class PreflightRuleResult
{
    internal PreflightRuleResult(string ruleId, PreflightOutcome outcome, string reasonCode, PreflightEvidence evidence)
    { RuleId = ruleId; Outcome = outcome; ReasonCode = reasonCode; Evidence = evidence; }
    /// <summary>Report-local identity assigned after stable sorting.</summary>
    public string ResultId { get; internal set; } = string.Empty;
    /// <summary>Local APF identifier.</summary>
    public string RuleId { get; }
    /// <summary>Local rule revision.</summary>
    public int RuleVersion => 1;
    /// <summary>Four-state observation.</summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Stable diagnostic reason; also explains applicability.</summary>
    public string ReasonCode { get; }
    /// <summary>Applicability explanation, expressed by the stable reason code and evidence.</summary>
    public string ApplicabilityReason => ReasonCode;
    /// <summary>Repository-local rule documentation.</summary>
    public string Source => "docs/archival-preflight.md";
    /// <summary>Severity; signature inventory failures are warnings but remain failures.</summary>
    public ValidationSeverity Severity => Outcome == PreflightOutcome.Unknown ? ValidationSeverity.Warning :
        Outcome == PreflightOutcome.Fail ? (RuleId == "APF-006" ? ValidationSeverity.Warning : ValidationSeverity.Error) : ValidationSeverity.Info;
    /// <summary>Evidence linked by this result's report-local identity.</summary>
    public PreflightEvidence Evidence { get; }
    /// <summary>Action for the caller; no automatic repair is performed.</summary>
    public string Recommendation => Outcome == PreflightOutcome.Fail || Outcome == PreflightOutcome.Unknown
        ? "Review the declaration and coverage before applying your acceptance policy." : string.Empty;
    /// <summary>Rule source kind.</summary>
    public string SourceKind => "LocalPolicy";
    /// <summary>No standard mapping has been verified.</summary>
    public string StandardStatus => "Unverified";
    /// <summary>Empty until a standard mapping is independently verified.</summary>
    public string Clause => string.Empty;
}

/// <summary>Read-only local preflight report. Pass covers only the explicitly listed observations.</summary>
public sealed class ArchivalPreflightReport
{
    internal ArchivalPreflightReport(PreflightRunState state, IEnumerable<PreflightRuleResult> results,
        IEnumerable<string> gaps, int entryCount, int documentCount, ArchivalPreflightOptions options)
    {
        RunState = state;
        var sorted = results.OrderBy(x => x.RuleId, StringComparer.Ordinal).ThenBy(x => x.Evidence.Entry, StringComparer.Ordinal)
            .ThenBy(x => x.Evidence.Selector, StringComparer.Ordinal).ThenBy(x => x.ReasonCode, StringComparer.Ordinal).ToList();
        for (var i = 0; i < sorted.Count; i++)
        {
            sorted[i].ResultId = "R" + (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
            sorted[i].Evidence.ResultId = sorted[i].ResultId;
            sorted[i].Evidence.EvidenceId = sorted[i].ResultId + "-E1";
        }
        RuleResults = sorted.AsReadOnly(); CoverageGaps = gaps.Distinct(StringComparer.Ordinal).OrderBy(x => x, StringComparer.Ordinal).ToList().AsReadOnly();
        EntryCount = entryCount; DocumentCount = documentCount;
        RequireSingleDocument = options.RequireSingleDocument; RequireNoEncryption = options.RequireNoEncryption;
        BudgetDescription = $"XML bytes={options.MaxXmlBytes}; single XML={options.MaxSingleXmlBytes}; depth={options.MaxXmlDepth}; nodes={options.MaxXmlNodes}; references={options.MaxReferences}; results={options.MaxResults}; input={options.LoadOptions.MaxInputBytes}; entries={options.LoadOptions.MaxEntryCount}; entry bytes={options.LoadOptions.MaxEntryUncompressedBytes}; expanded bytes={options.LoadOptions.MaxTotalUncompressedBytes}; ratio={options.LoadOptions.MaxCompressionRatio}";
    }
    /// <summary>Version of this local policy, not an OFD or archival standard version.</summary>
    public string ProfileVersion => "ofdrw-local-preflight/0.1-draft";
    /// <summary>Execution integrity, independent of rule outcomes.</summary>
    public PreflightRunState RunState { get; }
    /// <summary>Finite inspected scope and package lookup semantics.</summary>
    public string Scope => "Known OFD control declarations and package payload presence; case-insensitive package lookup; no rendering, decryption, digest or trust verification.";
    /// <summary>Certification is never assessed by this API.</summary>
    public string Certification => "NotAssessed";
    /// <summary>Whether the single-document policy was requested.</summary>
    public bool RequireSingleDocument { get; }
    /// <summary>Whether absence of known encryption markers was requested.</summary>
    public bool RequireNoEncryption { get; }
    /// <summary>Snapshot of budgets used for this run.</summary>
    public string BudgetDescription { get; }
    /// <summary>Loaded package entries; zero if loading did not finish.</summary>
    public int EntryCount { get; }
    /// <summary>Recognized DocBody count; zero does not establish applicability.</summary>
    public int DocumentCount { get; }
    /// <summary>Stable ordered object observations, including every local rule.</summary>
    public IReadOnlyList<PreflightRuleResult> RuleResults { get; }
    /// <summary>Unsupported semantics and intentionally excluded assessments.</summary>
    public IReadOnlyList<string> CoverageGaps { get; }
    /// <summary>Fail dominates Unknown, then Pass; all inapplicable remains inapplicable.</summary>
    public PreflightOutcome SummaryOutcome => RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail) ? PreflightOutcome.Fail :
        RunState == PreflightRunState.Incomplete || RuleResults.Any(x => x.Outcome == PreflightOutcome.Unknown) ? PreflightOutcome.Unknown :
        RuleResults.Any(x => x.Outcome == PreflightOutcome.Pass) ? PreflightOutcome.Pass : PreflightOutcome.NotApplicable;
    /// <summary>Aggregate a local rule while retaining all object results in the report.</summary>
    public PreflightOutcome GetRuleOutcome(string ruleId)
    {
        var items = RuleResults.Where(x => x.RuleId == ruleId).ToList();
        if (items.Count == 0) throw new ArgumentException("Unknown rule identifier.", nameof(ruleId));
        return items.Any(x => x.Outcome == PreflightOutcome.Fail) ? PreflightOutcome.Fail :
            items.Any(x => x.Outcome == PreflightOutcome.Unknown) ? PreflightOutcome.Unknown :
            items.Any(x => x.Outcome == PreflightOutcome.Pass) ? PreflightOutcome.Pass : PreflightOutcome.NotApplicable;
    }
}
