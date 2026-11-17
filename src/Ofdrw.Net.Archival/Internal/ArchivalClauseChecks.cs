using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Internal;

/// <summary>
/// Identity of a clause-mapped check owned by Ofdrw.Net.Archival. The AC- series continues from AC-016 here;
/// AC-001..AC-015 stay owned by Ofdrw.Net.Packaging's ArchivalCoverageRuleCatalog, which this package never changes.
/// </summary>
internal sealed class ArchivalOwnedRuleDefinition
{
    internal ArchivalOwnedRuleDefinition(string ruleId, string clause, string condition, int version = 1,
        params string[] additionalClauses)
    {
        RuleId = ruleId; Clause = clause; Condition = condition; Version = version;
        Clauses = Array.AsReadOnly(new[] { clause }.Concat(additionalClauses).ToArray());
    }
    internal string RuleId { get; }
    /// <summary>Primary clause; used for package-level Unknown observations (budget, load failure).</summary>
    internal string Clause { get; }
    /// <summary>Every ledger clause this rule may observe, primary first.</summary>
    internal IReadOnlyList<string> Clauses { get; }
    internal string Condition { get; }
    internal int Version { get; }
}

/// <summary>Read-only inputs available to a package-owned clause check.</summary>
internal sealed class ArchivalCheckContext
{
    private readonly Dictionary<Type, object> shared = new Dictionary<Type, object>();

    internal ArchivalCheckContext(ArchivalXmlReader xml, ArchivalPreflightOptions options, ArchivalScopeResult scope,
        CancellationToken token, ArchivalFontPolicy? fonts = null)
    {
        Xml = xml; Options = options; Scope = scope; CancellationToken = token;
        Fonts = fonts ?? new ArchivalAssessmentOptions().SnapshotFonts();
    }
    internal ArchivalFontPolicy Fonts { get; }

    /// <summary>One analysis shared by several checks; a failure (such as a budget overrun) is replayed to each of them.</summary>
    internal T Shared<T>(Func<T> create) where T : class
    {
        if (!shared.TryGetValue(typeof(T), out var value))
            shared[typeof(T)] = value = new Lazy<T>(create, LazyThreadSafetyMode.None);
        return ((Lazy<T>)value).Value;
    }
    /// <summary>Bounded XML access; budget overruns throw ArchivalBudgetExceededException.</summary>
    internal ArchivalXmlReader Xml { get; }
    internal OfdPackageArchive Archive => Xml.Archive;
    internal ArchivalPreflightOptions Options { get; }
    /// <summary>Root, document and page entries established by the scope inspection; any may be null/empty.</summary>
    internal ArchivalScopeResult Scope { get; }
    internal CancellationToken CancellationToken { get; }
}

/// <summary>
/// Extension seam for additional clause-mapped checks (not a public plugin API). Planned users:
/// A1b known-CustomTags reference closure (6.11, 6.1) and A1c font payload / used-glyph checks
/// (6.2.6(b)(1), (b)(2), (c)(1), (c)(2)).
/// </summary>
internal interface IArchivalClauseCheck
{
    ArchivalOwnedRuleDefinition Rule { get; }
    /// <summary>Document gap this check can close, or null.</summary>
    string? GapCode { get; }
    /// <summary>Returns at least one observation for <see cref="Rule"/>. Unknown when evidence is insufficient.</summary>
    IEnumerable<ArchivalClauseObservation> Check(ArchivalCheckContext context);
    /// <summary>Whether these observations decided the question behind <see cref="GapCode"/> for this input.</summary>
    bool Decided(IReadOnlyList<ArchivalClauseObservation> observations);
}

internal static class ArchivalOwnedRules
{
    internal const string Owner = "Ofdrw.Net.Archival";
    internal const string PackagingOwner = "Ofdrw.Net.Packaging";
    /// <summary>First AC number owned by this package. Packaging owns AC-001..AC-015.</summary>
    internal const int FirstRuleNumber = 16;

    /// <summary>Registered checks, in report order.</summary>
    internal static IReadOnlyList<IArchivalClauseCheck> Registered { get; } = Array.AsReadOnly(new IArchivalClauseCheck[]
    {
        new KnownCustomTagClosureCheck(),
        new ArchivalFontCheck(ArchivalFontAnalysis.UsedTextRule),
        new ArchivalFontCheck(ArchivalFontAnalysis.ClippingTextRule),
        new ArchivalFontCheck(ArchivalFontAnalysis.OutlineRule),
        new ArchivalFontCheck(ArchivalFontAnalysis.TablesRule)
    });

    internal static IReadOnlyList<IArchivalClauseCheck> Validate(IEnumerable<IArchivalClauseCheck> checks)
    {
        if (checks == null) throw new ArgumentNullException(nameof(checks));
        var list = checks.ToList();
        var ids = new HashSet<string>(StringComparer.Ordinal);
        foreach (var check in list)
        {
            var rule = check?.Rule ?? throw new ArgumentException("A check and its rule are required.", nameof(checks));
            if (!IsOwnedId(rule.RuleId) || !ids.Add(rule.RuleId) ||
                ArchivalCoverageRuleCatalog.Rules.Any(x => x.RuleId == rule.RuleId) ||
                ArchivalRuleCatalog.Rules.Any(x => x.RuleId == rule.RuleId) ||
                rule.Clauses.Any(clause => !ArchivalClauseCatalog.Clauses.Any(x => x.Clause == clause)))
                throw new ArgumentException("Owned checks need unique AC-016+ identities mapped to ledger clauses.", nameof(checks));
        }
        return list.AsReadOnly();
    }

    internal static bool IsOwnedId(string? id) =>
        id != null && id.Length >= 6 && id.StartsWith("AC-", StringComparison.Ordinal) &&
        int.TryParse(id.Substring(3), NumberStyles.None, CultureInfo.InvariantCulture, out var number) &&
        number >= FirstRuleNumber && id.Substring(3).Length >= 3;

    internal static ArchivalClauseObservation Observation(ArchivalOwnedRuleDefinition rule, PreflightOutcome outcome,
        string reason, string entry, string selector, string observed) =>
        Observation(rule, rule.Clause, outcome, reason, entry, selector, observed);

    internal static ArchivalClauseObservation Observation(ArchivalOwnedRuleDefinition rule, string clause, PreflightOutcome outcome,
        string reason, string entry, string selector, string observed) =>
        new ArchivalClauseObservation(rule.RuleId, Owner, rule.Version, clause, outcome, reason,
            new ArchivalEvidenceLocation(entry, selector, observed, rule.Condition));

    internal static ArchivalClauseObservation Observation(ArchivalOwnedRuleDefinition rule, PreflightOutcome outcome,
        string reason, string entry, XElement node, string observed) =>
        Observation(rule, outcome, reason, entry, ArchivalPackagePath.Selector(node), observed);
}

/// <summary>Clause-level aggregation, including the strength policy.</summary>
internal static class ArchivalClauseAggregation
{
    internal static bool IsRecommendation(ArchivalClauseStrength strength) =>
        strength == ArchivalClauseStrength.Should || strength == ArchivalClauseStrength.May;

    /// <summary>Unmet Should/May observations, surfaced instead of a clause Fail.</summary>
    internal static IReadOnlyList<ArchivalClauseObservation> Advisories(ArchivalClauseStrength strength,
        IEnumerable<ArchivalClauseObservation> observations) =>
        Array.AsReadOnly(IsRecommendation(strength)
            ? observations.Where(x => x.Outcome == PreflightOutcome.Fail).ToArray()
            : Array.Empty<ArchivalClauseObservation>());

    /// <summary>
    /// Fail &gt; Unknown &gt; Pass &gt; NotApplicable over mapped observations. No observation means Unknown.
    /// Statement and Shall keep rule Fail; Should and May turn a rule Fail into Unknown plus an advisory.
    /// An incomplete run never reports Pass or NotApplicable.
    /// </summary>
    internal static PreflightOutcome Outcome(ArchivalClauseStrength strength,
        IReadOnlyCollection<ArchivalClauseObservation> observations, bool complete)
    {
        if (observations.Count == 0) return PreflightOutcome.Unknown;
        var effective = observations.Select(x => x.Outcome == PreflightOutcome.Fail && IsRecommendation(strength)
            ? PreflightOutcome.Unknown : x.Outcome).ToList();
        var merged = effective.Contains(PreflightOutcome.Fail) ? PreflightOutcome.Fail :
            effective.Contains(PreflightOutcome.Unknown) ? PreflightOutcome.Unknown :
            effective.Contains(PreflightOutcome.Pass) ? PreflightOutcome.Pass : PreflightOutcome.NotApplicable;
        return !complete && merged != PreflightOutcome.Fail ? PreflightOutcome.Unknown : merged;
    }
}
