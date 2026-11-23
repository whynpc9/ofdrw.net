using System;
using Ofdrw.Net.Packaging.Archive;
using System.Collections.Generic;
using System.Linq;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Internal;

/// <summary>
/// AC-016: known custom tags are retained and every edge of their chain closes inside the package. The traversal is the
/// same one the local preflight runs with <see cref="ArchivalPreflightOptions.RecognizeKnownCustomTags"/>, not a copy.
/// 6.11 is a Should clause with an unverified modal scope, so an unmet observation is an advisory; a known tag edge whose
/// target is missing or outside the package is also observed against 6.1 (first paragraph), a Statement clause.
/// </summary>
internal sealed class KnownCustomTagClosureCheck : IArchivalClauseCheck
{
    internal const string RuleId = "AC-016";
    internal const string NoTagsReason = "NoCustomTagsInResult";
    private const string Retention = "6.11", SelfContained = "6.1（第一段）";

    public ArchivalOwnedRuleDefinition Rule { get; } = new ArchivalOwnedRuleDefinition(RuleId, Retention,
        "Known custom tags (ofdrw-kv; legacy EMR as written by this library) are retained and their index and payload references resolve inside the package",
        1, SelfContained);

    public string? GapCode => "CustomTagClosureNotAssessed";

    public IEnumerable<ArchivalClauseObservation> Check(ArchivalCheckContext context)
    {
        var scope = context.Scope;
        if (scope.Document == null || scope.DocumentEntry == null)
        {
            context.SpendResult();
            yield return ArchivalOwnedRules.Observation(Rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable",
                "<document>", "<package>", "Document not resolved by scope inspection");
            yield break;
        }
        if (!scope.Document.Elements().Any(x => x.Name.LocalName == "CustomTags"))
        {
            // 6.11 depends on tags existing before archiving; the result alone cannot show that.
            context.SpendResult();
            yield return ArchivalOwnedRules.Observation(Rule, PreflightOutcome.Unknown, NoTagsReason, scope.DocumentEntry,
                OfdPackageReference.Selector(scope.Document), "No CustomTags declaration; source tag presence not evidenced");
            yield break;
        }
        var reserved = new HashSet<string>(scope.Pages.Select(x => x.Key), StringComparer.OrdinalIgnoreCase) { "OFD.xml" };
        var findings = KnownCustomTagScope.Traverse(context.Archive, scope.DocumentEntry, scope.Document,
            context.Xml.Read, context.Xml.Visit, reserved.Contains, context.CancellationToken);
        foreach (var finding in findings)
        {
            var selector = finding.Node == null ? "<package>" : OfdPackageReference.Selector(finding.Node);
            var observed = finding.Observed.Length == 0 ? finding.Reason : finding.Observed;
            context.SpendResult();
            yield return ArchivalOwnedRules.Observation(Rule, Retention, finding.Outcome, finding.Reason, finding.Entry, selector, observed);
            context.SpendResult();
            yield return ArchivalOwnedRules.Observation(Rule, SelfContained, finding.Outcome, finding.Reason, finding.Entry, selector, observed);
        }
    }

    // Decided when every observation is located Pass/Fail, or there was nothing to close.
    public bool Decided(IReadOnlyList<ArchivalClauseObservation> observations) =>
        observations.All(x => x.Outcome != PreflightOutcome.Unknown || x.ReasonCode == NoTagsReason);
}
