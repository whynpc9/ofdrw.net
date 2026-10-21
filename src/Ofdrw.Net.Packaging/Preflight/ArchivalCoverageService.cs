using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>One independently checked, narrow condition in GB/T 42133-2022.</summary>
public sealed class ArchivalCoverageRuleDefinition
{
    internal ArchivalCoverageRuleDefinition(string id, string clause, string condition)
    { RuleId = id; Clause = clause; Condition = condition; }
    public string RuleId { get; }
    public int RuleVersion => 1;
    public string StandardId => "GB/T 42133-2022";
    public string StandardVersion => "2022";
    public string Clause { get; }
    public string Condition { get; }
    public string StandardStatus => "VerifiedSubset";
    public string Source => "https://tsgdag.jhun.edu.cn/_upload/article/files/93/cc/4d8bf26643f49ee8a76ee25d4ef6/82cca584-f57d-47db-9363-a1497f03326a.pdf";
    public string SourceSha256 => "7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82";
}

/// <summary>New R20 conditions; R17's two-rule catalog and copy handler are unchanged.</summary>
public static class ArchivalCoverageRuleCatalog
{
    private static readonly IReadOnlyList<ArchivalCoverageRuleDefinition> RulesValue = Array.AsReadOnly(new[]
    {
        new ArchivalCoverageRuleDefinition("AC-001", "6.2.1(a)", "Recognized OFD root declares DocType=OFD-A"),
        new ArchivalCoverageRuleDefinition("AC-002", "6.2.1(c)", "Recognized OFD root contains exactly one DocBody"),
        new ArchivalCoverageRuleDefinition("AC-003", "6.2.2(c)", "Document actions are intra-document Goto only")
    });
    public static IReadOnlyList<ArchivalCoverageRuleDefinition> Rules => RulesValue;
}

/// <summary>An object-level result, not an archival compliance decision.</summary>
public sealed class ArchivalCoverageRuleResult
{
    internal ArchivalCoverageRuleResult(ArchivalCoverageRuleDefinition rule, PreflightOutcome outcome,
        string reason, string entry, string selector, string observed)
    {
        Rule = rule; Outcome = outcome; ReasonCode = reason;
        ResultId = "AC-RESULT-" + rule.RuleId;
        Evidence = new PreflightEvidence(entry, selector, observed, rule.Condition)
        { ResultId = ResultId, EvidenceId = ResultId + "-E1" };
    }
    public string ResultId { get; }
    public ArchivalCoverageRuleDefinition Rule { get; }
    public PreflightOutcome Outcome { get; }
    public string ReasonCode { get; }
    public PreflightEvidence Evidence { get; }
}

/// <summary>Explicitly incomplete standard coverage plus independently observed subset results.</summary>
public sealed class ArchivalCoverageReport
{
    internal ArchivalCoverageReport(ArchivalRuleReport r17, IReadOnlyList<ArchivalCoverageRuleResult> results)
    { FoundationalRules = r17; RuleResults = results; }
    /// <summary>R17's unchanged two-clause report, including the six local preflight rules.</summary>
    public ArchivalRuleReport FoundationalRules { get; }
    public IReadOnlyList<ArchivalCoverageRuleResult> RuleResults { get; }
    /// <summary>The rest of GB/T 42133-2022 is inventoried separately, not validated by this API.</summary>
    public string CoverageStatus => "Incomplete";
    public string Certification => "NotAssessed";
    public string Scope => "GB/T 42133-2022 6.2.1(a),(c), 6.2.2(c) plus R17 6.2.2(a)-(b); remaining clauses are unassessed by this API.";
    /// <summary>A known failure is reported; otherwise complete-standard status stays Unknown.</summary>
    public PreflightOutcome StandardOutcome => RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail) ||
        FoundationalRules.RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail)
        ? PreflightOutcome.Fail : PreflightOutcome.Unknown;
}

/// <summary>Read-only checks for three additional conditions, preserving R17 default behavior.</summary>
public sealed class ArchivalCoverageService
{
    public async Task<ArchivalCoverageReport> CheckAsync(Stream input, ArchivalPreflightOptions options,
        CancellationToken cancellationToken = default)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (!input.CanRead || !input.CanSeek || input.Position != 0)
            throw new ArgumentException("A seekable input at position zero is required.", nameof(input));
        var snapshot = options.Snapshot();
        try
        {
            var foundational = await new ArchivalRuleService().CheckAsync(input, snapshot, cancellationToken).ConfigureAwait(false);
            if (foundational.LocalPreflight.RunState != PreflightRunState.Complete)
                return Unknown(foundational, "LocalInspectionIncomplete");
            OfdPackageArchive archive;
            try { archive = await new OfdPackageLoader().LoadAsync(input, snapshot.LoadOptions, cancellationToken).ConfigureAwait(false); }
            catch (InvalidDataException) { return Unknown(foundational, "InputRejected"); }
            var root = ArchivalRuleService.Read(archive, "OFD.xml", snapshot)?.Root;
            if (root == null || root.Name.LocalName != "OFD" || !Supported(root.Name.NamespaceName))
                return Unknown(foundational, "RootScopeUnavailable");
            return new ArchivalCoverageReport(foundational, Array.AsReadOnly(new[]
            {
                DocType(root), DocumentCount(root), DocumentActions(archive, foundational, snapshot)
            }));
        }
        finally { input.Position = 0; }
    }

    private static ArchivalCoverageReport Unknown(ArchivalRuleReport foundational, string reason) =>
        new ArchivalCoverageReport(foundational, Array.AsReadOnly(ArchivalCoverageRuleCatalog.Rules.Select(rule =>
            Result(rule, PreflightOutcome.Unknown, reason, "<package>", reason, "<input>")).ToArray()));

    private static ArchivalCoverageRuleResult DocType(XElement root)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[0];
        var selector = Selector(root) + "/@DocType";
        if (root.Attributes().Any(x => x.Name.LocalName == "DocType" && x.Name.Namespace != XNamespace.None))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousDocType", selector, "Foreign DocType attribute present");
        var value = (string?)root.Attribute("DocType");
        if (value == null) return Result(rule, PreflightOutcome.Fail, "DocTypeMissing", selector, "DocType absent");
        return value == "OFD-A"
            ? Result(rule, PreflightOutcome.Pass, "ArchivalDocType", selector, "DocType=OFD-A")
            : Result(rule, PreflightOutcome.Fail, "NonArchivalDocType", selector, "DocType is not OFD-A");
    }

    private static ArchivalCoverageRuleResult DocumentCount(XElement root)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[1];
        var selector = Selector(root) + "/{" + root.Name.NamespaceName + "}DocBody";
        var count = root.Elements(root.Name.Namespace + "DocBody").Count();
        if (count > 1) return Result(rule, PreflightOutcome.Fail, "MultipleDocuments", selector, "Recognized DocBody count=" + count);
        if (root.Elements().Any(x => x.Name.LocalName == "DocBody" && x.Name.Namespace != root.Name.Namespace))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousDocumentScope", selector, "Foreign DocBody present");
        return count == 1
            ? Result(rule, PreflightOutcome.Pass, "SingleDocument", selector, "Recognized DocBody count=1")
            : Result(rule, PreflightOutcome.Fail, "DocumentMissing", selector, "Recognized DocBody count=0");
    }

    private static ArchivalCoverageRuleResult DocumentActions(OfdPackageArchive archive,
        ArchivalRuleReport foundational, ArchivalPreflightOptions options)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[2];
        var entry = foundational.DocumentEntry;
        if (entry == null)
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unresolved", "<input>");
        var document = ArchivalRuleService.Read(archive, entry, options)?.Root;
        if (document == null || document.Name.LocalName != "Document" || !Supported(document.Name.NamespaceName))
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unsupported", entry);
        var ns = document.Name.Namespace;
        var selector = "/{" + document.Name.NamespaceName + "}Document[1]/{" + document.Name.NamespaceName + "}Actions[1]";
        if (document.Elements().Any(x => x.Name.LocalName == "Actions" && x.Name.Namespace != ns))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousActions", selector, "Foreign Actions declaration", entry);
        var containers = document.Elements(ns + "Actions").ToArray();
        if (containers.Length == 0)
            return Result(rule, PreflightOutcome.NotApplicable, "NoDocumentActions", selector, "No Actions declaration", entry);
        if (containers.Length != 1 || containers[0].Attributes().Any(x => !x.IsNamespaceDeclaration) ||
            containers[0].Nodes().Any(x => x is XText text && !string.IsNullOrWhiteSpace(text.Value) ||
                x is not XElement && x is not XText))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousActions", selector, "Actions structure unsupported", entry);
        var actions = containers[0].Elements().ToArray();
        if (actions.Length == 0)
            return Result(rule, PreflightOutcome.NotApplicable, "NoDocumentActions", selector, "No Action objects", entry);
        var unknown = false;
        for (var i = 0; i < actions.Length; i++)
        {
            var action = actions[i];
            var location = selector + "/{" + document.Name.NamespaceName + "}Action[" + (i + 1) + "]";
            if (action.Name != ns + "Action" || action.Attributes().Any(x => !x.IsNamespaceDeclaration &&
                    (x.Name.Namespace != XNamespace.None || x.Name.LocalName != "Event")) ||
                action.Nodes().Any(x => x is XText text && !string.IsNullOrWhiteSpace(text.Value) ||
                    x is not XElement && x is not XText))
            { unknown = true; continue; }
            var regions = action.Elements(ns + "Region").ToArray();
            var kinds = action.Elements().Where(x => x.Name != ns + "Region").ToArray();
            if (kinds.Any(x => x.Name.Namespace == ns && x.Name.LocalName is "URI" or "GotoA" or "Sound" or "Movie"))
                return Result(rule, PreflightOutcome.Fail, "NonGotoDocumentAction", location,
                    "Known non-Goto action present", entry);
            if (regions.Length > 1 || kinds.Length != 1 || kinds[0].Name != ns + "Goto" ||
                kinds[0].Elements().Count(x => x.Name == ns + "Dest" || x.Name == ns + "Bookmark") != 1 ||
                kinds[0].Elements().Any(x => x.Name != ns + "Dest" && x.Name != ns + "Bookmark"))
                unknown = true;
        }
        return unknown
            ? Result(rule, PreflightOutcome.Unknown, "UnsupportedDocumentAction", selector, "Action type uncertain", entry)
            : Result(rule, PreflightOutcome.Pass, "GotoOnly", selector, "All recognized Action objects are Goto", entry);
    }

    private static ArchivalCoverageRuleResult Result(ArchivalCoverageRuleDefinition rule, PreflightOutcome outcome,
        string reason, string selector, string observed, string entry = "OFD.xml") =>
        new ArchivalCoverageRuleResult(rule, outcome, reason, entry, selector, observed);

    private static string Selector(XElement root) => "/{" + root.Name.NamespaceName + "}OFD[1]";
    private static bool Supported(string ns) => ns == "http://www.ofdspec.org" || ns == "http://www.ofdspec.org/2016";
}
