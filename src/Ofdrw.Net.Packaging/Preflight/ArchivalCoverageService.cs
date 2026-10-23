using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>One independently checked, narrow condition in GB/T 42133-2022.</summary>
public sealed class ArchivalCoverageRuleDefinition
{
    internal ArchivalCoverageRuleDefinition(string id, string clause, string condition)
    { RuleId = id; Clause = clause; Condition = condition; }
    /// <summary>Stable identifier of this narrow check.</summary>
    public string RuleId { get; }
    /// <summary>Implementation revision of this check.</summary>
    public int RuleVersion => 1;
    /// <summary>Identifier of the mapped standard.</summary>
    public string StandardId => "GB/T 42133-2022";
    /// <summary>Edition year of the mapped standard.</summary>
    public string StandardVersion => "2022";
    /// <summary>Clause containing the checked condition.</summary>
    public string Clause { get; }
    /// <summary>Bounded condition evaluated by this check.</summary>
    public string Condition { get; }
    /// <summary>Mapping status for this subset, not full-standard conformity.</summary>
    public string StandardStatus => "VerifiedSubset";
    /// <summary>Public location of the inspected standard PDF.</summary>
    public string Source => "https://tsgdag.jhun.edu.cn/_upload/article/files/93/cc/4d8bf26643f49ee8a76ee25d4ef6/82cca584-f57d-47db-9363-a1497f03326a.pdf";
    /// <summary>SHA-256 of the inspected standard PDF.</summary>
    public string SourceSha256 => "7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82";
}

/// <summary>New R20 conditions; R17's two-rule catalog and copy handler are unchanged.</summary>
public static class ArchivalCoverageRuleCatalog
{
    private static readonly IReadOnlyList<ArchivalCoverageRuleDefinition> RulesValue = Array.AsReadOnly(new[]
    {
        new ArchivalCoverageRuleDefinition("AC-001", "6.2.1(a)", "Recognized OFD root declares DocType=OFD-A"),
        new ArchivalCoverageRuleDefinition("AC-002", "6.2.1(c)", "Recognized OFD root contains exactly one DocBody"),
        new ArchivalCoverageRuleDefinition("AC-003", "6.2.2(c)", "Document actions are intra-document Goto only"),
        new ArchivalCoverageRuleDefinition("AC-004", "6.2.3(c)", "Page actions are intra-document Goto only"),
        new ArchivalCoverageRuleDefinition("AC-005", "6.2.5(a)", "Outline actions are intra-document Goto only"),
        new ArchivalCoverageRuleDefinition("AC-006", "6.2.5(b)", "Outline Goto targets exist in the selected document"),
        new ArchivalCoverageRuleDefinition("AC-007", "6.5(b)", "Used ImageObject Interpolate attributes are false"),
        new ArchivalCoverageRuleDefinition("AC-008", "6.6(c)", "TextObject size is identified by a valid Size attribute"),
        new ArchivalCoverageRuleDefinition("AC-009", "6.10.1(a)(3)", "Retained annotations declare ReadOnly, NoZoom and NoRotate true"),
        new ArchivalCoverageRuleDefinition("AC-010", "6.3.3(b)", "Graphic objects in pages, templates and annotation appearances retain no Action declarations; migration needs source evidence"),
        new ArchivalCoverageRuleDefinition("AC-011", "6.2.3(d)", "Named page layers are distinguishable within each page"),
        new ArchivalCoverageRuleDefinition("AC-012", "6.10.1(a)(2)", "Retained annotation appearances have no nested PageBlock"),
        new ArchivalCoverageRuleDefinition("AC-013", "6.9(a)", "Recognized Action/URI declarations with explicit external targets or targets resolved against an explicit external Base are not retained; other dependencies remain unverified"),
        new ArchivalCoverageRuleDefinition("AC-014", "6.12(a)", "Recognized Document/Extensions declarations with a direct text target are not retained; nested content and other extension locations remain unverified"),
        new ArchivalCoverageRuleDefinition("AC-015", "6.16", "The package retains no root Encryptions.xml marker; complete plaintext and decryption provenance remain unverified")
    });
    /// <summary>Immutable ordered list of supported coverage checks.</summary>
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
    /// <summary>Identity of this observation within the report.</summary>
    public string ResultId { get; }
    /// <summary>Check identity, standard clause and source.</summary>
    public ArchivalCoverageRuleDefinition Rule { get; }
    /// <summary>Four-state result for the checked condition alone.</summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Stable explanation of the result or uncertainty.</summary>
    public string ReasonCode { get; }
    /// <summary>Package entry and selector with redacted observed metadata.</summary>
    public PreflightEvidence Evidence { get; }
}

/// <summary>Explicitly incomplete standard coverage plus independently observed subset results.</summary>
public sealed class ArchivalCoverageReport
{
    internal ArchivalCoverageReport(ArchivalRuleReport r17, IReadOnlyList<ArchivalCoverageRuleResult> results)
    { FoundationalRules = r17; RuleResults = results; }
    /// <summary>R17's unchanged two-clause report, including the six local preflight rules.</summary>
    public ArchivalRuleReport FoundationalRules { get; }
    /// <summary>One result per R20 coverage check.</summary>
    public IReadOnlyList<ArchivalCoverageRuleResult> RuleResults { get; }
    /// <summary>The rest of GB/T 42133-2022 is inventoried separately, not validated by this API.</summary>
    public string CoverageStatus => "Incomplete";
    /// <summary>Archival certification has not been assessed.</summary>
    public string Certification => "NotAssessed";
    /// <summary>Mapped clause subset and explicit remainder boundary.</summary>
    public string Scope => "GB/T 42133-2022 6.2.1(a),(c), 6.2.2(c), 6.2.3(c)-(d), 6.2.5(a)-(b), 6.3.3(b), 6.5(b), 6.6(c), 6.9(a), 6.10.1(a)(2)-(3), 6.12(a), 6.16 plus R17 6.2.2(a)-(b); 6.9(a), 6.12(a) and 6.16 are violation-only subsets and remaining obligations are unassessed by this API.";
    /// <summary>A known failure is reported; otherwise complete-standard status stays Unknown.</summary>
    public PreflightOutcome StandardOutcome => RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail) ||
        FoundationalRules.RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail)
        ? PreflightOutcome.Fail : PreflightOutcome.Unknown;
}

/// <summary>Read-only standard-mapped checks, preserving R17 default behavior.</summary>
public sealed class ArchivalCoverageService
{
    /// <summary>Checks a borrowed, seekable OFD package at position zero and restores its position.</summary>
    /// <param name="input">Stable package stream, left open.</param>
    /// <param name="options">Local policy and resource budgets, snapshotted before inspection.</param>
    /// <param name="cancellationToken">Cancellation propagates to the package loader.</param>
    /// <returns>R17 observations and additional checks; never a full-standard Pass.</returns>
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
            OfdPackageArchive archive;
            try { archive = await new OfdPackageLoader().LoadAsync(input, snapshot.LoadOptions, cancellationToken).ConfigureAwait(false); }
            catch (InvalidDataException) { return Unknown(foundational, "InputRejected"); }
            var marker = KnownEncryptionMarker(archive);
            if (foundational.LocalPreflight.RunState != PreflightRunState.Complete)
                return Unknown(foundational, "LocalInspectionIncomplete", marker);
            var root = ArchivalRuleService.Read(archive, "OFD.xml", snapshot)?.Root;
            if (root == null || root.Name.LocalName != "OFD" || !Supported(root.Name.NamespaceName))
                return Unknown(foundational, "RootScopeUnavailable", marker);
            var results = new List<ArchivalCoverageRuleResult>
            {
                DocType(root), DocumentCount(root), DocumentActions(archive, root, snapshot)
            };
            results.AddRange(new ArchivalStructuralCoverage(archive, root, snapshot, cancellationToken).Check());
            results.Add(marker);
            return new ArchivalCoverageReport(foundational, results.AsReadOnly());
        }
        finally { input.Position = 0; }
    }

    private static ArchivalCoverageReport Unknown(ArchivalRuleReport foundational, string reason,
        ArchivalCoverageRuleResult? marker = null) =>
        new ArchivalCoverageReport(foundational, Array.AsReadOnly(ArchivalCoverageRuleCatalog.Rules.Select(rule =>
            rule.RuleId == "AC-015" && marker?.Outcome == PreflightOutcome.Fail ? marker :
            Result(rule, PreflightOutcome.Unknown, reason, "<package>", reason, "<input>")).ToArray()));

    private static ArchivalCoverageRuleResult DocType(XElement root)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[0];
        var selector = Selector(root) + "/@DocType";
        var value = (string?)root.Attribute("DocType");
        if (value == null) return Result(rule, PreflightOutcome.Fail, "DocTypeMissing", selector, "DocType absent");
        if (value != "OFD-A")
            return Result(rule, PreflightOutcome.Fail, "NonArchivalDocType", selector, "DocType is not OFD-A");
        if (root.Attributes().Any(x => x.Name.LocalName == "DocType" && x.Name.Namespace != XNamespace.None))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousDocType", selector, "Foreign DocType attribute present");
        return Result(rule, PreflightOutcome.Pass, "ArchivalDocType", selector, "DocType=OFD-A");
    }

    private static ArchivalCoverageRuleResult DocumentCount(XElement root)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[1];
        var selector = Selector(root) + "/{" + root.Name.NamespaceName + "}DocBody";
        var count = root.Elements(root.Name.Namespace + "DocBody").Count();
        if (count > 1) return Result(rule, PreflightOutcome.Fail, "MultipleDocuments", selector, "Recognized DocBody count=" + count);
        if (count == 0) return Result(rule, PreflightOutcome.Fail, "DocumentMissing", selector, "Recognized DocBody count=0");
        if (root.Elements().Any(x => x.Name.LocalName == "DocBody" && x.Name.Namespace != root.Name.Namespace))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousDocumentScope", selector, "Foreign DocBody present");
        return Result(rule, PreflightOutcome.Pass, "SingleDocument", selector, "Recognized DocBody count=1");
    }

    private static ArchivalCoverageRuleResult DocumentActions(OfdPackageArchive archive,
        XElement root, ArchivalPreflightOptions options)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[2];
        var rootNs = root.Name.Namespace;
        var bodies = root.Elements(rootNs + "DocBody").ToArray();
        if (bodies.Length != 1)
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unresolved", "<input>");
        var foreignBody = root.Elements().Any(x => x.Name.LocalName == "DocBody" && x.Name.Namespace != rootNs);
        var roots = bodies[0].Elements(rootNs + "DocRoot").ToArray();
        if (roots.Length != 1 || roots[0].HasElements || roots[0].Attributes().Any(x => !x.IsNamespaceDeclaration))
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unresolved", "<input>");
        var foreignRoot = bodies[0].Elements().Any(x => x.Name.LocalName == "DocRoot" && x.Name.Namespace != rootNs);
        var documentPath = roots[0].Value.Trim();
        if (string.IsNullOrWhiteSpace(documentPath) || documentPath.Contains(":") || documentPath.Contains("?") ||
            documentPath.Contains("#") || documentPath.Contains("%") || documentPath.Contains("\\") ||
            documentPath.StartsWith("//", StringComparison.Ordinal))
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unresolved", "<input>");
        string entry;
        try { entry = OfdPackagePath.Resolve("OFD.xml", documentPath); }
        catch (InvalidDataException)
        {
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unresolved", "<input>");
        }
        var document = ArchivalRuleService.Read(archive, entry, options)?.Root;
        if (document == null || document.Name.LocalName != "Document" || !Supported(document.Name.NamespaceName))
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>",
                "Document root unsupported", entry);
        var ns = document.Name.Namespace;
        var selector = "/{" + document.Name.NamespaceName + "}Document[1]/{" + document.Name.NamespaceName + "}Actions[1]";
        var containers = document.Elements(ns + "Actions").ToArray();
        for (var containerIndex = 0; containerIndex < containers.Length; containerIndex++)
        {
            var recognized = containers[containerIndex].Elements(ns + "Action").ToArray();
            for (var actionIndex = 0; actionIndex < recognized.Length; actionIndex++)
                if (recognized[actionIndex].Elements().Any(x => x.Name.Namespace == ns &&
                        x.Name.LocalName is "URI" or "GotoA" or "Sound" or "Movie"))
                    return Result(rule, PreflightOutcome.Fail, "NonGotoDocumentAction",
                        "/{" + document.Name.NamespaceName + "}Document[1]/{" + document.Name.NamespaceName +
                        "}Actions[" + (containerIndex + 1) + "]/{" + document.Name.NamespaceName +
                        "}Action[" + (actionIndex + 1) + "]", "Known non-Goto action present", entry);
        }
        if (foreignBody || foreignRoot)
            return Result(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", selector,
                "Foreign document declaration present", entry);
        if (document.Elements().Any(x => x.Name.LocalName == "Actions" && x.Name.Namespace != ns))
            return Result(rule, PreflightOutcome.Unknown, "AmbiguousActions", selector, "Foreign Actions declaration", entry);
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

    private static ArchivalCoverageRuleResult KnownEncryptionMarker(OfdPackageArchive archive)
    {
        var rule = ArchivalCoverageRuleCatalog.Rules[14];
        return archive.Contains("Encryptions.xml")
            ? Result(rule, PreflightOutcome.Fail, "EncryptionMarkerRetained", "/Encryptions.xml", "Known encryption marker present", "Encryptions.xml")
            : Result(rule, PreflightOutcome.Unknown, "PlaintextUnverified", "<package>", "No known root encryption marker; complete plaintext unverified", "<input>");
    }

    private static string Selector(XElement root) => "/{" + root.Name.NamespaceName + "}OFD[1]";
    private static bool Supported(string ns) => ns == "http://www.ofdspec.org" || ns == "http://www.ofdspec.org/2016";
}
