using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>A narrowly verified archival treatment from GB/T 42133-2022; not a certification rule set.</summary>
public sealed class ArchivalRuleDefinition
{
    internal ArchivalRuleDefinition(string id, string clause, string element)
    { RuleId = id; Clause = clause; Element = element; }
    /// <summary>Stable local rule identifier.</summary>
    public string RuleId { get; }
    /// <summary>Revision of this implementation.</summary>
    public int RuleVersion => 1;
    /// <summary>Source standard and edition.</summary>
    public string StandardId => "GB/T 42133-2022";
    /// <summary>Edition year of the mapped standard.</summary>
    public string StandardVersion => "2022";
    /// <summary>Visually checked paragraph in the named edition.</summary>
    public string Clause { get; }
    /// <summary>Only this narrow rule has a verified mapping; full standard coverage is not claimed.</summary>
    public string StandardStatus => "VerifiedSubset";
    /// <summary>Public source of the inspected standard copy.</summary>
    public string Source => "https://tsgdag.jhun.edu.cn/_upload/article/files/93/cc/4d8bf26643f49ee8a76ee25d4ef6/82cca584-f57d-47db-9363-a1497f03326a.pdf";
    /// <summary>SHA-256 of the inspected PDF bytes.</summary>
    public string SourceSha256 => "7d76c086e1384ff200c0df8f5a40bf67660e841340c5e2630e582c81c22f1b82";
    internal string Element { get; }
}

/// <summary>The two document-root removal rules supported by this profile.</summary>
public static class ArchivalRuleCatalog
{
    private static readonly IReadOnlyList<ArchivalRuleDefinition> RulesValue = Array.AsReadOnly(new[]
    {
        new ArchivalRuleDefinition("AR-001", "6.2.2(a)", "Permissions"),
        new ArchivalRuleDefinition("AR-002", "6.2.2(b)", "VPreferences")
    });
    /// <summary>Immutable ordered rule definitions.</summary>
    public static IReadOnlyList<ArchivalRuleDefinition> Rules => RulesValue;
}

/// <summary>One document-root observation under a specified archival clause.</summary>
public sealed class ArchivalRuleResult
{
    internal ArchivalRuleResult(ArchivalRuleDefinition rule, PreflightOutcome outcome, string reason, PreflightEvidence evidence)
    {
        Rule = rule; Outcome = outcome; ReasonCode = reason; Evidence = evidence;
        ResultId = "AR-RESULT-" + rule.RuleId;
        evidence.ResultId = ResultId; evidence.EvidenceId = ResultId + "-E1";
    }
    /// <summary>Identity within this two-rule report.</summary>
    public string ResultId { get; }
    /// <summary>Verified-subset rule identity and source.</summary>
    public ArchivalRuleDefinition Rule { get; }
    /// <summary>Four-state result for the inspected declaration.</summary>
    public PreflightOutcome Outcome { get; }
    /// <summary>Stable explanation of applicability or uncertainty.</summary>
    public string ReasonCode { get; }
    /// <summary>Package entry and expanded-name selector, without source content.</summary>
    public PreflightEvidence Evidence { get; }
}

/// <summary>Existing local preflight and two opt-in archival observations, kept distinct.</summary>
public sealed class ArchivalRuleReport
{
    internal ArchivalRuleReport(ArchivalPreflightReport local, IReadOnlyList<ArchivalRuleResult> results, string? documentEntry)
    { LocalPreflight = local; RuleResults = results; DocumentEntry = documentEntry; }
    /// <summary>Unchanged six-rule local preflight.</summary>
    public ArchivalPreflightReport LocalPreflight { get; }
    /// <summary>One result per supported archival rule.</summary>
    public IReadOnlyList<ArchivalRuleResult> RuleResults { get; }
    /// <summary>Full archival certification is not evaluated.</summary>
    public string Certification => "NotAssessed";
    /// <summary>Fixed scope of the standard-mapped subset.</summary>
    public string Scope => "GB/T 42133-2022 6.2.2(a)-(b), one recognized Document.xml only; other clauses unassessed.";
    /// <summary>Fail dominates Unknown across local and subset observations; this is still not certification.</summary>
    public PreflightOutcome SummaryOutcome => LocalPreflight.SummaryOutcome == PreflightOutcome.Fail ||
        RuleResults.Any(x => x.Outcome == PreflightOutcome.Fail) ? PreflightOutcome.Fail :
        LocalPreflight.RunState == PreflightRunState.Incomplete || LocalPreflight.SummaryOutcome == PreflightOutcome.Unknown ||
        RuleResults.Any(x => x.Outcome == PreflightOutcome.Unknown) ? PreflightOutcome.Unknown :
        PreflightOutcome.Pass;
    internal string? DocumentEntry { get; }
}

/// <summary>Read-only inspection of two standard-mapped document-root declarations.</summary>
public sealed class ArchivalRuleService
{
    /// <summary>Checks an owned, seekable snapshot at position zero and restores its position.</summary>
    public async Task<ArchivalRuleReport> CheckAsync(Stream input, ArchivalPreflightOptions options,
        CancellationToken cancellationToken = default)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (!input.CanRead || !input.CanSeek || input.Position != 0)
            throw new ArgumentException("A seekable input at position zero is required.", nameof(input));
        var snapshot = options.Snapshot();
        try
        {
            var local = await new ArchivalPreflightChecker().CheckAsync(input, snapshot, cancellationToken).ConfigureAwait(false);
            if (local.RunState != PreflightRunState.Complete || local.GetRuleOutcome("APF-001") != PreflightOutcome.Pass)
                return Unknown(local, "DocumentScopeUnavailable");
            OfdPackageArchive archive;
            try { archive = await new OfdPackageLoader().LoadAsync(input, snapshot.LoadOptions, cancellationToken).ConfigureAwait(false); }
            catch (InvalidDataException) { return Unknown(local, "InputRejected"); }
            var root = Read(archive, "OFD.xml", snapshot)?.Root;
            if (root == null || !Supported(root.Name.NamespaceName)) return Unknown(local, "RootScopeUnavailable");
            var bodies = root.Elements(root.Name.Namespace + "DocBody").ToArray();
            if (bodies.Length != 1) return Unknown(local, "DocumentScopeUnavailable");
            var roots = bodies[0].Elements(root.Name.Namespace + "DocRoot").ToArray();
            if (roots.Length != 1) return Unknown(local, "DocumentScopeUnavailable");
            var documentPath = roots[0].Value.Trim();
            if (UnsafePath(documentPath)) return Unknown(local, "DocumentScopeUnavailable");
            string entry;
            try { entry = OfdPackagePath.Resolve("OFD.xml", documentPath); }
            catch (InvalidDataException) { return Unknown(local, "DocumentScopeUnavailable"); }
            var document = Read(archive, entry, snapshot)?.Root;
            if (document == null || document.Name != root.Name.Namespace + "Document") return Unknown(local, "DocumentScopeUnavailable");
            var results = ArchivalRuleCatalog.Rules.Select(rule => Evaluate(rule, document, entry)).ToArray();
            return new ArchivalRuleReport(local, Array.AsReadOnly(results), entry);
        }
        finally { input.Position = 0; }
    }

    private static ArchivalRuleReport Unknown(ArchivalPreflightReport local, string reason) =>
        new ArchivalRuleReport(local, Array.AsReadOnly(ArchivalRuleCatalog.Rules.Select(rule =>
            Result(rule, PreflightOutcome.Unknown, reason, "<input>", "<package>")).ToArray()), null);

    private static ArchivalRuleResult Evaluate(ArchivalRuleDefinition rule, XElement document, string entry)
    {
        var ns = document.Name.Namespace;
        var matches = document.Elements(ns + rule.Element).ToArray();
        var foreign = document.Elements().Any(x => x.Name.LocalName == rule.Element && x.Name.Namespace != ns);
        if (foreign) return Result(rule, PreflightOutcome.Unknown, "UnsupportedNamespace", entry, Selector(document, rule.Element));
        if (matches.Length == 0) return Result(rule, PreflightOutcome.Pass, "DeclarationAbsent", entry, Selector(document, rule.Element));
        return Result(rule, PreflightOutcome.Fail, matches.Length == 1 ? "DeclarationPresent" : "DuplicateDeclarations",
            entry, Selector(document, rule.Element));
    }

    private static ArchivalRuleResult Result(ArchivalRuleDefinition rule, PreflightOutcome outcome, string reason,
        string entry, string selector) => new ArchivalRuleResult(rule, outcome, reason,
            new PreflightEvidence(entry, selector, reason, "No " + rule.Element + " declaration in recognized Document.xml"));

    private static string Selector(XElement document, string element) =>
        "/{" + document.Name.NamespaceName + "}Document[1]/{" + document.Name.NamespaceName + "}" + element + "[1]";

    private static bool Supported(string ns) => ns == "http://www.ofdspec.org" || ns == "http://www.ofdspec.org/2016";
    private static bool UnsafePath(string path) => string.IsNullOrWhiteSpace(path) || path.Contains(":") ||
        path.Contains("?") || path.Contains("#") || path.Contains("%") || path.Contains("\\") || path.StartsWith("//", StringComparison.Ordinal);

    internal static XDocument? Read(OfdPackageArchive archive, string entry, ArchivalPreflightOptions options)
    {
        if (!archive.TryGetBytes(entry, out var bytes) || bytes.Length > options.MaxSingleXmlBytes) return null;
        var settings = new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
            MaxCharactersInDocument = options.MaxSingleXmlBytes, MaxCharactersFromEntities = 1 };
        try
        {
            using var source = new MemoryStream(bytes, false);
            using var reader = XmlReader.Create(source, settings);
            return XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        }
        catch (XmlException) { return null; }
    }
}
