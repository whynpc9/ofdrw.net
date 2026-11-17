using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Tests;

// A1d: the roadmap §7 A1d table, row by row, on real library output. The DOCX sample is converted with the pinned
// Noto Sans CJK SC from tests/fixtures/trusted-fonts (no host font); the test is skipped visibly when that LFS fixture
// is not materialized. Assertions use reference relationships, never incidental resource IDs.
public sealed class PipelineRegressionTests : IClassFixture<PipelineRegressionTests.Outputs>
{
    private static readonly XNamespace Ofd = "http://www.ofdspec.org";
    private const string B1 = "6.2.6(b)(1)";
    private readonly Outputs outputs;

    public PipelineRegressionTests(Outputs outputs) { this.outputs = outputs; }

    // Row 1: simple Layout + fixed font + complete references.
    [Fact]
    public async Task SimpleLayoutWithFixedFontIsDecidableButKeepsOverallGaps()
    {
        var document = new FlowDocument();
        document.Options.DefaultFontFamily = "Noto Sans";
        document.Blocks.Add(new Paragraph("Archival assessment of a simple Layout document."));
        var package = document.Render();
        Assert.NotEmpty(package.Fonts);
        foreach (var font in package.Fonts) font.Data = FontPayloadTests.NotoSans();
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);

        var report = await Assess(output.ToArray());
        Assert.Equal((ArchivalExecutionState.Complete, ArchivalInputSupport.Supported), (report.ExecutionState, report.InputSupport.Support));
        foreach (var clause in new[] { B1, "6.2.6(c)(1)", "6.2.6(c)(2)" })
            Assert.Equal(PreflightOutcome.Pass, report.GetClause(clause).Outcome);
        Assert.Equal(PreflightOutcome.NotApplicable, report.GetClause("6.2.6(b)(2)").Outcome);
        AssertFontsMatchReferences(Unzip(output.ToArray()), report);
        Assert.DoesNotContain("FontPayloadNotAssessed", report.Document.Gaps);
        Assert.Contains("NoFullStandardAssessment", report.Document.Gaps);
        Assert.Contains("ReferencedEditionsNotFixed", report.Document.Gaps);
        Assert.NotEmpty(report.Document.UncheckedClauses);
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome); // ordinary DocType stays a declaration Fail
    }

    // Row 2: Native DOCX, explicit native and default mode.
    [PinnedFontTheory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task NativeDocxKeepsTextAndProvenanceAndKnownTagsAreDecided(bool explicitNative)
    {
        var package = outputs.Get(explicitNative);
        var entries = Unzip(package);
        var text = string.Concat(entries.Where(x => x.Key.EndsWith("Content.xml", StringComparison.Ordinal))
            .SelectMany(x => Xml(x.Value).Descendants(Ofd + "TextCode")).Select(x => x.Value));
        foreach (var phrase in new[] { "文档转换视觉回归样例", "Generated DOCX / PDF / OFD Fixture", "待复核", "这段文字应为红色粗体。" })
            Assert.Contains(phrase, text);
        var tags = Xml(entries["Doc_0/Tags/CustomTag_KeyValue.xml"]).Root!.Elements().ToDictionary(x => (string)x.Attribute("Key")!, x => (string)x.Attribute("Value")!);
        Assert.Equal("DOCX/OpenXML", tags["source-text-origin"]);

        var report = await Assess(package);
        Assert.Equal((ArchivalExecutionState.Complete, ArchivalInputSupport.Supported), (report.ExecutionState, report.InputSupport.Support));
        foreach (var clause in new[] { "6.11", "6.1（第一段）" })
            Assert.Contains(report.GetClause(clause).Observations, x => x.RuleId == "AC-016" && x.ReasonCode == "KnownCustomTagPayload" && x.Outcome == PreflightOutcome.Pass);
        Assert.DoesNotContain("CustomTagClosureNotAssessed", report.Document.Gaps);
        // The archival entry evaluates the known tags; APF-002..006 are no longer blanket Unknown there.
        var local = report.Coverage!.FoundationalRules.LocalPreflight;
        Assert.Equal(ArchivalPreflightReport.KnownCustomTagsProfileVersion, local.ProfileVersion);
        Assert.DoesNotContain(local.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope");
        // The legacy APF profile keeps its own contract on the same bytes: both coexist.
        var legacy = await Legacy(package);
        Assert.Equal(ArchivalPreflightReport.DefaultProfileVersion, legacy.ProfileVersion);
        foreach (var rule in new[] { "APF-002", "APF-003", "APF-005", "APF-006" })
            Assert.Equal(PreflightOutcome.Unknown, legacy.GetRuleOutcome(rule));
        Assert.Contains(legacy.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope");

        foreach (var clause in new[] { B1, "6.2.6(c)(1)", "6.2.6(c)(2)" })
            Assert.Equal(PreflightOutcome.Pass, report.GetClause(clause).Outcome);
        AssertFontsMatchReferences(entries, report);
        Assert.DoesNotContain("FontPayloadNotAssessed", report.Document.Gaps);
        Assert.Equal("NonArchivalDocType", Assert.Single(report.GetClause("6.2.1(a)").Observations).ReasonCode);
        Assert.Equal("NoClippingText", Assert.Single(report.GetClause("6.2.6(b)(2)").Observations).ReasonCode);
        // Clauses needing source or external evidence remain Unknown after A1; that is the correct result.
        foreach (var (clause, reason) in new[] { ("6.3.3(b)", "SourceActionProvenanceUnavailable"), ("6.9(a)", "ExternalActionDependenciesUnverified"),
                     ("6.12(a)", "ExtensionScopeUnverified"), ("6.16", "PlaintextUnverified") })
            Assert.Equal((PreflightOutcome.Unknown, reason), (report.GetClause(clause).Outcome, Assert.Single(report.GetClause(clause).Observations).ReasonCode));
        // The excerpt in docs/content/guides/archival-assessment.md is pinned here.
        Assert.Equal(new[] { "NoFullStandardAssessment", "ReferencedEditionsNotFixed", "MatrixScopeUnverified:6.11,6.13.1(a),6.13.1(b),6.13.1(c)",
            "LocalPreflightRulesNotMappedToClauses" }, report.Document.Gaps);
        Assert.Equal((1, 10, 8, 86), (Count(PreflightOutcome.Fail), Count(PreflightOutcome.Pass), Count(PreflightOutcome.NotApplicable), Count(PreflightOutcome.Unknown)));
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome);
        int Count(PreflightOutcome outcome) => report.Clauses.Count(x => x.Outcome == outcome);
    }

    // Row 3: missing FontFile / missing payload, built from the same output.
    [PinnedFontTheory]
    [InlineData("no-fontfile")]
    [InlineData("no-payload")]
    public async Task MissingFontPayloadIsLocatedUnderTheNoExemptionPolicy(string variant)
    {
        var entries = Unzip(outputs.Get(true));
        var used = UsedFontIds(entries).First();
        var resEntry = entries.Keys.Single(x => x.EndsWith("PublicRes.xml", StringComparison.Ordinal));
        var res = Xml(entries[resEntry]);
        var font = res.Descendants(Ofd + "Font").Single(x => (string?)x.Attribute("ID") == used);
        var file = font.Element(Ofd + "FontFile")!;
        string? removedPath = null;
        if (variant == "no-fontfile") { file.Remove(); entries[resEntry] = Bytes(res); }
        else
        {
            removedPath = ResolveFontFile(resEntry, res, file.Value);
            Assert.True(entries.Remove(removedPath));
        }
        var negative = Zip(entries);

        var report = await Assess(negative, new ArchivalAssessmentOptions { FontExemptionPolicy = ArchivalFontExemptionPolicy.NoExemptionsAccepted });
        var failure = Assert.Single(report.GetClause(B1).Observations, x => x.Evidence.Observed.StartsWith("Font ID=" + used + ";", StringComparison.Ordinal));
        Assert.Equal(("AC-017", PreflightOutcome.Fail, variant == "no-fontfile" ? "FontPayloadNotEmbedded" : "FontFileMissing", resEntry),
            (failure.RuleId, failure.Outcome, failure.ReasonCode, failure.Evidence.Entry));
        if (removedPath != null) Assert.Contains("FontFile=" + removedPath, failure.Evidence.Observed);
        Assert.Equal(PreflightOutcome.Fail, report.GetClause(B1).Outcome);
        Assert.Equal(PreflightOutcome.Fail, report.Document.StandardOutcome);
        // Under the default policy a name-only font is an unverified exemption, not a Fail.
        if (variant == "no-fontfile")
        {
            var lenient = await Assess(negative);
            Assert.Equal("FontExemptionBasisUnverified",
                Assert.Single(lenient.GetClause(B1).Observations, x => x.Outcome != PreflightOutcome.Pass).ReasonCode);
            Assert.Equal(PreflightOutcome.Unknown, lenient.GetClause(B1).Outcome);
        }
        // The legacy APF report keeps its own contract: still the CustomTags scope Unknown, unaffected by font checks.
        var legacy = await Legacy(negative);
        Assert.Equal(ArchivalPreflightReport.DefaultProfileVersion, legacy.ProfileVersion);
        Assert.Contains(legacy.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope");
    }

    // Row 4: unknown tags or foreign structure.
    [PinnedFontTheory]
    [InlineData("unknown-tag")]
    [InlineData("foreign-structure")]
    public async Task UnknownTagsOrForeignStructureStayUnknown(string variant)
    {
        var entries = Unzip(outputs.Get(true));
        if (variant == "unknown-tag")
        {
            var index = Xml(entries["Doc_0/Tags/CustomTags.xml"]);
            index.Root!.Add(new XElement(Ofd + "CustomTag", new XAttribute("TypeID", "vendor"), new XAttribute("NameSpace", "urn:vendor:tags"),
                new XElement(Ofd + "FileLoc", "Vendor.xml")));
            entries["Doc_0/Tags/CustomTags.xml"] = Bytes(index);
            entries["Doc_0/Tags/Vendor.xml"] = Encoding.UTF8.GetBytes("<Vendor xmlns='urn:vendor:tags'/>");
        }
        else
        {
            var document = Xml(entries["Doc_0/Document.xml"]);
            document.Root!.Add(new XElement(XName.Get("Vendor", "urn:vendor")));
            entries["Doc_0/Document.xml"] = Bytes(document);
        }
        var report = await Assess(Zip(entries));
        Assert.NotEqual(PreflightOutcome.Pass, report.Document.StandardOutcome);
        if (variant == "unknown-tag")
        {
            Assert.Equal(PreflightOutcome.Unknown, report.GetClause("6.11").Outcome);
            Assert.Contains(report.GetClause("6.11").Observations, x => x.ReasonCode == "UnknownCustomTagSchema");
            Assert.Contains("CustomTagClosureNotAssessed", report.Document.Gaps);
        }
        else
        {
            Assert.Contains(report.InputSupport.Findings, x => x.Code == "UnrecognizedDocumentContent");
            Assert.NotEqual(ArchivalInputSupport.Supported, report.InputSupport.Support);
            Assert.Equal(PreflightOutcome.Unknown, report.GetClause(B1).Outcome);
        }
    }

    // Row 5: only the root relabelled OFD-A.
    [PinnedFontFact]
    public async Task LabelOnlyArchivalDeclarationIsSeparateFromSubstantiveResults()
    {
        var ordinary = outputs.Get(true);
        var entries = Unzip(ordinary);
        var root = Xml(entries["OFD.xml"]);
        root.Root!.SetAttributeValue("DocType", "OFD-A");
        entries["OFD.xml"] = Bytes(root);
        var before = await Assess(ordinary);
        var after = await Assess(Zip(entries));
        Assert.Equal((ArchivalDeclarationStatus.Archival, PreflightOutcome.Pass), (after.Declaration.Status, after.Declaration.Outcome));
        foreach (var clause in before.Clauses.Where(x => x.Clause != "6.2.1(a)"))
            Assert.Equal(clause.Observations.Select(x => (x.RuleId, x.Outcome, x.ReasonCode)),
                after.GetClause(clause.Clause).Observations.Select(x => (x.RuleId, x.Outcome, x.ReasonCode)));
        Assert.Equal(PreflightOutcome.Unknown, after.Document.StandardOutcome);
        Assert.Equal(before.Document.UncheckedClauses, after.Document.UncheckedClauses);
        Assert.Contains("NoFullStandardAssessment", after.Document.Gaps);
    }

    // Row 6: missing DocType, several DocBody, signed, encrypted, attachments, Versions — each reported on its own.
    [PinnedFontTheory]
    [InlineData("missing-doctype", null)]
    [InlineData("multiple", "MultipleDocuments")]
    [InlineData("signatures", "SignaturesDeclared")]
    [InlineData("encryption", "EncryptionMarkerPresent")]
    [InlineData("attachments", "AttachmentsDeclared")]
    [InlineData("versions", "VersionsDeclared")]
    public async Task DeclarationAndStructureOutsideTheProfileAreReportedSeparately(string variant, string? code)
    {
        var entries = Unzip(outputs.Get(true));
        var root = Xml(entries["OFD.xml"]);
        var body = root.Root!.Element(Ofd + "DocBody")!;
        var document = Xml(entries["Doc_0/Document.xml"]);
        switch (variant)
        {
            case "missing-doctype": root.Root.Attribute("DocType")!.Remove(); break;
            case "multiple": root.Root.Add(new XElement(body)); break;
            case "signatures":
                body.Add(new XElement(Ofd + "Signatures", "Doc_0/Signs/Signatures.xml"));
                entries["Doc_0/Signs/Signatures.xml"] = Bytes(new XDocument(new XElement(Ofd + "Signatures",
                    new XElement(Ofd + "Signature", new XAttribute("ID", "1"), new XAttribute("BaseLoc", "Sign_0/Signature.xml")))));
                break;
            case "encryption": entries["Encryptions.xml"] = Bytes(new XDocument(new XElement(Ofd + "Encryptions"))); break;
            case "attachments":
                document.Root!.Element(Ofd + "CustomTags")!.AddBeforeSelf(new XElement(Ofd + "Attachments", "Attachs/Attachments.xml"));
                entries["Doc_0/Attachs/Attachments.xml"] = Bytes(new XDocument(new XElement(Ofd + "Attachments",
                    new XElement(Ofd + "Attachment", new XAttribute("ID", "A1"), new XAttribute("Name", "a"), new XElement(Ofd + "FileLoc", "a.txt")))));
                entries["Doc_0/Attachs/a.txt"] = Encoding.UTF8.GetBytes("attachment");
                break;
            case "versions":
                body.Add(new XElement(Ofd + "Versions", new XElement(Ofd + "Version", new XAttribute("ID", "v1"), new XAttribute("Index", "1"),
                    new XAttribute("BaseLoc", "Doc_0/Version.xml"))));
                break;
        }
        entries["OFD.xml"] = Bytes(root);
        entries["Doc_0/Document.xml"] = Bytes(document);
        var report = await Assess(Zip(entries));
        if (code == null)
        {
            Assert.Equal((ArchivalDeclarationStatus.Missing, "DocTypeMissing"), (report.Declaration.Status, report.Declaration.ReasonCode));
            Assert.Equal(PreflightOutcome.Fail, report.GetClause("6.2.1(a)").Outcome);
            Assert.Equal(ArchivalInputSupport.Supported, report.InputSupport.Support);
            return;
        }
        Assert.Equal(ArchivalDeclarationStatus.NotArchival, report.Declaration.Status);
        Assert.Equal(ArchivalInputSupport.OutOfScope, report.InputSupport.Support);
        var finding = Assert.Single(report.InputSupport.Findings, x => x.Support == ArchivalInputSupport.OutOfScope);
        Assert.Equal(code, finding.Code);
        Assert.Contains("InputOutsideProfile", report.Document.Gaps);
    }

    private static void AssertFontsMatchReferences(Dictionary<string, byte[]> entries, ArchivalAssessmentReport report)
    {
        // Every font ID used by a page has an AC-017 observation located at its declaring table and payload entry.
        var used = UsedFontIds(entries).ToHashSet();
        Assert.NotEmpty(used);
        var observed = report.GetClause(B1).Observations.Select(x => x.Evidence.Observed.Split(';')[0]["Font ID=".Length..]).ToHashSet();
        Assert.Equal(used, observed);
        var resEntry = entries.Keys.Single(x => x.EndsWith("PublicRes.xml", StringComparison.Ordinal));
        var res = Xml(entries[resEntry]);
        foreach (var observation in report.GetClause(B1).Observations)
        {
            var id = observation.Evidence.Observed.Split(';')[0]["Font ID=".Length..];
            var file = res.Descendants(Ofd + "Font").Single(x => (string?)x.Attribute("ID") == id).Element(Ofd + "FontFile")!;
            var target = ResolveFontFile(resEntry, res, file.Value);
            Assert.True(entries.ContainsKey(target));
            Assert.Equal(resEntry, observation.Evidence.Entry);
            Assert.Contains("FontFile=" + target + ";", observation.Evidence.Observed);
        }
    }

    private static IEnumerable<string> UsedFontIds(Dictionary<string, byte[]> entries) =>
        entries.Where(x => x.Key.EndsWith("Content.xml", StringComparison.Ordinal))
            .SelectMany(x => Xml(x.Value).Descendants(Ofd + "TextObject"))
            .Where(x => x.Elements(Ofd + "TextCode").Any(c => c.Value.Trim().Length > 0))
            .Select(x => (string)x.Attribute("Font")!).Distinct(StringComparer.Ordinal);

    // Test-side resolution of the writer's relative FontFile against the table's BaseLoc (both relative here).
    private static string ResolveFontFile(string resEntry, XDocument res, string location)
    {
        var folder = resEntry[..(resEntry.LastIndexOf('/') + 1)];
        var baseLoc = ((string?)res.Root!.Attribute("BaseLoc"))?.Trim('/');
        return folder + (string.IsNullOrEmpty(baseLoc) ? "" : baseLoc + "/") + location.Trim();
    }

    private static async Task<ArchivalAssessmentReport> Assess(byte[] package, ArchivalAssessmentOptions? options = null)
    {
        var hash = SHA256.HashData(package);
        using var input = new MemoryStream(package, false);
        var report = await new ArchivalAssessmentService().AssessAsync(input, options ?? new ArchivalAssessmentOptions());
        Assert.Equal(hash, SHA256.HashData(input.ToArray())); // read-only: nothing is rewritten or cleaned
        Assert.Equal("NotAssessed", report.SoftwareRequirements.Status);
        return report;
    }

    private static async Task<ArchivalPreflightReport> Legacy(byte[] package)
    {
        using var input = new MemoryStream(package, false);
        return await new ArchivalPreflightChecker().CheckAsync(input, new ArchivalPreflightOptions());
    }

    private static XDocument Xml(byte[] bytes) => XDocument.Load(new MemoryStream(bytes, false));

    private static byte[] Bytes(XDocument document)
    {
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
                using (var stream = zip.CreateEntry(entry.Key).Open()) stream.Write(entry.Value);
        return output.ToArray();
    }

    private static Dictionary<string, byte[]> Unzip(byte[] package)
    {
        using var zip = new ZipArchive(new MemoryStream(package, false), ZipArchiveMode.Read);
        var result = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        foreach (var entry in zip.Entries.Where(x => x.Name.Length > 0))
        {
            using var stream = entry.Open();
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            result[entry.FullName] = copy.ToArray();
        }
        return result;
    }

    /// <summary>Converts generated-layout.docx once per mode with the pinned CJK font; nothing from the host.</summary>
    public sealed class Outputs : IDisposable
    {
        private readonly Dictionary<bool, byte[]> cache = new();
        private string? fontDirectory;

        internal byte[] Get(bool explicitNative)
        {
            lock (cache)
            {
                if (cache.TryGetValue(explicitNative, out var bytes)) return bytes;
                fontDirectory ??= PinnedFont.Extract();
                var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
                options.FontDirectories.Add(fontDirectory);
                if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
                using var input = File.OpenRead(Path.Combine(PinnedFont.Root, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx"));
                using var output = new MemoryStream();
                new DocxToOfdConverter(options).ConvertAsync(input, output).GetAwaiter().GetResult();
                return cache[explicitNative] = output.ToArray();
            }
        }

        public void Dispose()
        {
            if (fontDirectory != null) Directory.Delete(fontDirectory, true);
        }
    }
}

internal static class PinnedFont
{
    private const string Fixture = "tests/fixtures/trusted-fonts/sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip";
    private const string Member = "fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf";

    internal static string Root
    {
        get
        {
            var root = new DirectoryInfo(AppContext.BaseDirectory);
            while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
            return root.FullName;
        }
    }

    // A Git LFS pointer is small text; the materialized fixture is a ZIP.
    internal static string? SkipReason
    {
        get
        {
            var path = Path.Combine(Root, Fixture);
            if (!File.Exists(path)) return "Pinned CJK font fixture missing: " + Fixture;
            using var stream = File.OpenRead(path);
            var header = new byte[2];
            return stream.Read(header, 0, 2) == 2 && header[0] == (byte)'P' && header[1] == (byte)'K' ? null :
                "Pinned CJK font fixture is not materialized (run git lfs pull): " + Fixture;
        }
    }

    internal static string Extract()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-archival-cjk-" + Guid.NewGuid().ToString("N"));
        using var archive = ZipFile.OpenRead(Path.Combine(Root, Fixture));
        var entry = archive.GetEntry(Member)!;
        Assert.Equal(21659544, entry.Length);
        var bytes = new byte[entry.Length];
        using (var input = entry.Open()) input.ReadExactly(bytes);
        Assert.Equal("3012A9B63F5ECA3E3B38F23A1BE5ED504675E394ABF8E7A4FA981506582C04AA", Convert.ToHexString(SHA256.HashData(bytes)));
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "noto-cjk.ttf"), bytes);
        return directory;
    }
}

/// <summary>A fact that is reported as skipped, not passed, when the pinned font fixture is unavailable.</summary>
internal sealed class PinnedFontFactAttribute : FactAttribute
{
    public PinnedFontFactAttribute() { if (PinnedFont.SkipReason is { } reason) Skip = reason; }
}

/// <summary>A theory that is reported as skipped, not passed, when the pinned font fixture is unavailable.</summary>
internal sealed class PinnedFontTheoryAttribute : TheoryAttribute
{
    public PinnedFontTheoryAttribute() { if (PinnedFont.SkipReason is { } reason) Skip = reason; }
}
