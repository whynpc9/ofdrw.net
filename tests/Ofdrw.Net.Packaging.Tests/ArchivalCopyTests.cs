using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public class ArchivalCopyTests
{
    private const string Ns = "http://www.ofdspec.org";
    private const string Ns2016 = "http://www.ofdspec.org/2016";
    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);
    private static Dictionary<string, byte[]> Seed(string extra = "") => new(StringComparer.OrdinalIgnoreCase)
    {
        ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
        ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages>{extra}</Document>"),
        ["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer><PathObject/></Layer></Content></Page>")
    };
    private static byte[] Zip(Dictionary<string, byte[]> entries, CompressionLevel level = CompressionLevel.Optimal)
    {
        using var memory = new MemoryStream();
        using (var zip = new ZipArchive(memory, ZipArchiveMode.Create, true))
            foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                using var target = zip.CreateEntry(entry.Key, level).Open();
                target.Write(entry.Value);
            }
        return memory.ToArray();
    }
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    [Theory]
    [InlineData("OFD-A", PreflightOutcome.Pass, "ArchivalDocType")]
    [InlineData("OFD", PreflightOutcome.Fail, "NonArchivalDocType")]
    [InlineData("", PreflightOutcome.Fail, "DocTypeMissing")]
    public async Task CoverageChecksVerifiedArchiveDocTypeWithoutChangingR17(string docType,
        PreflightOutcome expected, string reason)
    {
        var entries = Seed();
        var attribute = docType.Length == 0 ? "" : $" DocType='{docType}'";
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'{attribute}><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        using var input = new MemoryStream(Zip(entries), false);
        var report = await new ArchivalCoverageService().CheckAsync(input, new());
        Assert.Equal(0, input.Position);
        Assert.Equal(expected, report.RuleResults[0].Outcome);
        Assert.Equal(reason, report.RuleResults[0].ReasonCode);
        Assert.Equal(PreflightOutcome.Pass, report.RuleResults[1].Outcome);
        Assert.Equal(PreflightOutcome.NotApplicable, report.RuleResults[2].Outcome);
        Assert.Equal("6.2.1(a)", report.RuleResults[0].Rule.Clause);
        Assert.Equal("6.2.1(c)", report.RuleResults[1].Rule.Clause);
        Assert.Equal("Incomplete", report.CoverageStatus);
        Assert.Equal("NotAssessed", report.Certification);
        Assert.Equal(expected == PreflightOutcome.Fail ? PreflightOutcome.Fail : PreflightOutcome.Unknown, report.StandardOutcome);
        Assert.Equal(2, report.FoundationalRules.RuleResults.Count);
        Assert.Equal("OFD.xml", report.RuleResults[0].Evidence.Entry);
        Assert.NotEmpty(report.RuleResults[0].Evidence.Observed);
    }

    [Theory]
    [InlineData("", PreflightOutcome.NotApplicable, "NoDocumentActions")]
    [InlineData("<Actions><Action><Goto><Bookmark Name='next'/></Goto></Action></Actions>",
        PreflightOutcome.Pass, "GotoOnly")]
    [InlineData("<Actions><Action><Region/><Goto><Bookmark Name='next'/></Goto></Action></Actions>",
        PreflightOutcome.Pass, "GotoOnly")]
    [InlineData("<Actions><Action><URI URI='https://example.org'/></Action></Actions>",
        PreflightOutcome.Fail, "NonGotoDocumentAction")]
    [InlineData("<Actions><Action><FutureAction xmlns='urn:foreign'/></Action></Actions>",
        PreflightOutcome.Unknown, "UnsupportedDocumentAction")]
    public async Task CoverageClassifiesDocumentActionTypeWithoutRewritingIt(string extra,
        PreflightOutcome expected, string reason)
    {
        var entries = Seed(extra);
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        using var input = new MemoryStream(Zip(entries), false);
        var before = Hash(input.ToArray());
        var report = await new ArchivalCoverageService().CheckAsync(input, new());
        Assert.Equal(expected, report.RuleResults[2].Outcome);
        Assert.Equal(reason, report.RuleResults[2].ReasonCode);
        Assert.Equal("6.2.2(c)", report.RuleResults[2].Rule.Clause);
        Assert.Equal("Doc_0/Document.xml", report.RuleResults[2].Evidence.Entry);
        Assert.Equal(expected == PreflightOutcome.Fail ? PreflightOutcome.Fail : PreflightOutcome.Unknown,
            report.StandardOutcome);
        Assert.Equal(before, Hash(input.ToArray()));
        Assert.Equal(0, input.Position);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CoverageFindsKnownDocumentActionWhenLocalPagePolicyFails(bool missingPageEntry)
    {
        var action = "<Actions><Action><URI URI='https://example.org'/></Action></Actions>";
        var entries = Seed(action);
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        if (missingPageEntry) entries.Remove("Doc_0/Pages/P0/Content.xml");
        else entries["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'>{action}</Document>");
        using var input = new MemoryStream(Zip(entries), false);
        var report = await new ArchivalCoverageService().CheckAsync(input, new());
        Assert.Equal(PreflightOutcome.Fail, report.FoundationalRules.LocalPreflight.GetRuleOutcome("APF-001"));
        Assert.All(report.FoundationalRules.RuleResults, x => Assert.Equal(PreflightOutcome.Unknown, x.Outcome));
        Assert.Equal(PreflightOutcome.Fail, report.RuleResults[2].Outcome);
        Assert.Equal("NonGotoDocumentAction", report.RuleResults[2].ReasonCode);
        Assert.Equal("Doc_0/Document.xml", report.RuleResults[2].Evidence.Entry);
        Assert.Equal(PreflightOutcome.Fail, report.StandardOutcome);
    }

    [Fact]
    public async Task CoverageFailsForMultipleDocumentsAndKeepsForeignScopeUnknown()
    {
        var entries = Seed();
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        using (var input = new MemoryStream(Zip(entries), false))
        {
            var report = await new ArchivalCoverageService().CheckAsync(input,
                new ArchivalPreflightOptions { RequireSingleDocument = false });
            Assert.Equal(PreflightOutcome.NotApplicable,
                report.FoundationalRules.LocalPreflight.GetRuleOutcome("APF-001"));
            Assert.Equal(PreflightOutcome.Fail, report.RuleResults[1].Outcome);
            Assert.Equal("MultipleDocuments", report.RuleResults[1].ReasonCode);
            Assert.Equal(PreflightOutcome.Fail, report.StandardOutcome);
        }
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A' xmlns:x='urn:foreign'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody><x:DocBody/></OFD>");
        using (var input = new MemoryStream(Zip(entries), false))
        {
            var report = await new ArchivalCoverageService().CheckAsync(input, new());
            Assert.Equal(PreflightOutcome.Unknown, report.RuleResults[1].Outcome);
            Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
        }
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A' xmlns:x='urn:foreign' x:DocType='OFD'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        using (var input = new MemoryStream(Zip(entries), false))
        {
            var report = await new ArchivalCoverageService().CheckAsync(input, new());
            Assert.Equal(PreflightOutcome.Unknown, report.RuleResults[0].Outcome);
            Assert.Equal("AmbiguousDocType", report.RuleResults[0].ReasonCode);
        }
    }

    [Fact]
    public async Task CoverageCannotTurnRejectedInputIntoStandardPass()
    {
        using var input = new MemoryStream(Utf8("not a ZIP"), false);
        var report = await new ArchivalCoverageService().CheckAsync(input, new());
        Assert.Equal(0, input.Position);
        Assert.Equal(PreflightRunState.Incomplete, report.FoundationalRules.LocalPreflight.RunState);
        Assert.All(report.RuleResults, result => Assert.Equal(PreflightOutcome.Unknown, result.Outcome));
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
    }

    [Fact]
    public async Task CoverageDoesNotInterpretForeignRootAsKnownStandardEvidence()
    {
        var entries = Seed();
        entries["OFD.xml"] = Utf8("<OFD xmlns='urn:foreign' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        using var input = new MemoryStream(Zip(entries), false);
        var report = await new ArchivalCoverageService().CheckAsync(input, new());
        Assert.All(report.RuleResults, result => Assert.Equal(PreflightOutcome.Unknown, result.Outcome));
        Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
    }

    [Fact]
    public async Task R17CopyPreservesR20RootObservationsAndOriginal()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions>");
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        var folder = Path.Combine(Path.GetTempPath(), "r20-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries));
            var inputHash = Hash(File.ReadAllBytes(source));
            var copy = await new ArchivalCopyConverter().ConvertAsync(source, destination, new());
            Assert.True(File.Exists(copy.AuditPath));
            Assert.Equal(inputHash, Hash(File.ReadAllBytes(source)));
            using var output = File.OpenRead(destination);
            var report = await new ArchivalCoverageService().CheckAsync(output, new());
            Assert.Equal(PreflightOutcome.Pass, report.RuleResults[0].Outcome);
            Assert.Equal(PreflightOutcome.Pass, report.RuleResults[1].Outcome);
            Assert.Equal(PreflightOutcome.NotApplicable, report.RuleResults[2].Outcome);
            Assert.Equal(PreflightOutcome.Unknown, report.StandardOutcome);
            Assert.Equal(PreflightOutcome.Pass, report.FoundationalRules.SummaryOutcome);
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task ArchivalCopyKeepsAttachmentOwnershipForSeparatePhysicalRemoval()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions><Attachments>Attachs/Attachments.xml</Attachments>");
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}' DocType='OFD-A'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        entries["Doc_0/Attachs/Attachments.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='A' Name='sample'><FileLoc>a.bin</FileLoc></Attachment></Attachments>");
        entries["Doc_0/Attachs/a.bin"] = [1, 2, 3, 4];
        var folder = Path.Combine(Path.GetTempPath(), "r20-r18-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd");
            var copyPath = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries));
            await new ArchivalCopyConverter().ConvertAsync(source, copyPath, new());
            using var copyFile = File.OpenRead(copyPath);
            var coverage = await new ArchivalCoverageService().CheckAsync(copyFile, new());
            Assert.Equal(PreflightOutcome.Pass, coverage.RuleResults[0].Outcome);
            Assert.Equal(PreflightOutcome.Pass, coverage.RuleResults[1].Outcome);
            Assert.Equal(PreflightOutcome.Unknown, coverage.StandardOutcome);
            copyFile.Position = 0;
            var package = await new OfdReader().ReadAsync(copyFile);
            Assert.Single(package.Attachments);
            Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
            using var removed = new MemoryStream();
            var result = await new OfdPackageWriter().WriteWithResultAsync(package, removed);
            Assert.Contains("Doc_0/Attachs/a.bin", result.RemovedEntries);
            using var output = new ZipArchive(new MemoryStream(removed.ToArray()), ZipArchiveMode.Read);
            Assert.Null(output.GetEntry("Doc_0/Attachs/a.bin"));
            Assert.Null(output.GetEntry("Doc_0/Attachs/Attachments.xml"));
            using var retainedCopy = File.OpenRead(copyPath);
            var retained = await new OfdPackageLoader().LoadAsync(retainedCopy);
            Assert.Equal(entries["Doc_0/Attachs/a.bin"], retained.GetBytes("Doc_0/Attachs/a.bin"));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData("", PreflightOutcome.Pass, PreflightOutcome.Pass)]
    [InlineData("<Permissions><Edit>false</Edit></Permissions>", PreflightOutcome.Fail, PreflightOutcome.Pass)]
    [InlineData("<VPreferences><HideToolbar>true</HideToolbar></VPreferences>", PreflightOutcome.Pass, PreflightOutcome.Fail)]
    [InlineData("<other:Permissions xmlns:other='urn:unknown'/>", PreflightOutcome.Unknown, PreflightOutcome.Unknown)]
    public async Task StandardMappedSubsetReportsFourStateAndKeepsLegacySix(string extra, PreflightOutcome permission, PreflightOutcome view)
    {
        using var input = new MemoryStream(Zip(Seed(extra)), false);
        var report = await new ArchivalRuleService().CheckAsync(input, new());
        Assert.Equal(0, input.Position);
        Assert.Equal(6, report.LocalPreflight.RuleResults.Select(x => x.RuleId).Distinct().Count());
        Assert.Equal(permission, report.RuleResults[0].Outcome);
        Assert.Equal(view, report.RuleResults[1].Outcome);
        Assert.Equal("GB/T 42133-2022", report.RuleResults[0].Rule.StandardId);
        Assert.Equal("2022", report.RuleResults[0].Rule.StandardVersion);
        Assert.Equal("6.2.2(a)", report.RuleResults[0].Rule.Clause);
        Assert.Equal("6.2.2(b)", report.RuleResults[1].Rule.Clause);
        Assert.All(report.RuleResults, x => { Assert.NotEmpty(x.ResultId); Assert.NotEmpty(x.Evidence.EvidenceId); });
        Assert.Equal("NotAssessed", report.Certification);
    }

    [Fact]
    public async Task LocalPackageFailureCannotProduceSubsetSummaryPass()
    {
        var entries = Seed(); entries.Remove("Doc_0/Pages/P0/Content.xml");
        using var input = new MemoryStream(Zip(entries), false);
        var report = await new ArchivalRuleService().CheckAsync(input, new());
        Assert.Equal(PreflightOutcome.Fail, report.LocalPreflight.SummaryOutcome);
        Assert.Equal(PreflightOutcome.Fail, report.SummaryOutcome);
    }

    [Fact]
    public async Task RuleInspectionKeepsInitialBudgetsWhenCallerMutatesOptionsDuringRead()
    {
        var options = new ArchivalPreflightOptions();
        using var input = new OnReadStream(Zip(Seed()), () => options.MaxSingleXmlBytes = 1);
        var report = await new ArchivalRuleService().CheckAsync(input, options);
        Assert.Equal(1, options.MaxSingleXmlBytes);
        Assert.Equal(PreflightOutcome.Pass, report.LocalPreflight.SummaryOutcome);
        Assert.All(report.RuleResults, result => Assert.Equal(PreflightOutcome.Pass, result.Outcome));
    }

    [Fact]
    public async Task CopyKeepsStoredRepetitiveAttachmentWithinCompressionRatioBudget()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions><Attachments>Attachs/Attachments.xml</Attachments>");
        entries["Doc_0/Attachs/Attachments.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='A'><FileLoc>a.txt</FileLoc></Attachment></Attachments>");
        entries["Doc_0/Attachs/a.txt"] = Utf8(new string('x', 32_768));
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries, CompressionLevel.NoCompression));
            var options = new ArchivalPreflightOptions(); options.LoadOptions.MaxCompressionRatio = 12;
            var result = await new ArchivalCopyConverter().ConvertAsync(source, destination, options);
            Assert.Equal(PreflightOutcome.Pass, result.Postflight.SummaryOutcome);
            using var file = File.OpenRead(destination);
            using var zip = new ZipArchive(file, ZipArchiveMode.Read);
            var attachment = zip.GetEntry("Doc_0/Attachs/a.txt")!;
            Assert.True(attachment.Length / (double)attachment.CompressedLength <= 12);
            using var output = File.OpenRead(destination);
            var package = await new OfdPackageLoader().LoadAsync(output, options.LoadOptions);
            Assert.Equal(entries["Doc_0/Attachs/a.txt"], package.GetBytes("Doc_0/Attachs/a.txt"));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task FormattedDocRootUsesSameReferenceAsLocalPreflight()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions>");
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>\n  Doc_0/Document.xml\n</DocRoot></DocBody></OFD>");
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries));
            using (var input = File.OpenRead(source))
            {
                var before = await new ArchivalRuleService().CheckAsync(input, new());
                Assert.Equal(PreflightOutcome.Fail, before.RuleResults[0].Outcome);
                Assert.Equal("DeclarationPresent", before.RuleResults[0].ReasonCode);
            }
            var result = await new ArchivalCopyConverter().ConvertAsync(source, destination, new());
            Assert.Equal(PreflightOutcome.Pass, result.Postflight.SummaryOutcome);
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task DisabledSingleDocumentPolicyStillInspectsOneUnambiguousDocument()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions>");
        var options = new ArchivalPreflightOptions { RequireSingleDocument = false };
        using (var input = new MemoryStream(Zip(entries), false))
        {
            var report = await new ArchivalRuleService().CheckAsync(input, options);
            Assert.Equal(PreflightOutcome.NotApplicable, report.LocalPreflight.GetRuleOutcome("APF-001"));
            Assert.Equal(PreflightOutcome.Fail, report.RuleResults[0].Outcome);
            Assert.Equal("DeclarationPresent", report.RuleResults[0].ReasonCode);
        }
        var ambiguous = Seed("<Permissions><Edit>false</Edit></Permissions>");
        ambiguous["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        using (var input = new MemoryStream(Zip(ambiguous), false))
        {
            var report = await new ArchivalRuleService().CheckAsync(input, options);
            Assert.All(report.RuleResults, result => Assert.Equal(PreflightOutcome.Unknown, result.Outcome));
        }
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries));
            var result = await new ArchivalCopyConverter().ConvertAsync(source, destination, options);
            Assert.Equal(PreflightOutcome.Pass, result.Postflight.SummaryOutcome);
        }
        finally { Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData(Ns, Ns2016)]
    [InlineData(Ns2016, Ns)]
    public async Task DistinctSupportedPackageAndDocumentNamespacesRemainInspectable(string packageNs, string documentNs)
    {
        var entries = Seed();
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{packageNs}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        entries["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{documentNs}'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages><Permissions><Edit>false</Edit></Permissions></Document>");
        entries["Doc_0/Pages/P0/Content.xml"] = Utf8($"<Page xmlns='{documentNs}'><Content><Layer><PathObject/></Layer></Content></Page>");
        using (var input = new MemoryStream(Zip(entries), false))
        {
            var report = await new ArchivalRuleService().CheckAsync(input, new());
            Assert.Equal(PreflightOutcome.Fail, report.RuleResults[0].Outcome);
            Assert.Equal("DeclarationPresent", report.RuleResults[0].ReasonCode);
        }
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries));
            var result = await new ArchivalCopyConverter().ConvertAsync(source, destination, new());
            Assert.Equal(PreflightOutcome.Pass, result.Postflight.SummaryOutcome);
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task UnsupportedDocumentNamespaceRemainsUnknown()
    {
        var entries = Seed();
        entries["Doc_0/Document.xml"] = Utf8("<Document xmlns='urn:unsupported'><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages><Permissions/></Document>");
        using var input = new MemoryStream(Zip(entries), false);
        var report = await new ArchivalRuleService().CheckAsync(input, new());
        Assert.All(report.RuleResults, result => Assert.Equal(PreflightOutcome.Unknown, result.Outcome));
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            File.WriteAllBytes(source, Zip(entries));
            await Assert.ThrowsAsync<InvalidDataException>(() => new ArchivalCopyConverter().ConvertAsync(source, destination, new()));
            Assert.False(File.Exists(destination));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task CaseDistinctDestinationWorksOnCaseSensitiveVolume()
    {
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "archive.ofd"); var destination = Path.Combine(folder, "Archive.ofd");
            File.WriteAllBytes(source, Zip(Seed("<Permissions><Edit>false</Edit></Permissions>")));
            if (File.Exists(destination)) return; // This volume treats the two names as one file.
            var inputHash = Hash(File.ReadAllBytes(source));
            var result = await new ArchivalCopyConverter().ConvertAsync(source, destination, new());
            Assert.Equal(PreflightOutcome.Pass, result.Postflight.SummaryOutcome);
            Assert.Equal(inputHash, Hash(File.ReadAllBytes(source)));
            Assert.True(File.Exists(destination));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task CopyRemovesOnlyTwoDeclarationsAndAuditsEveryEntry()
    {
        var entries = Seed("<Permissions><Edit>false</Edit><Print Printable='true' Copies='2'/></Permissions>" +
                           "<VPreferences><HideToolbar>true</HideToolbar></VPreferences>" +
                           "<Attachments>Attachs/Attachments.xml</Attachments>");
        entries["Doc_0/Attachs/Attachments.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='A'><FileLoc>a.txt</FileLoc></Attachment></Attachments>");
        entries["Doc_0/Attachs/a.txt"] = Array.Empty<byte>();
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            var bytes = Zip(entries); File.WriteAllBytes(source, bytes);
            var result = await new ArchivalCopyConverter().ConvertAsync(source, destination, new());
            Assert.Equal(Hash(bytes), result.InputSha256, ignoreCase: true);
            Assert.Equal(Hash(bytes), Hash(File.ReadAllBytes(source)));
            Assert.Equal(new[] { "Permissions", "VPreferences" }, result.RemovedDeclarations);
            Assert.False(result.SignaturesInvalidated);
            Assert.Equal(PreflightOutcome.Pass, result.Postflight.SummaryOutcome);
            var audit = XDocument.Load(result.AuditPath);
            Assert.Equal("NotAssessed", (string?)audit.Root?.Attribute("certification"));
            Assert.Equal(entries.Count, audit.Descendants("Entry").Count());
            Assert.Single(audit.Descendants("Entry"), x => (string?)x.Attribute("changed") == "true");
            using var output = File.OpenRead(destination);
            var package = await new OfdPackageLoader().LoadAsync(output);
            Assert.Equal(entries.Keys.OrderBy(x => x), package.EntryNames.OrderBy(x => x));
            foreach (var entry in entries.Keys.Where(x => x != "Doc_0/Document.xml")) Assert.Equal(entries[entry], package.GetBytes(entry));
            var document = XDocument.Parse(package.ReadUtf8Text("Doc_0/Document.xml"));
            Assert.Null(document.Root?.Element(XName.Get("Permissions", Ns)));
            Assert.Null(document.Root?.Element(XName.Get("VPreferences", Ns)));
            Assert.NotNull(document.Root?.Element(XName.Get("Attachments", Ns)));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Theory]
    [InlineData("<Permissions><Edit>false</Edit><u:Note xmlns:u='urn:unknown'/></Permissions>")]
    [InlineData("<VPreferences><HideToolbar>true</HideToolbar><u:Note xmlns:u='urn:unknown'/></VPreferences>")]
    [InlineData("<Permissions/><Permissions/>")]
    public async Task OpaqueOrAmbiguousTargetDoesNotPublish(string extra)
    {
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            var bytes = Zip(Seed(extra)); File.WriteAllBytes(source, bytes);
            await Assert.ThrowsAnyAsync<Exception>(() => new ArchivalCopyConverter().ConvertAsync(source, destination, new()));
            Assert.Equal(Hash(bytes), Hash(File.ReadAllBytes(source)));
            Assert.False(File.Exists(destination)); Assert.False(File.Exists(destination + ".audit.xml"));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task SignedInputIsRejectedWithoutChangingSignatures()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions>");
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature ID='1' BaseLoc='Sign_0/Signature.xml'/></Signatures>");
        entries["Doc_0/Signs/Sign_0/Signature.xml"] = Utf8($"<Signature xmlns='{Ns}'><SignedInfo><References><Reference FileRef='/Doc_0/Document.xml'/></References></SignedInfo><SignedValue>SignedValue.dat</SignedValue></Signature>");
        entries["Doc_0/Signs/Sign_0/SignedValue.dat"] = new byte[] { 1, 2, 3 };
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            var bytes = Zip(entries); File.WriteAllBytes(source, bytes);
            var error = await Assert.ThrowsAsync<InvalidDataException>(() => new ArchivalCopyConverter().ConvertAsync(source, destination, new()));
            Assert.Contains("SignedInputRequiresResigning", error.Message);
            Assert.Equal(Hash(bytes), Hash(File.ReadAllBytes(source)));
            Assert.False(File.Exists(destination));
        }
        finally { Directory.Delete(folder, true); }
    }

    [Fact]
    public async Task UnsupportedUnsignedControlScopeIsNotReportedAsSigned()
    {
        var entries = Seed("<Permissions><Edit>false</Edit></Permissions><Annotations/>");
        var folder = Path.Combine(Path.GetTempPath(), "r17-" + Guid.NewGuid().ToString("N")); Directory.CreateDirectory(folder);
        try
        {
            var source = Path.Combine(folder, "source.ofd"); var destination = Path.Combine(folder, "copy.ofd");
            var bytes = Zip(entries); File.WriteAllBytes(source, bytes);
            using (var input = new MemoryStream(bytes, false))
            {
                var preflight = await new ArchivalPreflightChecker().CheckAsync(input, new());
                Assert.Equal(PreflightOutcome.Unknown, preflight.GetRuleOutcome("APF-006"));
            }
            var error = await Assert.ThrowsAsync<InvalidDataException>(() =>
                new ArchivalCopyConverter().ConvertAsync(source, destination, new()));
            Assert.DoesNotContain("SignedInputRequiresResigning", error.Message);
            Assert.Contains("Preflight", error.Message);
            Assert.Equal(Hash(bytes), Hash(File.ReadAllBytes(source)));
            Assert.False(File.Exists(destination));
        }
        finally { Directory.Delete(folder, true); }
    }

    private sealed class OnReadStream : Stream
    {
        private readonly MemoryStream inner;
        private readonly Action onRead;
        private bool called;
        public OnReadStream(byte[] bytes, Action onRead) { inner = new MemoryStream(bytes, false); this.onRead = onRead; }
        private void Mutate() { if (!called) { called = true; onRead(); } }
        public override bool CanRead => true;
        public override bool CanSeek => true;
        public override bool CanWrite => false;
        public override long Length => inner.Length;
        public override long Position { get => inner.Position; set => inner.Position = value; }
        public override int Read(byte[] buffer, int offset, int count) { Mutate(); return inner.Read(buffer, offset, count); }
        public override int Read(Span<byte> buffer) { Mutate(); return inner.Read(buffer); }
        public override Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        { Mutate(); return inner.ReadAsync(buffer, offset, count, cancellationToken); }
        public override long Seek(long offset, SeekOrigin origin) => inner.Seek(offset, origin);
        public override void Flush() { }
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) { if (disposing) inner.Dispose(); base.Dispose(disposing); }
    }
}
