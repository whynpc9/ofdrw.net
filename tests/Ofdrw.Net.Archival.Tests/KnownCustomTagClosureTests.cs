using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Tests;

// A1b: AC-016 observes known custom-tag closure against 6.11 (Should, advisory only) and 6.1 first paragraph (Statement).
public sealed class KnownCustomTagClosureTests
{
    private const string Ofd = "http://www.ofdspec.org";
    private const string EmrNamespace = "urn:ofdrw-net:custom-tags:emr";
    private const string Index = "Doc_0/Tags/CustomTags.xml";
    private const string KvPayload = "Doc_0/Tags/CustomTag_KeyValue.xml";
    private const string EmrPayload = "Doc_0/Tags/CustomTag_EMR.xml";
    private const string SelfContained = "6.1（第一段）";
    private static readonly XName Record = XName.Get("CustomTag", Ofd), FileLoc = XName.Get("FileLoc", Ofd);

    [Theory]
    [InlineData("neutral", "Shape=KeyValue")]
    [InlineData("emr", "Shape=LegacyEmr")]
    [InlineData("both", "Shape=LegacyEmr")]
    public async Task KnownCompleteTagsCloseAndRemoveTheGap(string shape, string expected)
    {
        var neutral = await Neutral();
        var report = await Assess(shape == "neutral" ? neutral : shape == "emr" ? Emr(neutral) : Both(neutral));
        Assert.Equal(ArchivalExecutionState.Complete, report.ExecutionState);
        Assert.Equal(ArchivalInputSupport.Supported, report.InputSupport.Support);
        var retention = report.GetClause("6.11");
        Assert.Equal(PreflightOutcome.Pass, retention.Outcome);
        Assert.All(retention.Observations, x => Assert.Equal(("AC-016", "Ofdrw.Net.Archival"), (x.RuleId, x.RuleOwner)));
        Assert.Contains(retention.Observations, x => x.ReasonCode == "KnownCustomTagPayload" && x.Evidence.Observed.StartsWith(expected, StringComparison.Ordinal));
        Assert.Equal(PreflightOutcome.Pass, report.GetClause(SelfContained).Outcome);
        Assert.DoesNotContain("CustomTagClosureNotAssessed", report.Document.Gaps);
        // The embedded local preflight ran with the opt-in tag scope: no blanket UnsupportedResourceScope.
        var local = report.Coverage!.FoundationalRules.LocalPreflight;
        Assert.Equal(ArchivalPreflightReport.KnownCustomTagsProfileVersion, local.ProfileVersion);
        Assert.DoesNotContain(local.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope");
    }

    [Theory]
    [InlineData("missing", "MissingTarget")]
    [InlineData("../../../outside.xml", "RootEscape")]
    [InlineData("file:///etc/tags.xml", "ExternalReference")]
    public async Task UnavailableKnownTargetIsAnAdvisoryOn611AndAFailOn61(string location, string reason)
    {
        var entries = Entries(await Neutral());
        if (location == "missing") entries.Remove(KvPayload);
        else entries[Index] = Edit(entries[Index], root => root.Element(Record)!.Element(FileLoc)!.Value = location);
        var report = await Assess(Zip(entries));
        var retention = report.GetClause("6.11");
        Assert.Equal(ArchivalClauseStrength.Should, retention.Strength);
        Assert.Equal(PreflightOutcome.Unknown, retention.Outcome); // Should: never a clause Fail
        var advisory = Assert.Single(retention.Advisories);
        Assert.Equal((reason, Index), (advisory.ReasonCode, advisory.Evidence.Entry));
        Assert.EndsWith("FileLoc[1]", advisory.Evidence.Selector, StringComparison.Ordinal);
        var selfContained = report.GetClause(SelfContained);
        Assert.Equal(ArchivalClauseStrength.Statement, selfContained.Strength);
        Assert.Equal(PreflightOutcome.Fail, selfContained.Outcome);
        Assert.Contains(SelfContained, report.Clauses.Where(x => x.Outcome == PreflightOutcome.Fail).Select(x => x.Clause));
        Assert.DoesNotContain("CustomTagClosureNotAssessed", report.Document.Gaps); // decided: a located failure
    }

    [Theory]
    [InlineData("unknown-schema", "UnknownCustomTagSchema")]
    [InlineData("foreign-record", "UnrecognizedCustomTagRecord")]
    [InlineData("cycle", "CustomTagRoleConflict")]
    public async Task UnknownTagContentStaysUnknownAndKeepsTheGap(string variant, string reason)
    {
        var entries = Entries(await Neutral());
        entries[Index] = Edit(entries[Index], root =>
        {
            var record = root.Element(Record)!;
            if (variant == "unknown-schema") record.SetAttributeValue("TypeID", "vendor");
            else if (variant == "foreign-record") record.Name = XName.Get("CustomTag", "urn:foreign");
            else record.Element(FileLoc)!.Value = "../Pages/Page_0/Content.xml";
        });
        var report = await Assess(Zip(entries));
        foreach (var clause in new[] { "6.11", SelfContained })
        {
            Assert.Equal(PreflightOutcome.Unknown, report.GetClause(clause).Outcome);
            Assert.Contains(report.GetClause(clause).Observations, x => x.ReasonCode == reason);
        }
        Assert.Empty(report.GetClause("6.11").Advisories);
        Assert.Contains("CustomTagClosureNotAssessed", report.Document.Gaps);
    }

    [Fact]
    public async Task MixedKnownAndUnknownRecordsLocateBoth()
    {
        var entries = Entries(await Neutral());
        entries[Index] = Edit(entries[Index], root => root.Add(new XElement(Record, new XAttribute("TypeID", "vendor"),
            new XAttribute("NameSpace", "urn:vendor"), new XElement(FileLoc, "Vendor.xml"))));
        entries["Doc_0/Tags/Vendor.xml"] = Encoding.UTF8.GetBytes("<v:Data xmlns:v='urn:vendor'/>");
        var observations = (await Assess(Zip(entries))).GetClause("6.11").Observations;
        Assert.Contains(observations, x => x.Outcome == PreflightOutcome.Pass && x.ReasonCode == "KnownCustomTagPayload");
        Assert.Contains(observations, x => x.Outcome == PreflightOutcome.Unknown && x.ReasonCode == "UnknownCustomTagSchema" &&
            x.Evidence.Selector.EndsWith("CustomTag[2]", StringComparison.Ordinal));
    }

    [Fact]
    public async Task TagTraversalBudgetMakesTheRunIncomplete()
    {
        var report = await Assess(await Neutral(), new ArchivalAssessmentOptions { Preflight = { MaxReferences = 4 } });
        Assert.Equal(ArchivalExecutionState.BudgetExhausted, report.ExecutionState);
        Assert.Contains("CustomTagClosureNotAssessed", report.Document.Gaps);
        Assert.Equal(PreflightOutcome.Unknown, report.GetClause("6.11").Outcome);
    }

    private static async Task<ArchivalAssessmentReport> Assess(byte[] package, ArchivalAssessmentOptions? options = null)
    {
        using var input = new MemoryStream(package, false);
        var report = await new ArchivalAssessmentService().AssessAsync(input, options ?? new ArchivalAssessmentOptions());
        Assert.Equal(package, input.ToArray()); // tag bytes and provenance untouched
        Assert.Equal(0, input.Position);
        return report;
    }

    private static async Task<byte[]> Neutral()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.CustomTags["source-text-origin"] = "DOCX/OpenXML";
        package.CustomTags["docx-ofd-mode"] = "Native";
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        return stream.ToArray();
    }

    private static byte[] Emr(byte[] neutral)
    {
        var entries = Entries(neutral);
        entries[EmrPayload] = EmrTags(entries[KvPayload], int.MaxValue);
        entries.Remove(KvPayload);
        entries[Index] = Edit(entries[Index], root =>
        {
            var record = root.Element(Record)!;
            record.SetAttributeValue("TypeID", "EMR"); record.SetAttributeValue("NameSpace", EmrNamespace);
            record.Element(FileLoc)!.Value = "CustomTag_EMR.xml";
        });
        return Zip(entries);
    }

    private static byte[] Both(byte[] neutral)
    {
        var entries = Entries(neutral);
        entries[EmrPayload] = EmrTags(entries[KvPayload], 1);
        entries[Index] = Edit(entries[Index], root => root.AddFirst(new XElement(Record, new XAttribute("TypeID", "EMR"),
            new XAttribute("NameSpace", EmrNamespace), new XElement(FileLoc, "CustomTag_EMR.xml"))));
        return Zip(entries);
    }

    private static byte[] EmrTags(byte[] kv, int count)
    {
        var tags = XDocument.Parse(Encoding.UTF8.GetString(kv)).Root!.Elements().Take(count);
        return Encoding.UTF8.GetBytes(new XDocument(new XElement(XName.Get("EMRTags", Ofd), new XAttribute(XNamespace.Xmlns + "ofd", Ofd),
            tags.Select(tag => new XElement(XName.Get("Tag", Ofd), new XAttribute("Key", tag.Attribute("Key")!.Value),
                new XAttribute("Value", tag.Attribute("Value")!.Value))))).ToString());
    }

    private static byte[] Edit(byte[] xml, Action<XElement> change)
    {
        var document = XDocument.Parse(Encoding.UTF8.GetString(xml));
        change(document.Root!);
        return Encoding.UTF8.GetBytes(document.ToString());
    }

    private static Dictionary<string, byte[]> Entries(byte[] data)
    {
        using var zip = new ZipArchive(new MemoryStream(data), ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry =>
        { using var source = entry.Open(); using var copy = new MemoryStream(); source.CopyTo(copy); return copy.ToArray(); });
    }

    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var entry in entries)
            { using var target = zip.CreateEntry(entry.Key).Open(); target.Write(entry.Value); }
        return stream.ToArray();
    }
}
