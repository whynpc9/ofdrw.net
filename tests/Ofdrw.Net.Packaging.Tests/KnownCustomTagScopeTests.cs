using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Packaging.Tests;

// A1b: the opt-in known custom-tag scope of the local preflight. With the option off, results stay exactly as before.
public sealed class KnownCustomTagScopeTests
{
    private const string Ofd = "http://www.ofdspec.org";
    private const string KvNamespace = "urn:ofdrw-net:custom-tags:kv:1";
    private const string EmrNamespace = "urn:ofdrw-net:custom-tags:emr";
    private const string Document = "Doc_0/Document.xml";
    private const string Index = "Doc_0/Tags/CustomTags.xml";
    private const string KvPayload = "Doc_0/Tags/CustomTag_KeyValue.xml";
    private const string EmrPayload = "Doc_0/Tags/CustomTag_EMR.xml";
    private static readonly XName Record = XName.Get("CustomTag", Ofd), FileLoc = XName.Get("FileLoc", Ofd);

    // Captured from the unmodified checker (base f8b3b2b9) for the D2 neutral-channel package; it matches the
    // roadmap §2 baseline: APF-002/003/005/006 UnsupportedResourceScope and APF-004 UnrecognizedControlOrEncryptionScope.
    private const string DefaultBaseline = """
        ofdrw-local-preflight/0.1-draft
        Complete
        Unknown
        R1|APF-001|1|Pass|PageCount|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]|count=1|At least one existing page
        R2|APF-001|1|Pass|DocumentCount|OFD.xml|/{http://www.ofdspec.org}OFD[1]|DocBody count=1|Exactly one recognized DocBody
        R3|APF-001|1|Pass|LocalPayloadPresent|OFD.xml|/{http://www.ofdspec.org}OFD[1]/{http://www.ofdspec.org}DocBody[1]/{http://www.ofdspec.org}DocRoot[1]|Reference=Doc_0/Document.xml; containing=; Target=Doc_0/Document.xml; length=464|Known local declaration and existing package target
        R4|APF-002|1|Unknown|UnsupportedResourceScope|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]||Known local declaration and existing package target
        R5|APF-002|1|Pass|LocalPayloadPresent|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]/{http://www.ofdspec.org}CommonData[1]/{http://www.ofdspec.org}PublicRes[1]|Reference=PublicRes.xml; containing=Doc_0/Document.xml; Target=Doc_0/PublicRes.xml; length=101|Known local declaration and existing package target
        R6|APF-002|1|Pass|LocalPayloadPresent|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]/{http://www.ofdspec.org}Pages[1]/{http://www.ofdspec.org}Page[1]|Reference=Pages/Page_0/Content.xml; containing=Doc_0/Document.xml; Target=Doc_0/Pages/Page_0/Content.xml; length=244; ID=1; Name=|Known local declaration and existing package target
        R7|APF-003|1|Unknown|UnsupportedResourceScope|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]||Known local declaration and existing package target
        R8|APF-004|1|Pass|KnownEncryptionMarkersAbsent|<input>|<package>|Encryptions.xml=False; PasswordCrypto/=False; entries=6|Known local declaration and existing package target
        R9|APF-004|1|Unknown|UnrecognizedControlOrEncryptionScope|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]||Known local declaration and existing package target
        R10|APF-005|1|Unknown|UnsupportedResourceScope|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]||Known local declaration and existing package target
        R11|APF-006|1|Unknown|UnsupportedResourceScope|Doc_0/Document.xml|/{http://www.ofdspec.org}Document[1]||Known local declaration and existing package target
        Font format, glyph coverage, embedding rights and licensing are not assessed.
        Signature digests, cryptography, certificates, timestamps and legal validity are not assessed.
        Standard clauses, OFD conformance and archival certification are not assessed.
        Unknown XML, actions, nested payload contents, full resource-ID closure and XSD validation are not covered.
        UnsupportedResourceScope: Doc_0/Document.xml
        """;

    [Fact]
    public async Task DefaultOffReproducesThePreviousReportExactly()
    {
        var report = await Check(await Neutral(), on: false);
        Assert.Equal(DefaultBaseline.ReplaceLineEndings("\n"), Snapshot(report));
        Assert.Equal(ArchivalPreflightReport.DefaultProfileVersion, report.ProfileVersion);
        // The historical EMR shape also keeps the same blanket results with the option off.
        var emr = await Check(Emr(await Neutral()), on: false);
        Assert.Equal(Outcomes(report), Outcomes(emr));
        Assert.False(new ArchivalPreflightOptions().RecognizeKnownCustomTags);
    }

    [Theory]
    [InlineData("neutral")]
    [InlineData("emr")]
    [InlineData("both")]
    public async Task KnownCompleteChainsAreDecidedAndLocated(string shape)
    {
        var neutral = await Neutral();
        var package = shape == "neutral" ? neutral : shape == "emr" ? Emr(neutral) : Both(neutral);
        var report = await Check(package);
        Assert.Equal(ArchivalPreflightReport.KnownCustomTagsProfileVersion, report.ProfileVersion);
        Assert.Equal(PreflightRunState.Complete, report.RunState);
        Assert.DoesNotContain(report.RuleResults, x => x.Outcome is PreflightOutcome.Unknown or PreflightOutcome.Fail);
        Assert.Contains(report.CoverageGaps, x => x.Contains("GB/T 33190", StringComparison.Ordinal));
        var index = Assert.Single(report.RuleResults, x => x.ReasonCode == "KnownCustomTagIndex");
        Assert.Equal(("APF-002", Document), (index.RuleId, index.Evidence.Entry));
        Assert.Equal("Target=" + Index, index.Evidence.Observed);
        var payloads = report.RuleResults.Where(x => x.ReasonCode == "KnownCustomTagPayload").ToList();
        Assert.All(payloads, x => Assert.Equal(Index, x.Evidence.Entry));
        Assert.All(payloads, x => Assert.EndsWith("{http://www.ofdspec.org}FileLoc[1]", x.Evidence.Selector, StringComparison.Ordinal));
        var expected = shape switch
        {
            "neutral" => new[] { "Shape=KeyValue; Target=" + KvPayload + "; tags=2" },
            "emr" => new[] { "Shape=LegacyEmr; Target=" + EmrPayload + "; tags=2" },
            _ => new[] { "Shape=KeyValue; Target=" + KvPayload + "; tags=2", "Shape=LegacyEmr; Target=" + EmrPayload + "; tags=1" }
        };
        Assert.Equal(expected.Order(), payloads.Select(x => x.Evidence.Observed).Order());
    }

    [Theory]
    [InlineData("missing", "MissingTarget")]
    [InlineData("../../../outside.xml", "RootEscape")]
    [InlineData("https://example.invalid/tags.xml", "ExternalReference")]
    [InlineData("", "EmptyReference")]
    public async Task KnownPayloadEdgeOutsideThePackageFailsAtTheRecord(string location, string reason)
    {
        var entries = Entries(await Neutral());
        if (location == "missing") entries.Remove(KvPayload);
        else entries[Index] = Edit(entries[Index], root => root.Element(Record)!.Element(FileLoc)!.Value = location);
        var report = await Check(Zip(entries));
        var failure = Assert.Single(report.RuleResults, x => x.Outcome == PreflightOutcome.Fail);
        Assert.Equal(("APF-002", reason, Index), (failure.RuleId, failure.ReasonCode, failure.Evidence.Entry));
        Assert.EndsWith("FileLoc[1]", failure.Evidence.Selector, StringComparison.Ordinal);
        Assert.DoesNotContain(report.RuleResults, x => x.Outcome == PreflightOutcome.Unknown);
    }

    [Fact]
    public async Task MissingIndexFailsAtTheDocumentDeclaration()
    {
        var entries = Entries(await Neutral());
        entries.Remove(Index);
        var report = await Check(Zip(entries));
        var failure = Assert.Single(report.RuleResults, x => x.Outcome == PreflightOutcome.Fail);
        Assert.Equal(("MissingTarget", Document), (failure.ReasonCode, failure.Evidence.Entry));
        Assert.EndsWith("{http://www.ofdspec.org}CustomTags[1]", failure.Evidence.Selector, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("unknown-schema", "UnknownCustomTagSchema")]
    [InlineData("emr-type-with-kv-namespace", "UnknownCustomTagSchema")]
    [InlineData("foreign-record", "UnrecognizedCustomTagRecord")]
    [InlineData("schema-location", "UnrecognizedCustomTagRecord")]
    [InlineData("foreign-tag-in-payload", "UnrecognizedCustomTagPayload")]
    [InlineData("payload-root-mismatch", "UnrecognizedCustomTagPayload")]
    [InlineData("nested-tag-content", "UnrecognizedCustomTagPayload")]
    [InlineData("malformed-payload", "UnreadableCustomTagPayload")]
    [InlineData("duplicate-target", "DuplicateCustomTagTarget")]
    [InlineData("cycle-to-index", "CustomTagRoleConflict")]
    [InlineData("cycle-to-document", "CustomTagRoleConflict")]
    [InlineData("two-locations", "AmbiguousFileLoc")]
    public async Task UnknownOrAmbiguousTagContentStaysUnknown(string variant, string reason)
    {
        var entries = Entries(await Neutral());
        XElement Known(XElement root) => root.Element(Record)!;
        switch (variant)
        {
            case "unknown-schema":
                entries[Index] = Edit(entries[Index], root => Known(root).SetAttributeValue("TypeID", "vendor")); break;
            case "emr-type-with-kv-namespace":
                entries[Index] = Edit(entries[Index], root => Known(root).SetAttributeValue("TypeID", "EMR")); break;
            case "foreign-record":
                entries[Index] = Edit(entries[Index], root => Known(root).Name = XName.Get("CustomTag", "urn:foreign")); break;
            case "schema-location":
                entries[Index] = Edit(entries[Index], root => Known(root).Add(new XElement(XName.Get("SchemaLoc", Ofd), "tags.xsd"))); break;
            case "foreign-tag-in-payload":
                entries[KvPayload] = Edit(entries[KvPayload], root => root.Add(new XElement(XName.Get("Tag", "urn:foreign"),
                    new XAttribute("Key", "k"), new XAttribute("Value", "v")))); break;
            case "payload-root-mismatch":
                // Declared as ofdrw-kv, but the payload has the legacy EMR root and tags.
                entries[KvPayload] = EmrTags(entries[KvPayload], int.MaxValue); break;
            case "nested-tag-content":
                entries[KvPayload] = Edit(entries[KvPayload], root => root.Elements().First().Add(new XElement(XName.Get("Ref", KvNamespace), "Doc_0/x.xml"))); break;
            case "malformed-payload":
                entries[KvPayload] = Encoding.UTF8.GetBytes("<Tags xmlns='" + KvNamespace + "'>"); break;
            case "duplicate-target":
                entries[Index] = Edit(entries[Index], root => root.Add(new XElement(Known(root)))); break;
            case "cycle-to-index":
                entries[Index] = Edit(entries[Index], root => Known(root).Element(FileLoc)!.Value = "CustomTags.xml"); break;
            case "cycle-to-document":
                entries[Index] = Edit(entries[Index], root => Known(root).Element(FileLoc)!.Value = "../Document.xml"); break;
            case "two-locations":
                entries[Index] = Edit(entries[Index], root => Known(root).SetAttributeValue("FileLoc", "CustomTag_KeyValue.xml")); break;
        }
        var package = Zip(entries);
        var report = await Check(package);
        Assert.Equal(PreflightRunState.Complete, report.RunState);
        var unknown = report.RuleResults.Where(x => x.ReasonCode == reason).ToList();
        Assert.Contains(unknown, x => x.RuleId == "APF-002" && x.Outcome == PreflightOutcome.Unknown && x.Evidence.Entry == Index);
        // Unknown tag scope is not decided for the other control rules either; nothing is whitelisted.
        Assert.Contains(report.RuleResults, x => x.RuleId == "APF-003" && x.ReasonCode == reason);
        Assert.DoesNotContain(report.RuleResults, x => x.Outcome == PreflightOutcome.Fail);
        // Off, the same package keeps the blanket result.
        var off = await Check(package, on: false);
        Assert.Contains(off.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope");
        Assert.DoesNotContain(off.RuleResults, x => x.ReasonCode == reason);
    }

    [Fact]
    public async Task MixedKnownAndUnknownRecordsLocateBoth()
    {
        var entries = Entries(Emr(await Neutral()));
        entries[Index] = Edit(entries[Index], root => root.Add(new XElement(Record, new XAttribute("TypeID", "vendor"),
            new XAttribute("NameSpace", "urn:vendor"), new XElement(FileLoc, "Vendor.xml"))));
        entries["Doc_0/Tags/Vendor.xml"] = Encoding.UTF8.GetBytes("<v:Data xmlns:v='urn:vendor'><v:Ref>../Res/x.png</v:Ref></v:Data>");
        var report = await Check(Zip(entries));
        Assert.Single(report.RuleResults, x => x.ReasonCode == "KnownCustomTagPayload" && x.Evidence.Observed.StartsWith("Shape=LegacyEmr", StringComparison.Ordinal));
        var unknown = Assert.Single(report.RuleResults, x => x.RuleId == "APF-002" && x.ReasonCode == "UnknownCustomTagSchema");
        Assert.EndsWith("{http://www.ofdspec.org}CustomTag[2]", unknown.Evidence.Selector, StringComparison.Ordinal);
    }

    [Fact]
    public async Task ForeignSameNameDeclarationInTheDocumentStaysUnsupported()
    {
        var entries = Entries(await Neutral());
        entries[Document] = Edit(entries[Document], root => root.Add(new XElement(XName.Get("CustomTags", "urn:foreign"), "Tags/Other.xml")));
        var report = await Check(Zip(entries));
        Assert.Contains(report.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope" && x.Evidence.Entry == Document);
        Assert.DoesNotContain(report.RuleResults, x => x.ReasonCode == "KnownCustomTagIndex");
    }

    [Fact]
    public async Task OtherScopesStayUnsupportedWhenOptedIn()
    {
        var entries = Entries(await Neutral());
        entries[Document] = Edit(entries[Document], root => root.Add(new XElement(XName.Get("Extensions", Ofd), "Ext.xml")));
        var report = await Check(Zip(entries));
        Assert.Contains(report.RuleResults, x => x.ReasonCode == "UnsupportedResourceScope" && x.Evidence.Entry == Document);
    }

    [Fact]
    public async Task TagTraversalHonoursTheReferenceBudget()
    {
        var package = await Neutral();
        // DocRoot, PublicRes and the page use three references; the two tag edges need two more.
        var off = await Check(package, on: false, references: 4);
        Assert.Equal(PreflightRunState.Complete, off.RunState);
        var on = await Check(package, references: 4);
        Assert.Equal(PreflightRunState.Incomplete, on.RunState);
        Assert.Contains("InspectionBudgetExceeded", on.CoverageGaps);
        var nodes = await Check(package, nodes: 40);
        Assert.Equal(PreflightRunState.Incomplete, nodes.RunState);
    }

    private static IEnumerable<string> Outcomes(ArchivalPreflightReport report) =>
        report.RuleResults.Select(x => x.RuleId + "|" + x.Outcome + "|" + x.ReasonCode + "|" + x.Evidence.Entry + "|" + x.Evidence.Selector);

    internal static string Snapshot(ArchivalPreflightReport report) =>
        string.Join("\n", new[] { report.ProfileVersion, report.RunState.ToString(), report.SummaryOutcome.ToString() }
            .Concat(report.RuleResults.Select(r => string.Join("|", r.ResultId, r.RuleId, r.RuleVersion, r.Outcome, r.ReasonCode,
                r.Evidence.Entry, r.Evidence.Selector, r.Evidence.Observed, r.Evidence.Expected)))
            .Concat(report.CoverageGaps));

    private static async Task<ArchivalPreflightReport> Check(byte[] package, bool on = true, int references = 100_000, int nodes = 200_000)
    {
        using var input = new MemoryStream(package, false);
        var report = await new ArchivalPreflightChecker().CheckAsync(input, new ArchivalPreflightOptions
        { RecognizeKnownCustomTags = on, MaxReferences = references, MaxXmlNodes = nodes });
        Assert.Equal(package, input.ToArray()); // read-only: tag bytes and provenance unchanged
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

    // The pre-D2 writer output: EMR index record and an OFD-namespace EMRTags payload.
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

    // One package carrying both channels, as an edited historical input does.
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
