using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Packaging.Tests;

public class ArchivalCopyTests
{
    private const string Ns = "http://www.ofdspec.org";
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
