using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Packaging.Tests;

public class ArchivalPreflightTests
{
    private const string Ns = "http://www.ofdspec.org";
    private static string Xml(string name, string body = "", string attributes = "") => $"<{name} xmlns='{Ns}' {attributes}>{body}</{name}>";
    private static Dictionary<string, byte[]> Seed(bool font = false)
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Bytes(Xml("OFD", "<DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody>")),
            ["Doc_0/Document.xml"] = Bytes(Xml("Document", (font ? "<CommonData><PublicRes>PublicRes.xml</PublicRes></CommonData>" : "") + "<Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages>")),
            ["Doc_0/Pages/P0/Content.xml"] = Bytes(Xml("Page", font ? "<Content><Layer><TextObject Font='1'><TextCode>归档 Archive</TextCode></TextObject></Layer></Content>" : "<Content><Layer><PathObject/></Layer></Content>"))
        };
        if (font)
        {
            entries["Doc_0/PublicRes.xml"] = Bytes(Xml("Res", "<Fonts><Font ID='1'><FontFile>font.bin</FontFile></Font></Fonts>", "BaseLoc='Res'"));
            entries["Doc_0/Res/font.bin"] = new byte[] { 1, 2, 3, 4 }; // Deliberately not a valid font; presence only.
        }
        return entries;
    }
    private static byte[] Bytes(string text) => Encoding.UTF8.GetBytes(text);
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var pair in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                var entry = zip.CreateEntry(pair.Key);
                entry.LastWriteTime = new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var stream = entry.Open(); stream.Write(pair.Value);
            }
        return output.ToArray();
    }
    private static Task<ArchivalPreflightReport> Run(Dictionary<string, byte[]> entries, ArchivalPreflightOptions? options = null) =>
        new ArchivalPreflightChecker().CheckAsync(new MemoryStream(Zip(entries), false), options ?? new());
    private static void Replace(Dictionary<string, byte[]> entries, string entry, string from, string to) =>
        entries[entry] = Bytes(Encoding.UTF8.GetString(entries[entry]).Replace(from, to));
    private static void Attach(Dictionary<string, byte[]> entries, string item = "<Attachment ID='A1'><FileLoc>a.txt</FileLoc></Attachment>")
    {
        Replace(entries, "Doc_0/Document.xml", "</Document>", "<Attachments>Attachs/Attachments.xml</Attachments></Document>");
        entries["Doc_0/Attachs/Attachments.xml"] = Bytes(Xml("Attachments", item));
        entries["Doc_0/Attachs/a.txt"] = Array.Empty<byte>(); // Empty attachments are allowed.
    }
    private static void Sign(Dictionary<string, byte[]> entries)
    {
        Replace(entries, "OFD.xml", "</DocBody>", "<Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody>");
        entries["Doc_0/Signs/Signatures.xml"] = Bytes(Xml("Signatures", "<Signature ID='S1' BaseLoc='Sign_0/Signature.xml'/>"));
        entries["Doc_0/Signs/Sign_0/Signature.xml"] = Bytes(Xml("Signature", "<SignedInfo><References><Reference FileRef='/Doc_0/Document.xml'><CheckValue>invalid</CheckValue></Reference></References></SignedInfo><SignedValue>SignedValue.dat</SignedValue>"));
        entries["Doc_0/Signs/Sign_0/SignedValue.dat"] = new byte[] { 0, 1, 2, 3 }; // Invalid signature, not verified.
    }

    [Theory]
    [InlineData(1, PreflightOutcome.Pass)] [InlineData(1, PreflightOutcome.Fail)]
    [InlineData(1, PreflightOutcome.NotApplicable)] [InlineData(1, PreflightOutcome.Unknown)]
    [InlineData(2, PreflightOutcome.Pass)] [InlineData(2, PreflightOutcome.Fail)]
    [InlineData(2, PreflightOutcome.NotApplicable)] [InlineData(2, PreflightOutcome.Unknown)]
    [InlineData(3, PreflightOutcome.Pass)] [InlineData(3, PreflightOutcome.Fail)]
    [InlineData(3, PreflightOutcome.NotApplicable)] [InlineData(3, PreflightOutcome.Unknown)]
    [InlineData(4, PreflightOutcome.Pass)] [InlineData(4, PreflightOutcome.Fail)]
    [InlineData(4, PreflightOutcome.NotApplicable)] [InlineData(4, PreflightOutcome.Unknown)]
    [InlineData(5, PreflightOutcome.Pass)] [InlineData(5, PreflightOutcome.Fail)]
    [InlineData(5, PreflightOutcome.NotApplicable)] [InlineData(5, PreflightOutcome.Unknown)]
    [InlineData(6, PreflightOutcome.Pass)] [InlineData(6, PreflightOutcome.Fail)]
    [InlineData(6, PreflightOutcome.NotApplicable)] [InlineData(6, PreflightOutcome.Unknown)]
    public async Task RuleMatrix(int rule, PreflightOutcome expected)
    {
        var entries = Seed(rule == 3 || rule == 2 && expected != PreflightOutcome.NotApplicable);
        var options = new ArchivalPreflightOptions();
        switch (rule, expected)
        {
            case (1, PreflightOutcome.Fail): Replace(entries, "OFD.xml", "</OFD>", "<DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"); break;
            case (1, PreflightOutcome.NotApplicable): options.RequireSingleDocument = false; break;
            case (1, PreflightOutcome.Unknown): Replace(entries, "OFD.xml", Ns, "urn:unknown"); break;
            case (2, PreflightOutcome.Fail): Replace(entries, "Doc_0/PublicRes.xml", "font.bin", "https://example.invalid/font.bin"); break;
            case (2, PreflightOutcome.NotApplicable): entries["Doc_0/Document.xml"] = Bytes(Xml("Document", "<Pages/>")); break;
            case (2, PreflightOutcome.Unknown): Replace(entries, "Doc_0/PublicRes.xml", Ns, "urn:unknown"); break;
            case (3, PreflightOutcome.Fail): Replace(entries, "Doc_0/PublicRes.xml", "<FontFile>font.bin</FontFile>", ""); break;
            case (3, PreflightOutcome.NotApplicable): entries = Seed(); break;
            case (3, PreflightOutcome.Unknown): Replace(entries, "Doc_0/Pages/P0/Content.xml", "Font='1'", "Font='99'"); break;
            case (4, PreflightOutcome.Fail): entries["Encryptions.xml"] = Bytes("broken"); break;
            case (4, PreflightOutcome.NotApplicable): options.RequireNoEncryption = false; entries["Encryptions.xml"] = Bytes("broken"); break;
            case (4, PreflightOutcome.Unknown): entries["PasswordCrypto/entriesmap.dat"] = new byte[] { 1 }; break;
            case (5, PreflightOutcome.Pass): Attach(entries); break;
            case (5, PreflightOutcome.Fail): Attach(entries); entries.Remove("Doc_0/Attachs/Attachments.xml"); break;
            case (5, PreflightOutcome.Unknown): entries["Doc_0/Attachs/Attachments.xml"] = Bytes(Xml("Attachments")); break;
            case (6, PreflightOutcome.Pass): Sign(entries); break;
            case (6, PreflightOutcome.Fail): Sign(entries); entries.Remove("Doc_0/Signs/Signatures.xml"); break;
            case (6, PreflightOutcome.Unknown): entries["Doc_0/Signs/Sign_0/SignedValue.dat"] = new byte[] { 1 }; break;
        }
        var report = await Run(entries, options);
        Assert.Equal(expected, report.GetRuleOutcome("APF-00" + rule));
        Assert.Equal(6, report.RuleResults.Select(x => x.RuleId).Distinct().Count());
        Assert.All(report.RuleResults, x =>
        {
            Assert.NotEmpty(x.ResultId); Assert.NotEmpty(x.ReasonCode); Assert.NotEmpty(x.Evidence.Entry);
            Assert.NotEmpty(x.Evidence.Selector); Assert.NotEmpty(x.Evidence.Expected);
            Assert.Equal("LocalPolicy", x.SourceKind); Assert.Equal("Unverified", x.StandardStatus); Assert.Empty(x.Clause);
        });
        Assert.Equal("NotAssessed", report.Certification);
    }

    [Fact]
    public async Task DefiniteFailureRetainsUnknownAndWarningFailureIsStillFailure()
    {
        var entries = Seed(true); Sign(entries);
        entries.Remove("Doc_0/Res/font.bin");
        Replace(entries, "Doc_0/Pages/P0/Content.xml", "Font='1'", "Font='99'");
        entries.Remove("Doc_0/Signs/Sign_0/SignedValue.dat");
        var report = await Run(entries);
        Assert.Equal(PreflightOutcome.Fail, report.SummaryOutcome);
        Assert.Contains(report.RuleResults, x => x.RuleId == "APF-003" && x.Outcome == PreflightOutcome.Unknown);
        Assert.Contains(report.RuleResults, x => x.RuleId == "APF-006" && x.Outcome == PreflightOutcome.Fail && x.Severity == Core.Validation.ValidationSeverity.Warning);
    }

    [Theory]
    [InlineData("../../../escape.bin", PreflightOutcome.Fail)]
    [InlineData("https://user:secret@127.0.0.1/file?token=secret", PreflightOutcome.Fail)]
    [InlineData("file:/tmp/host-sentinel", PreflightOutcome.Fail)]
    [InlineData("ftp://host/file", PreflightOutcome.Fail)]
    [InlineData("custom:payload", PreflightOutcome.Fail)]
    [InlineData("\\\\server\\share\\file", PreflightOutcome.Fail)]
    [InlineData("font.bin?secret=value", PreflightOutcome.Unknown)]
    [InlineData("%66ont.bin", PreflightOutcome.Unknown)]
    [InlineData("font.bin#fragment", PreflightOutcome.Unknown)]
    [InlineData("/Doc_0/Res/font.bin", PreflightOutcome.Pass)]
    public async Task ReferencesAreClassifiedBeforePackageResolution(string location, PreflightOutcome expected)
    {
        var entries = Seed(true); Replace(entries, "Doc_0/PublicRes.xml", "font.bin", location);
        var report = await Run(entries);
        Assert.Equal(expected, report.GetRuleOutcome("APF-003"));
        Assert.DoesNotContain(report.RuleResults, x => x.Evidence.Observed.Contains("secret"));
    }

    [Theory]
    [InlineData("../escape.xml")] [InlineData("/absolute.xml")] [InlineData("ofd.XML")]
    public async Task UnsafeOrDuplicateZipStopsAllRules(string extra)
    {
        var entries = Seed(); entries[extra] = new byte[] { 1 };
        var report = await Run(entries);
        Assert.Equal(PreflightRunState.Incomplete, report.RunState);
        Assert.Equal(PreflightOutcome.Unknown, report.SummaryOutcome);
        Assert.All(report.RuleResults, x => Assert.Equal(PreflightOutcome.Unknown, x.Outcome));
    }

    [Fact]
    public async Task InvalidZipRestoresBorrowedReadOnlyInput()
    {
        using var stream = new BorrowedStream(new byte[] { 1, 2, 3 }, true);
        var report = await new ArchivalPreflightChecker().CheckAsync(stream, new());
        Assert.Equal(PreflightRunState.Incomplete, report.RunState);
        Assert.Equal(PreflightOutcome.Unknown, report.SummaryOutcome);
        Assert.Equal(0, stream.Position); Assert.False(stream.WasDisposed); Assert.Equal(0, stream.WriteCount);
    }

    [Theory]
    [InlineData("bytes")] [InlineData("depth")] [InlineData("nodes")] [InlineData("references")]
    [InlineData("results")] [InlineData("input")] [InlineData("totalXml")]
    public async Task BudgetsNeverBecomePass(string budget)
    {
        var options = new ArchivalPreflightOptions();
        switch (budget)
        {
            case "bytes": options.MaxSingleXmlBytes = 10; break;
            case "totalXml": options.MaxXmlBytes = 150; break;
            case "depth": options.MaxXmlDepth = 1; break;
            case "nodes": options.MaxXmlNodes = 1; break;
            case "references": options.MaxReferences = 1; break;
            case "results": options.MaxResults = 12; break;
            case "input": options.LoadOptions.MaxInputBytes = 1; break;
        }
        var report = await Run(Seed(true), options);
        Assert.Equal(PreflightRunState.Incomplete, report.RunState);
        Assert.Equal(PreflightOutcome.Unknown, report.SummaryOutcome);
        Assert.Equal(6, report.RuleResults.Where(x => x.Outcome == PreflightOutcome.Unknown).Select(x => x.RuleId).Distinct().Count());
        Assert.True(report.RuleResults.Count <= options.MaxResults);
    }

    [Theory]
    [InlineData(true)] [InlineData(false)]
    public async Task ReadOnlyLifetimeBytesAndCancellation(bool seekable)
    {
        var bytes = Zip(Seed(true)); var before = (byte[])bytes.Clone();
        using var stream = new BorrowedStream(bytes, seekable);
        var report = await new ArchivalPreflightChecker().CheckAsync(stream, new());
        Assert.Equal(PreflightOutcome.Pass, report.SummaryOutcome);
        Assert.Equal(before, bytes); Assert.False(stream.WasDisposed); Assert.Equal(0, stream.WriteCount);
        if (seekable) Assert.Equal(0, stream.Position); else Assert.Equal(bytes.Length, stream.Consumed);
        using var cancelled = new BorrowedStream(bytes, seekable);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new ArchivalPreflightChecker().CheckAsync(cancelled, new(), new CancellationToken(true)));
        Assert.False(cancelled.WasDisposed); Assert.Equal(0, cancelled.WriteCount);
        if (seekable) Assert.Equal(0, cancelled.Position);
    }

    [Fact]
    public async Task CancellationDuringLoadAndIoFailurePropagateAndRestorePosition()
    {
        using var cts = new CancellationTokenSource();
        using var cancel = new BorrowedStream(Zip(Seed()), true, () => { cts.Cancel(); throw new OperationCanceledException(cts.Token); });
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => new ArchivalPreflightChecker().CheckAsync(cancel, new(), cts.Token));
        Assert.Equal(0, cancel.Position); Assert.False(cancel.WasDisposed);
        using var io = new BorrowedStream(Zip(Seed()), true, () => throw new IOException("sentinel"));
        var ex = await Assert.ThrowsAsync<IOException>(() => new ArchivalPreflightChecker().CheckAsync(io, new()));
        Assert.Equal("sentinel", ex.Message); Assert.Equal(0, io.Position); Assert.False(io.WasDisposed);
    }

    [Fact]
    public async Task NoFetchUsesObservableNetworkAndExclusiveHostSentinels()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
        var host = Path.GetTempFileName();
        try
        {
            using var exclusiveHost = new FileStream(host, FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            var port = ((IPEndPoint)listener.LocalEndpoint).Port;
            var entries = Seed(true);
            Replace(entries, "Doc_0/PublicRes.xml", "font.bin", $"http://127.0.0.1:{port}/sentinel");
            var network = await Run(entries);
            Assert.Equal(PreflightOutcome.Fail, network.GetRuleOutcome("APF-003")); Assert.False(listener.Pending());
            entries = Seed(true); Replace(entries, "Doc_0/PublicRes.xml", "font.bin", new Uri(host).AbsoluteUri);
            var local = await Run(entries); Assert.Equal(PreflightOutcome.Fail, local.GetRuleOutcome("APF-003"));
            entries = Seed(); entries["OFD.xml"] = Bytes($"<!DOCTYPE OFD [<!ENTITY ext SYSTEM 'http://127.0.0.1:{port}/entity'>]>" + Xml("OFD", "&ext;"));
            var dtd = await Run(entries); Assert.Equal(PreflightOutcome.Fail, dtd.GetRuleOutcome("APF-001")); Assert.False(listener.Pending());
            entries["OFD.xml"] = Bytes($"<!DOCTYPE OFD [<!ENTITY ext SYSTEM '{new Uri(host).AbsoluteUri}'>]>" + Xml("OFD", "&ext;"));
            var fileDtd = await Run(entries); Assert.Equal(PreflightOutcome.Fail, fileDtd.GetRuleOutcome("APF-001"));
            Assert.Equal(0, exclusiveHost.Length);
        }
        finally { listener.Stop(); File.Delete(host); }
    }

    [Theory]
    [InlineData("<Attachment FileLoc='a.txt'><FileLoc>b.txt</FileLoc></Attachment>", PreflightOutcome.Unknown)]
    [InlineData("<Attachment External='true'><FileLoc>a.txt</FileLoc></Attachment>", PreflightOutcome.Fail)]
    [InlineData("<Attachment/>", PreflightOutcome.Fail)]
    [InlineData("<Attachment FileLoc='a.txt'/>", PreflightOutcome.Pass)]
    public async Task AttachmentConflictsAndCompatibility(string item, PreflightOutcome expected)
    {
        var entries = Seed(); Attach(entries, item);
        Assert.Equal(expected, (await Run(entries)).GetRuleOutcome("APF-005"));
    }

    [Fact]
    public async Task DuplicateFontUnknownAndBothDocumentsInspectedWhenPolicyDisabled()
    {
        var entries = Seed(true);
        Replace(entries, "Doc_0/PublicRes.xml", "</Fonts>", "<Font ID='1'><FontFile>font.bin</FontFile></Font></Fonts>");
        var duplicate = await Run(entries); Assert.Equal(PreflightOutcome.Unknown, duplicate.GetRuleOutcome("APF-003"));
        entries = Seed();
        Replace(entries, "OFD.xml", "</OFD>", "<DocBody><DocRoot>Doc_1/Document.xml</DocRoot></DocBody></OFD>");
        entries["Doc_1/Document.xml"] = Bytes(Xml("Document", "<Pages><Page BaseLoc='missing.xml'/></Pages>"));
        var multi = await Run(entries, new() { RequireSingleDocument = false });
        Assert.Equal(2, multi.DocumentCount); Assert.Equal(PreflightOutcome.NotApplicable, multi.GetRuleOutcome("APF-001"));
        Assert.Equal(PreflightOutcome.Fail, multi.GetRuleOutcome("APF-002"));
    }

    [Fact]
    public async Task PageResourcesResourceRootOverrideAndSignatureRootReferences()
    {
        var entries = Seed(true);
        Replace(entries, "Doc_0/PublicRes.xml", "BaseLoc='Res'", "BaseLoc='/Doc_0/Res'");
        Sign(entries);
        Replace(entries, "Doc_0/Signs/Sign_0/Signature.xml", "FileRef='/Doc_0/Document.xml'", "FileRef='Doc_0/Document.xml'");
        Assert.Equal(PreflightOutcome.Pass, (await Run(entries)).SummaryOutcome);
        entries.Remove("Doc_0/PublicRes.xml");
        Replace(entries, "Doc_0/Document.xml", "<CommonData><PublicRes>PublicRes.xml</PublicRes></CommonData>", "");
        entries["Doc_0/Pages/P0/PageRes.xml"] = Bytes(Xml("Res", "<Fonts><Font ID='1'><FontFile>/Doc_0/Res/font.bin</FontFile></Font></Fonts>"));
        Assert.Equal(PreflightOutcome.Pass, (await Run(entries)).GetRuleOutcome("APF-003"));
    }

    [Fact]
    public async Task UnknownNamespacesAndUnmodeledScopesAreNotEmptySuccess()
    {
        var entries = Seed(); Replace(entries, "Doc_0/Document.xml", "</Document>", "<Extensions/><vendor:Attachments xmlns:vendor='urn:vendor'/></Document>");
        var report = await Run(entries);
        foreach (var rule in new[] { 2, 3, 5, 6 }) Assert.Equal(PreflightOutcome.Unknown, report.GetRuleOutcome("APF-00" + rule));
        Assert.Contains(report.CoverageGaps, x => x.Contains("UnsupportedResourceScope"));
    }

    [Fact]
    public async Task DeterministicEvidenceAndOptionsSnapshot()
    {
        var options = new ArchivalPreflightOptions();
        var bytes = Zip(Seed(true));
        using var stream = new BorrowedStream(bytes, false, () => options.RequireSingleDocument = false);
        var report = await new ArchivalPreflightChecker().CheckAsync(stream, options);
        Assert.True(report.RequireSingleDocument);
        var second = await Run(Seed(true));
        Assert.Equal(report.RuleResults.Select(x => x.ResultId + x.RuleId + x.Evidence.Selector), second.RuleResults.Select(x => x.ResultId + x.RuleId + x.Evidence.Selector));
        Assert.Contains(report.RuleResults, x => x.Evidence.Selector.Contains("{" + Ns + "}FontFile[1]") && x.Evidence.Observed.Contains("length=4"));
    }

    [Fact]
    public async Task AllInapplicableAndNoObjectsNeverBecomePass()
    {
        var entries = Seed(); entries["Doc_0/Document.xml"] = Bytes(Xml("Document", "<Pages/>"));
        var report = await Run(entries, new() { RequireSingleDocument = false, RequireNoEncryption = false });
        Assert.Equal(PreflightOutcome.NotApplicable, report.SummaryOutcome);
        Assert.All(report.RuleResults, x => Assert.Equal(PreflightOutcome.NotApplicable, x.Outcome));
        Attach(entries, "");
        Assert.Equal(PreflightOutcome.Unknown, (await Run(entries)).GetRuleOutcome("APF-005"));
    }

    [Theory]
    [InlineData("zeroFont")] [InlineData("missingFont")] [InlineData("badDocument")]
    [InlineData("duplicateRoot")] [InlineData("missingPage")] [InlineData("mediaMissing")]
    [InlineData("transform")] [InlineData("badSignature")] [InlineData("cycle")]
    public async Task MissingConflictingAndUnsupportedControlSemantics(string variation)
    {
        var entries = Seed(true); var rule = 3; var expected = PreflightOutcome.Fail;
        switch (variation)
        {
            case "zeroFont": entries["Doc_0/Res/font.bin"] = Array.Empty<byte>(); break;
            case "missingFont": entries.Remove("Doc_0/Res/font.bin"); break;
            case "badDocument": entries["Doc_0/Document.xml"] = Bytes("<bad"); rule = 1; break;
            case "duplicateRoot": Replace(entries, "OFD.xml", "</DocBody>", "<DocRoot>Doc_0/Document.xml</DocRoot></DocBody>"); rule = 1; break;
            case "missingPage": entries.Remove("Doc_0/Pages/P0/Content.xml"); rule = 1; break;
            case "mediaMissing": Replace(entries, "Doc_0/PublicRes.xml", "</Res>", "<MultiMedias><MultiMedia ID='8'/></MultiMedias></Res>"); rule = 2; break;
            case "transform": Sign(entries); Replace(entries, "Doc_0/Signs/Sign_0/Signature.xml", "<CheckValue>", "<Transform/><CheckValue>"); rule = 6; expected = PreflightOutcome.Unknown; break;
            case "badSignature": Sign(entries); entries["Doc_0/Signs/Sign_0/Signature.xml"] = Bytes("<bad"); rule = 6; expected = PreflightOutcome.Unknown; break;
            case "cycle": Replace(entries, "Doc_0/Document.xml", "PublicRes.xml", "Document.xml"); rule = 2; expected = PreflightOutcome.Unknown; break;
        }
        var report = await Run(entries);
        Assert.Equal(expected, report.GetRuleOutcome("APF-00" + rule));
    }

    [Theory]
    [InlineData("ambiguousRoot")] [InlineData("foreignRoot")] [InlineData("foreignPage")]
    [InlineData("unknownObject")] [InlineData("unknownResource")]
    public async Task UnrecognizedOwnershipCannotBecomeDefiniteAbsenceOrPass(string variation)
    {
        var entries = Seed(); var rule = 1;
        switch (variation)
        {
            case "ambiguousRoot": Replace(entries, "OFD.xml", "Document.xml", "Document.xml?query"); break;
            case "foreignRoot": Replace(entries, "OFD.xml", "<DocRoot>", "<DocRoot xmlns='urn:vendor'>"); break;
            case "foreignPage": Replace(entries, "Doc_0/Document.xml", "<Page ID", "<Page xmlns='urn:vendor' ID"); break;
            case "unknownObject": Replace(entries, "Doc_0/Pages/P0/Content.xml", "PathObject", "VendorText"); rule = 3; break;
            case "unknownResource": entries = Seed(true); Replace(entries, "Doc_0/PublicRes.xml", "<Fonts>", "<Fonts><VendorFont/>"); rule = 3; break;
        }
        Assert.Equal(PreflightOutcome.Unknown, (await Run(entries)).GetRuleOutcome("APF-00" + rule));
    }

    [Fact]
    public async Task SupportedNamespaceVersionsAndResourceBaseWithoutDirectory()
    {
        var entries = Seed(true);
        Replace(entries, "Doc_0/PublicRes.xml", " BaseLoc='Res'", "");
        entries["Doc_0/font.bin"] = entries["Doc_0/Res/font.bin"]; entries.Remove("Doc_0/Res/font.bin");
        foreach (var entry in entries.Keys.Where(x => x.EndsWith(".xml")).ToList()) Replace(entries, entry, Ns, Ns + "/2016");
        Assert.Equal(PreflightOutcome.Pass, (await Run(entries)).SummaryOutcome);
    }

    [Fact]
    public async Task BorrowingPreconditionsRejectBadArguments()
    {
        var checker = new ArchivalPreflightChecker();
        await Assert.ThrowsAsync<ArgumentNullException>(() => checker.CheckAsync(null!, new()));
        using var input = new MemoryStream(Zip(Seed()), false);
        await Assert.ThrowsAsync<ArgumentNullException>(() => checker.CheckAsync(input, null!));
        input.Position = 1;
        await Assert.ThrowsAsync<ArgumentException>(() => checker.CheckAsync(input, new()));
        Assert.Equal(1, input.Position);
        input.Position = 0;
        await Assert.ThrowsAsync<ArgumentOutOfRangeException>(() => checker.CheckAsync(input, new() { MaxResults = 1 }));
        Assert.Equal(0, input.Position);
    }

    [Fact]
    public async Task UnsupportedEncodingIsUnknownAndCannotProveNoEncryption()
    {
        var entries = Seed(); entries["OFD.xml"] = Bytes("<?xml version='1.0' encoding='unknown-codepage'?>" + Xml("OFD"));
        var report = await Run(entries);
        Assert.Equal(PreflightOutcome.Unknown, report.GetRuleOutcome("APF-001"));
        Assert.Equal(PreflightOutcome.Unknown, report.GetRuleOutcome("APF-004"));
        Assert.Contains(report.RuleResults, x => x.ReasonCode == "UnsupportedXmlEncoding");
    }

    [Theory]
    [InlineData("font", 3)] [InlineData("attachment", 5)] [InlineData("signature", 6)]
    [InlineData("signedValue", 6)] [InlineData("media", 2)]
    public async Task ForeignNamespaceRequiredFieldsRemainUnknown(string field, int rule)
    {
        var entries = Seed(true);
        switch (field)
        {
            case "font": Replace(entries, "Doc_0/PublicRes.xml", "<FontFile>", "<FontFile xmlns='urn:unknown'>"); break;
            case "attachment": Attach(entries); Replace(entries, "Doc_0/Attachs/Attachments.xml", "<FileLoc>", "<FileLoc xmlns='urn:unknown'>"); break;
            case "signature": Sign(entries); Replace(entries, "Doc_0/Signs/Signatures.xml", "<Signature ID", "<Signature xmlns='urn:unknown' ID"); break;
            case "signedValue": Sign(entries); Replace(entries, "Doc_0/Signs/Sign_0/Signature.xml", "<SignedValue>", "<SignedValue xmlns='urn:unknown'>"); break;
            case "media": Replace(entries, "Doc_0/PublicRes.xml", "</Res>", "<MultiMedias><MultiMedia><MediaFile xmlns='urn:unknown'>missing.bin</MediaFile></MultiMedia></MultiMedias></Res>"); break;
        }
        Assert.Equal(PreflightOutcome.Unknown, (await Run(entries)).GetRuleOutcome("APF-00" + rule));
    }

    private sealed class BorrowedStream : Stream
    {
        private readonly MemoryStream inner;
        private readonly bool seekable;
        private readonly Action? onRead;
        public BorrowedStream(byte[] bytes, bool seekable, Action? onRead = null) { inner = new MemoryStream(bytes, false); this.seekable = seekable; this.onRead = onRead; }
        public bool WasDisposed { get; private set; }
        public int WriteCount { get; private set; }
        public long Consumed => inner.Position;
        public override bool CanRead => true;
        public override bool CanSeek => seekable;
        public override bool CanWrite => false;
        public override long Length => seekable ? inner.Length : throw new NotSupportedException();
        public override long Position { get => seekable ? inner.Position : throw new NotSupportedException(); set { if (!seekable) throw new NotSupportedException(); inner.Position = value; } }
        public override int Read(byte[] buffer, int offset, int count) { onRead?.Invoke(); return inner.Read(buffer, offset, count); }
        public override long Seek(long offset, SeekOrigin origin) => seekable ? inner.Seek(offset, origin) : throw new NotSupportedException();
        public override void Flush() => throw new NotSupportedException();
        public override void SetLength(long value) { WriteCount++; throw new NotSupportedException(); }
        public override void Write(byte[] buffer, int offset, int count) { WriteCount++; throw new NotSupportedException(); }
        protected override void Dispose(bool disposing) { WasDisposed = true; if (disposing) inner.Dispose(); base.Dispose(disposing); }
    }
}
