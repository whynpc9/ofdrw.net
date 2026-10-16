using System.IO.Compression;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging.Preflight;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Pr23AttachmentSplitTests
{
    private static readonly byte[] Original = [0, 1, 2, 255];
    private static readonly byte[] Edited = [9, 8, 7, 0, 255];

    [Theory]
    [InlineData("Doc_0/")]
    [InlineData("")]
    [InlineData("Nested/Doc/")]
    public async Task SplitPersistReadbackReusesAuthenticatedDescriptorAndPayload(string prefix)
    {
        var input = await Seed(prefix); var source = await Read(input);
        source.Attachments[0].Name = "renamed.bin"; source.Attachments[0].Data = Edited;
        var output = await Write(OfdDocumentSplitter.Split(source, [1]));
        var report = await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output), new());
        Save(prefix.Length == 0 ? "root" : prefix.StartsWith("Nested") ? "nested" : "canonical", input, output, report);
        Assert.DoesNotContain(report.RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
        Assert.Equal(2, report.RuleResults.Count(r => r.ReasonCode == "LocalPayloadPresent" && r.Evidence.Entry == prefix + "Attachs/Attachments.xml" && r.Outcome == PreflightOutcome.Pass));
        Assert.Equal("Attachs/Attachments.xml", Document(output, prefix).Root!.Element(Ns(output, prefix) + "Attachments")!.Value);
        Assert.Equal(AttachmentEntries(input, prefix), AttachmentEntries(output, prefix));
        Assert.Equal(Edited, Entries(output)[prefix + "Attachs/Attach_1_original.bin"]);
        Assert.Equal(Original, Entries(input)[prefix + "Attachs/Attach_1_original.bin"]);
        var readback = await Read(output); Assert.Single(readback.Pages);
        Assert.Equal("renamed.bin", readback.Attachments[0].Name); Assert.Equal(Edited, readback.Attachments[0].Data);
        Assert.Equal(source.Attachments.Select(a => a.Id), readback.Attachments.Select(a => a.Id));
        // A second split/read/save must not grow another suffix generation.
        var again = await Write(OfdDocumentSplitter.Split(readback, [0]));
        Assert.Equal(AttachmentEntries(output, prefix), AttachmentEntries(again, prefix));
        Assert.DoesNotContain((await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(again), new())).RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task SplitKeepsExplicitRemovalAndDoesNotReviveSourceSlots(bool all)
    {
        var input = await Seed(); var source = await Read(input);
        if (all) source.Attachments.Clear(); else source.Attachments.RemoveAt(0);
        var split = OfdDocumentSplitter.Split(source, [1]);
        var output = await Write(split); var readback = await Read(output);
        Assert.Equal(all ? 0 : 1, readback.Attachments.Count);
        Assert.Equal(AttachmentEntries(input), AttachmentEntries(output));
        var reference = Document(output).Root!.Element(Ns(output) + "Attachments");
        var report = await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output), new());
        if (all)
        {
            Assert.Null(reference);
            // Clear remains logical deletion: the original orphan is visible to
            // preflight; this change neither erases it nor suppresses the warning.
            Assert.Contains(report.RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
        }
        else
        {
            Assert.Equal("Attachs/Attachments.xml", reference!.Value);
            Assert.Equal("second.bin", Assert.Single(readback.Attachments).Name);
            Assert.Equal(source.Attachments[0].Id, readback.Attachments[0].Id);
            Assert.Equal(new byte[] { 3, 4, 5 }, readback.Attachments[0].Data);
            Assert.DoesNotContain(report.RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
            Assert.Single(XDocument.Parse(Encoding.UTF8.GetString(Entries(output)["Doc_0/Attachs/Attachments.xml"])).Root!.Elements());
        }
        Assert.Equal(Original, Entries(output)["Doc_0/Attachs/Attach_1_original.bin"]);
        Save(all ? "clear" : "remove-one", input, output, report);
    }

    [Fact]
    public async Task SplitRetainsProvenanceThroughControlledCloneAndPreparedWrite()
    {
        var source = await Read(await Seed()); source.Attachments.RemoveAt(0);
        var split = OfdDocumentSplitter.Split(source, [1]);
        Assert.NotSame(source.AttachmentSource, split.AttachmentSource);
        Assert.NotNull(split.AttachmentSource); Assert.All(split.Attachments, a => Assert.Same(split, a.SourceOwner));
        Assert.All(split.Attachments, a => Assert.Contains(source.Attachments, original => !ReferenceEquals(a, original) && a.Id == original.Id && a.Data.SequenceEqual(original.Data)));
        using var context = new OfdControlledEditContext(new());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(split, context.CaptureSource(split), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned); using var output = new MemoryStream(); await prepared.WriteAsync(output);
        Assert.Single((await Read(output.ToArray())).Attachments);
        Assert.DoesNotContain((await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output.ToArray()), new())).RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
    }

    [Theory]
    [InlineData("root")]
    [InlineData("attachment")]
    [InlineData("location")]
    [InlineData("reference")]
    public async Task SplitStillRefusesOpaqueEnvelopesButSameDocumentSaveKeepsThem(string scope)
    {
        var input = await Seed(); var entries = Entries(input);
        var path = scope == "reference" ? "Doc_0/Document.xml" : "Doc_0/Attachs/Attachments.xml";
        var xml = XDocument.Parse(Encoding.UTF8.GetString(entries[path])); var ns = xml.Root!.Name.Namespace;
        var target = scope switch
        {
            "reference" => xml.Root.Element(ns + "Attachments")!,
            "attachment" => xml.Root.Element(ns + "Attachment")!,
            "location" => xml.Root.Element(ns + "Attachment")!.Element(ns + "FileLoc")!,
            _ => xml.Root
        };
        target.SetAttributeValue(XName.Get("opaque", "urn:pr23-vendor"), "keep");
        entries[path] = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)); input = Zip(entries);
        var source = await Read(input); var output = await Write(source);
        var saved = XDocument.Parse(Encoding.UTF8.GetString(Entries(output)[path]));
        Assert.Equal("keep", saved.Root!.DescendantsAndSelf().Attributes(XName.Get("opaque", "urn:pr23-vendor")).Single().Value);
        Assert.Throws<NotSupportedException>(() => OfdDocumentSplitter.Split(source, [1]));
        Assert.Equal(entries[path], source.PreservedEntries[path]);
    }

    [Fact]
    public async Task SplitKeepsNoncanonicalButClosedImportedLocations()
    {
        var entries = Entries(await Seed()); var descriptor = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Attachs/Attachments.xml"]));
        foreach (var location in descriptor.Descendants(descriptor.Root!.Name.Namespace + "FileLoc")) location.Value = "../Attachs/" + location.Value;
        entries.Remove("Doc_0/Attachs/Attachments.xml"); entries["Doc_0/Metadata/List.xml"] = Encoding.UTF8.GetBytes(descriptor.ToString(SaveOptions.DisableFormatting));
        var document = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])); document.Root!.Element(document.Root.Name.Namespace + "Attachments")!.Value = "Metadata/List.xml";
        entries["Doc_0/Document.xml"] = Encoding.UTF8.GetBytes(document.ToString(SaveOptions.DisableFormatting));
        var input = Zip(entries); var source = await Read(input); source.Attachments[0].Data = Edited;
        var output = await Write(OfdDocumentSplitter.Split(source, [1]));
        Assert.Equal("Metadata/List.xml", Document(output).Root!.Element(Ns(output) + "Attachments")!.Value);
        Assert.Equal(Edited, Entries(output)["Doc_0/Attachs/Attach_1_original.bin"]);
        Assert.Equal(AttachmentEntries(input), AttachmentEntries(output));
        Assert.DoesNotContain(Entries(output).Keys, p => p.EndsWith(".1", StringComparison.Ordinal));
        var report = await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output), new());
        Assert.DoesNotContain(report.RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex"); Save("relocated", input, output, report);
    }

    [Fact]
    public async Task FreshUnownedConventionalCollisionsRemainUntouched()
    {
        var package = new OfdDocumentPackage(); var ns = "http://www.ofdspec.org/2016";
        var opaque = Encoding.UTF8.GetBytes($"<Attachments xmlns='{ns}' vendor='unowned'/>");
        package.PreservedEntries["Doc_0/Attachs/Attachments.xml"] = opaque;
        package.PreservedEntries["Doc_0/Attachs/Attach_1_new.bin"] = Original;
        package.Attachments.Add(new OfdAttachment { Name = "new.bin", Data = Edited });
        var output = await Write(package); var entries = Entries(output);
        Assert.Equal(opaque, entries["Doc_0/Attachs/Attachments.xml"]); Assert.Equal(Original, entries["Doc_0/Attachs/Attach_1_new.bin"]);
        Assert.Equal("Attachs/Attachments.xml.1", Document(output).Root!.Element(Ns(output) + "Attachments")!.Value);
        Assert.Equal(Edited, Assert.Single((await Read(output)).Attachments).Data);
        Assert.Contains((await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output), new())).RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
    }

    [Fact]
    public async Task SplitDoesNotOverwriteUnrelatedPayloadPlaceholderWhenAddingFreshAttachment()
    {
        var source = await Read(await Seed()); source.PreservedEntries["Doc_0/Attachs/Attach_3_new.bin"] = Original;
        source.Attachments.Add(new OfdAttachment { Name = "new.bin", Data = Edited });
        var output = await Write(OfdDocumentSplitter.Split(source, [1]));
        Assert.Equal(Original, Entries(output)["Doc_0/Attachs/Attach_3_new.bin"]);
        Assert.Equal(Edited, Entries(output)["Doc_0/Attachs/Attach_3_new.bin.1"]);
        Assert.Equal("Attachs/Attachments.xml", Document(output).Root!.Element(Ns(output) + "Attachments")!.Value);
        Assert.Equal(Edited, (await Read(output)).Attachments[2].Data);
        Assert.DoesNotContain((await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output), new())).RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
    }

    [Fact]
    public async Task SplitSaveAndPreflightDoNotFetchExternalAttachment()
    {
        var listener = new TcpListener(IPAddress.Loopback, 0); listener.Start();
        try
        {
            var package = await Read(await Seed()); var external = $"http://127.0.0.1:{((IPEndPoint)listener.LocalEndpoint).Port}/no-fetch.bin";
            package.Attachments.Add(new OfdAttachment { Name = "external.bin", IsExternal = true, ExternalPath = external });
            var input = await Write(package); var split = OfdDocumentSplitter.Split(await Read(input), [1]);
            var output = await Write(split); var readback = await Read(output);
            var report = await new ArchivalPreflightChecker().CheckAsync(new MemoryStream(output), new());
            Assert.Equal(external, readback.Attachments[2].ExternalPath); Assert.True(readback.Attachments[2].IsExternal); Assert.Empty(readback.Attachments[2].Data);
            Assert.False(listener.Pending()); Assert.Equal(AttachmentEntries(input), AttachmentEntries(output));
            Assert.DoesNotContain(report.RuleResults, r => r.ReasonCode == "OrphanAttachmentIndex");
            Assert.DoesNotContain(Entries(output).Keys, p => p.EndsWith("external.bin", StringComparison.Ordinal)); Save("external", input, output, report);
        }
        finally { listener.Stop(); }
    }

    private static async Task<byte[]> Seed(string prefix = "Doc_0/")
    {
        var package = new OfdDocumentPackage { DocumentEntryPath = prefix + "Document.xml" };
        for (var index = 0; index < 2; index++) package.Pages.Add(new OfdPage { Index = index, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.Attachments.Add(new OfdAttachment { Name = "original.bin", Data = Original });
        package.Attachments.Add(new OfdAttachment { Name = "second.bin", Data = [3, 4, 5] });
        return await Write(package);
    }
    private static async Task<OfdDocumentPackage> Read(byte[] input) => await new OfdReader().ReadAsync(new MemoryStream(input));
    private static async Task<byte[]> Write(OfdDocumentPackage package)
    {
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }
    private static Dictionary<string, byte[]> Entries(byte[] input)
    {
        using var zip = new ZipArchive(new MemoryStream(input));
        return zip.Entries.ToDictionary(e => e.FullName, e => { using var stream = e.Open(); using var output = new MemoryStream(); stream.CopyTo(output); return output.ToArray(); });
    }
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, true))
            foreach (var pair in entries)
            {
                var entry = zip.CreateEntry(pair.Key); entry.LastWriteTime = new DateTimeOffset(2026, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var stream = entry.Open(); stream.Write(pair.Value);
            }
        return output.ToArray();
    }
    private static XDocument Document(byte[] input, string prefix = "Doc_0/") => XDocument.Parse(Encoding.UTF8.GetString(Entries(input)[prefix + "Document.xml"]));
    private static XNamespace Ns(byte[] input, string prefix = "Doc_0/") => Document(input, prefix).Root!.Name.Namespace;
    private static string[] AttachmentEntries(byte[] input, string prefix = "Doc_0/") => Entries(input).Keys.Where(p => p.StartsWith(prefix + "Attachs/", StringComparison.Ordinal)).OrderBy(p => p, StringComparer.Ordinal).ToArray();
    private static void Save(string label, byte[] input, byte[] output, ArchivalPreflightReport report)
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_PR23_SPLIT_ARTIFACTS"); if (directory is null) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, label + "-input.ofd"), input); File.WriteAllBytes(Path.Combine(directory, label + "-output.ofd"), output);
        File.WriteAllText(Path.Combine(directory, label + "-entries.json"), JsonSerializer.Serialize(new { input = Entries(input).Keys, output = Entries(output).Keys }, new JsonSerializerOptions { WriteIndented = true }));
        File.WriteAllText(Path.Combine(directory, label + "-preflight.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    }
}
