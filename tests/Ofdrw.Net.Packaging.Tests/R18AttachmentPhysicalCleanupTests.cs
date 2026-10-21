using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class R18AttachmentPhysicalCleanupTests
{
    private const string Ns = "http://www.ofdspec.org/2016";
    private static readonly byte[] Payload = [0, 1, 2, 3, 255];

    [Fact]
    public async Task LogicalRemovalKeepsHistoricalPayloadWhilePhysicalRemovalDeletesOnlyAuthenticatedPayload()
    {
        var input = Seed(external: true);
        var logical = await Read(input);
        Assert.True(logical.RemoveAttachment(logical.Attachments[0]));
        var (logicalBytes, _) = await Write(logical);
        Assert.True(Entries(logicalBytes).ContainsKey("Doc_0/Payload/original.bin"));

        var physical = await Read(input);
        Assert.True(physical.RemoveAttachment(physical.Attachments[0], OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(physical);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.Contains("Doc_0/Payload/original.bin", result.RemovedEntries);
        Assert.Contains("Doc_0/Metadata/List.xml", Entries(output).Keys);
        Assert.Equal("external.bin", Assert.Single((await Read(output)).Attachments).Name);
        Assert.False(result.SignaturesInvalidated);
        Assert.Single((await Read(output)).Pages);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LastPhysicalRemovalDeletesDescriptorAndSurvivesSplitProvenance(bool requestBeforeSplit)
    {
        var loaded = await Read(Seed());
        if (requestBeforeSplit)
            Assert.True(loaded.RemoveAttachment(Assert.Single(loaded.Attachments), OfdAttachmentRemovalMode.Physical));
        var split = OfdDocumentSplitter.Split(loaded, [0]);
        if (!requestBeforeSplit)
            Assert.True(split.RemoveAttachment(Assert.Single(split.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(split);
        Assert.DoesNotContain("Doc_0/Metadata/List.xml", Entries(output).Keys);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.Contains("Doc_0/Metadata/List.xml", result.RemovedEntries);
        Assert.Contains("Doc_0/Payload/original.bin", result.RemovedEntries);
        Assert.Empty((await Read(output)).Attachments);
        Assert.Empty((await Read((await Write(await Read(output))).Bytes)).Attachments);
    }

    [Fact]
    public async Task NewObjectWithSameIdAndNameCannotBorrowRemovedOrigin()
    {
        var package = await Read(Seed(external: true));
        var original = package.Attachments[0];
        Assert.True(package.RemoveAttachment(original, OfdAttachmentRemovalMode.Physical));
        package.Attachments.Insert(0, new OfdAttachment { Id = original.Id, Name = original.Name, Data = [7, 8, 9] });
        var (output, _) = await Write(package);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.Equal(new byte[] { 7, 8, 9 }, (await Read(output)).Attachments[0].Data);
    }

    [Fact]
    public async Task SharedPayloadAndResourceBaseLocAliasRefuseBeforeOutput()
    {
        var shared = Entries(Seed());
        var descriptor = XDocument.Parse(Encoding.UTF8.GetString(shared["Doc_0/Metadata/List.xml"]));
        var root = descriptor.Root!;
        root.Add(new XElement(XName.Get("Attachment", Ns), new XAttribute("ID", "8"), new XAttribute("Name", "shared.bin"),
            new XElement(XName.Get("FileLoc", Ns), "../Payload/original.bin")));
        shared["Doc_0/Metadata/List.xml"] = Utf8(descriptor.ToString(SaveOptions.DisableFormatting));
        var package = await Read(Zip(shared));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        var aliased = Entries(Seed());
        aliased["Doc_0/Res/DocumentRes.dat"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='../Payload'><MultiMedias><MultiMedia ID='12'><MediaFile>original.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        package = await Read(Zip(aliased));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        aliased["Doc_0/Res/DocumentRes.dat"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='/Doc_0/Payload'><MultiMedias><MultiMedia ID='12'><MediaFile>original.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        package = await Read(Zip(aliased));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task UnknownAndMalformedXmlRefuseButLogicalRemovalStillPreservesSource()
    {
        var opaque = Entries(Seed());
        opaque["Doc_0/Extras/unknown.dat"] = Utf8("<x:Ref xmlns:x='urn:vendor'>../Payload/original.bin</x:Ref>");
        var package = await Read(Zip(opaque));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        opaque["Doc_0/Extras/unknown.dat"] = Utf8("<not-closed");
        package = await Read(Zip(opaque));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        opaque["Doc_0/Extras/unknown.dat"] = Utf8($"<Document xmlns='{Ns}'><EncodedAttachment>UGF5bG9hZC9vcmlnaW5hbC5iaW4=</EncodedAttachment></Document>");
        package = await Read(Zip(opaque));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        opaque["Doc_0/Extras/unknown.dat"] = [4, 5, 6];
        package = await Read(Zip(opaque));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task XmlLookingPayloadCannotBorrowAControlRole()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Payload/original.bin"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='2'/></Content></Page>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task FakeResourceSlotCannotAuthenticateAnOpaqueBinary()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Extras/manifest.bin"] = [4, 5, 6];
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</Document>", "<MediaFile>Extras/manifest.bin</MediaFile></Document>", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task DeclaredResourceTableKeepsItsOwnBinaryWhileCleaningAttachment()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<PublicRes>Res/PublicRes.xml</PublicRes></CommonData>", StringComparison.Ordinal));
        entries["Doc_0/Res/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><MultiMedias><MultiMedia ID='9'><MediaFile>image.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        entries["Doc_0/Res/Assets/image.bin"] = [4, 5, 6];
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.Equal(new byte[] { 4, 5, 6 }, Entries(output)["Doc_0/Res/Assets/image.bin"]);

        // The lexical MediaFile also resolves relative to the table itself,
        // but BaseLoc owns only the Assets path. It must not bless an orphan.
        entries["Doc_0/Res/image.bin"] = [7, 8, 9];
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task EncodedAliasInSignatureIsRejectedBeforeSignaturePruning()
    {
        var entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature BaseLoc='sig/Signature.xml'/></Signatures>");
        entries["Doc_0/Signs/sig/Signature.xml"] = Utf8($"<Signature xmlns='{Ns}'><SignedValue>../../Payload/original%2Ebin</SignedValue></Signature>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task BackslashResourceAliasIsRejectedBeforeOutput()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Res/DocumentRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='../Payload'><MultiMedias><MultiMedia ID='12'><MediaFile>\\original.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task MixedSignaturePathIsRejectedBeforeSignaturePruning()
    {
        var entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature BaseLoc='sig/Signature.xml'/></Signatures>");
        entries["Doc_0/Signs/sig/Signature.xml"] = Utf8($"<Signature xmlns='{Ns}'><SignedValue>../../Payload/orig<FileRef>inal.bin</FileRef></SignedValue></Signature>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task NonXmlSuffixUtf16DescriptorUsesSourcePathAndCanBeRemoved()
    {
        var entries = Entries(Seed());
        var descriptor = Encoding.UTF8.GetString(entries["Doc_0/Metadata/List.xml"]);
        entries.Remove("Doc_0/Metadata/List.xml");
        entries["Doc_0/Metadata/List.dat"] = Encoding.Unicode.GetPreamble()
            .Concat(Encoding.Unicode.GetBytes("<?xml version='1.0' encoding='utf-16'?>" + descriptor)).ToArray();
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("Metadata/List.xml", "Metadata/List.dat", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(package);
        Assert.Contains("Doc_0/Metadata/List.dat", result.RemovedEntries);
        Assert.DoesNotContain("Doc_0/Metadata/List.dat", Entries(output).Keys);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
    }

    [Fact]
    public async Task MultipleDocBodiesAndSignatureAliasRefuseEvenIfSignaturesWouldBePruned()
    {
        var entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody><DocBody><DocRoot>Other/Document.xml</DocRoot></DocBody></OFD>");
        entries["Other/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><CommonData><MaxUnitID>1</MaxUnitID><PageArea><PhysicalBox>0 0 210 297</PhysicalBox></PageArea></CommonData><Pages><Page ID='1' BaseLoc='Pages/Content.xml'/></Pages></Document>");
        entries["Other/Pages/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='2'/></Content></Page>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature BaseLoc='sig/Signature.xml'/></Signatures>");
        entries["Doc_0/Signs/sig/Signature.xml"] = Utf8($"<Signature xmlns='{Ns}'><SignedValue>../../Payload/original.bin</SignedValue></Signature>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task ControlledClonePreservesPhysicalRequestAndItsRemovalDisposition()
    {
        var input = Seed();
        using var context = new OfdControlledEditContext(new());
        var loaded = await new OfdReader().ReadControlledAsync(new MemoryStream(input), new(), context);
        Assert.True(loaded.RemoveAttachment(Assert.Single(loaded.Attachments), OfdAttachmentRemovalMode.Physical));
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(loaded, context.CaptureSource(loaded), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        using var output = new MemoryStream();
        await prepared.WriteAsync(output);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output.ToArray()).Keys);
        Assert.DoesNotContain("Doc_0/Metadata/List.xml", Entries(output.ToArray()).Keys);
    }

    [Fact]
    public async Task SignedPackageReportsInvalidationAndPhysicalDeletion()
    {
        var signed = Entries(Seed());
        signed["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        signed["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature BaseLoc='sig/Signature.xml'/></Signatures>");
        signed["Doc_0/Signs/sig/Signature.xml"] = Utf8($"<Signature xmlns='{Ns}'><SignedValue>SignedValue.dat</SignedValue></Signature>");
        signed["Doc_0/Signs/sig/SignedValue.dat"] = [7, 8, 9];
        var package = await Read(Zip(signed));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(package);
        Assert.True(result.SignaturesInvalidated);
        Assert.Contains(result.Diagnostics, item => item.Contains("sign", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.DoesNotContain("Doc_0/Signs/sig/SignedValue.dat", Entries(output).Keys);
        Assert.Null(XDocument.Parse(Encoding.UTF8.GetString(Entries(output)["OFD.xml"])).Root!
            .Descendants(XName.Get("Signatures", Ns)).FirstOrDefault());
    }

    [Fact]
    public async Task FreshAndExternalObjectsCannotRequestPhysicalDeletion()
    {
        var fresh = new OfdDocumentPackage();
        var attachment = new OfdAttachment { Name = "new.bin", Data = Payload };
        fresh.Attachments.Add(attachment);
        Assert.Throws<NotSupportedException>(() => fresh.RemoveAttachment(attachment, OfdAttachmentRemovalMode.Physical));
        Assert.Single(fresh.Attachments);
        var imported = await Read(Seed(external: true));
        Assert.Throws<NotSupportedException>(() => imported.RemoveAttachment(imported.Attachments[1], OfdAttachmentRemovalMode.Physical));
        Assert.Equal(2, imported.Attachments.Count);
    }

    private static async Task RefusesWithoutOutput(OfdDocumentPackage package)
    {
        var preserved = package.PreservedEntries.ToDictionary(entry => entry.Key, entry => entry.Value.ToArray());
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteWithResultAsync(package, output));
        Assert.Equal(0, output.Length);
        Assert.Equal(preserved.Keys.OrderBy(path => path), package.PreservedEntries.Keys.OrderBy(path => path));
        foreach (var entry in preserved)
            Assert.Equal(entry.Value, package.PreservedEntries[entry.Key]);
    }

    private static byte[] Seed(bool external = false)
    {
        var entries = new Dictionary<string, byte[]>
        {
            ["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>"),
            ["Doc_0/Document.xml"] = Utf8($"<Document xmlns='{Ns}'><CommonData><MaxUnitID>2</MaxUnitID><PageArea><PhysicalBox>0 0 210 297</PhysicalBox></PageArea></CommonData><Pages><Page ID='1' BaseLoc='Pages/Content.xml'/></Pages><Attachments>Metadata/List.xml</Attachments></Document>"),
            ["Doc_0/Pages/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='2'/></Content></Page>"),
            ["Doc_0/Metadata/List.xml"] = Utf8($"<Attachments xmlns='{Ns}'><Attachment ID='7' Name='original.bin'><FileLoc>../Payload/original.bin</FileLoc></Attachment>{(external ? "<Attachment ID='8' Name='external.bin' External='true'><FileLoc>https://invalid.example/never-fetch.bin</FileLoc></Attachment>" : "")}</Attachments>"),
            ["Doc_0/Payload/original.bin"] = Payload
        };
        return Zip(entries);
    }

    private static byte[] Utf8(string value) => Encoding.UTF8.GetBytes(value);
    private static Dictionary<string, byte[]> Entries(byte[] bytes)
    {
        using var zip = new ZipArchive(new MemoryStream(bytes));
        var result = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            using var source = entry.Open(); using var output = new MemoryStream(); source.CopyTo(output); result.Add(entry.FullName, output.ToArray());
        }
        return result;
    }
    private static byte[] Zip(Dictionary<string, byte[]> entries)
    {
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var pair in entries)
            {
                var entry = zip.CreateEntry(pair.Key);
                using var target = entry.Open(); target.Write(pair.Value);
            }
        return output.ToArray();
    }
    private static Task<OfdDocumentPackage> Read(byte[] bytes) => new OfdReader().ReadAsync(new MemoryStream(bytes));
    private static async Task<(byte[] Bytes, OfdPackageWriteResult Result)> Write(OfdDocumentPackage package)
    {
        using var output = new MemoryStream();
        var result = await new OfdPackageWriter().WriteWithResultAsync(package, output);
        return (output.ToArray(), result);
    }
}
