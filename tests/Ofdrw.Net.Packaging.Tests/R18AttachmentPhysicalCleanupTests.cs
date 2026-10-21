using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Signing;

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

        var templateAlias = Entries(Seed());
        templateAlias["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(templateAlias["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<TemplatePage ID='500' BaseLoc='Templates/Content.xml'/></CommonData>", StringComparison.Ordinal));
        templateAlias["Doc_0/Pages/Content.xml"] = Utf8(Encoding.UTF8.GetString(templateAlias["Doc_0/Pages/Content.xml"])
            .Replace("</Page>", "<Template TemplateID='500'/></Page>", StringComparison.Ordinal));
        templateAlias["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='501'/></Content></Page>");
        templateAlias["Doc_0/Templates/PageRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='../Payload'><MultiMedias><MultiMedia ID='502'><MediaFile>original.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        package = await Read(Zip(templateAlias));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
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

    [Theory]
    [InlineData(false, true)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(true, false)]
    public async Task BomlessMultibyteXmlAttachmentCannotBePhysicallyDeleted(bool utf32, bool bigEndian)
    {
        var entries = Entries(Seed());
        Encoding encoding = utf32 ? new UTF32Encoding(bigEndian, false) : new UnicodeEncoding(bigEndian, false);
        var name = utf32 ? (bigEndian ? "UTF-32BE" : "UTF-32LE") : (bigEndian ? "UTF-16BE" : "UTF-16LE");
        entries["Doc_0/Payload/original.bin"] = encoding.GetBytes($"<?xml version='1.0' encoding='{name}'?><Page xmlns='{Ns}'/>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "lost its reader-owned payload");
    }

    [Fact]
    public async Task BomlessMultibyteRetainedXmlIsScannedForAttachmentReferences()
    {
        var entries = Entries(Seed());
        var descriptor = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Metadata/List.xml"]));
        descriptor.Root!.Add(new XElement(XName.Get("Attachment", Ns), new XAttribute("ID", "8"),
            new XAttribute("Name", "retained.bin"), new XElement(XName.Get("FileLoc", Ns), "../Payload/retained.bin")));
        entries["Doc_0/Metadata/List.xml"] = Utf8(descriptor.ToString(SaveOptions.DisableFormatting));
        entries["Doc_0/Payload/retained.bin"] = new UnicodeEncoding(true, false)
            .GetBytes($"<Page xmlns='{Ns}'><Title>original.bin</Title></Page>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "remains referenced");
    }

    [Fact]
    public async Task RetainedAuthenticatedBinaryWithXmlSuffixDoesNotBlockOtherPayloadRemoval()
    {
        var entries = Entries(Seed());
        var descriptor = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Metadata/List.xml"]));
        descriptor.Root!.Add(new XElement(XName.Get("Attachment", Ns), new XAttribute("ID", "8"),
            new XAttribute("Name", "retained.xml"), new XElement(XName.Get("FileLoc", Ns), "../Payload/retained.xml")));
        entries["Doc_0/Metadata/List.xml"] = Utf8(descriptor.ToString(SaveOptions.DisableFormatting));
        entries["Doc_0/Payload/retained.xml"] = [4, 5, 6];

        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.Equal(new byte[] { 4, 5, 6 }, Entries(output)["Doc_0/Payload/retained.xml"]);
        Assert.Equal("retained.xml", Assert.Single((await Read(output)).Attachments).Name);

        entries["Doc_0/Extras/unowned.xml"] = [7, 8, 9];
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task ExternalAttachmentFileLocCannotAuthenticateOpaqueArchiveBytes()
    {
        var entries = Entries(Seed());
        var descriptor = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Metadata/List.xml"]));
        descriptor.Root!.Add(new XElement(XName.Get("Attachment", Ns), new XAttribute("ID", "8"),
            new XAttribute("Name", "external.bin"), new XAttribute("External", "true"),
            new XElement(XName.Get("FileLoc", Ns), "../Extras/orphan.bin")));
        entries["Doc_0/Metadata/List.xml"] = Utf8(descriptor.ToString(SaveOptions.DisableFormatting));
        entries["Doc_0/Extras/orphan.bin"] = [4, 5, 6];

        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "opaque binary entry ownership");
    }

    [Fact]
    public async Task FakeOfdCoverCannotAuthenticateOpaqueArchiveBytes()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Extras/Fake.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocInfo><Cover>manifest.bin</Cover></DocInfo></DocBody></OFD>");
        entries["Doc_0/Extras/manifest.bin"] = [4, 5, 6];

        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "opaque binary entry ownership");
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

        entries = Entries(Seed());
        entries["Doc_0/Extras/manifest.bin"] = [4, 5, 6];
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</Document>", "<MultiMedia ID='9' MediaFile='Extras/manifest.bin'/></Document>", StringComparison.Ordinal));
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<PublicRes>Res/PublicRes.xml</PublicRes></CommonData>", StringComparison.Ordinal));
        entries["Doc_0/Res/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}'><Fonts><Font><FontFile>orphan.bin</FontFile></Font></Fonts><MultiMedias><MultiMedia MediaFile='orphan.bin'/></MultiMedias></Res>");
        entries["Doc_0/Res/orphan.bin"] = [7, 8, 9];
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task NestedBaseLocCannotAuthenticateOpaqueResourceOrAttachmentBytes()
    {
        var resource = Entries(Seed());
        resource["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(resource["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<PublicRes>Res/PublicRes.xml</PublicRes></CommonData>", StringComparison.Ordinal));
        resource["Doc_0/Res/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><MultiMedias><MultiMedia ID='9' BaseLoc='../Extras' MediaFile='orphan.bin'/></MultiMedias></Res>");
        resource["Doc_0/Res/Extras/orphan.bin"] = [4, 5, 6];
        var package = await Read(Zip(resource));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "opaque binary entry ownership");

        resource.Remove("Doc_0/Res/Extras/orphan.bin");
        resource["Doc_0/Res/Assets/orphan.bin"] = [4, 5, 6];
        package = await Read(Zip(resource));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (resourceOutput, _) = await Write(package);
        Assert.Equal(resource["Doc_0/Res/Assets/orphan.bin"], Entries(resourceOutput)["Doc_0/Res/Assets/orphan.bin"]);

        var attachment = Entries(Seed());
        var descriptor = XDocument.Parse(Encoding.UTF8.GetString(attachment["Doc_0/Metadata/List.xml"]));
        descriptor.Root!.Add(new XElement(XName.Get("Attachment", Ns), new XAttribute("ID", "8"),
            new XAttribute("Name", "retained.bin"), new XAttribute("BaseLoc", "../Extras"),
            new XElement(XName.Get("FileLoc", Ns), "retained.bin")));
        attachment["Doc_0/Metadata/List.xml"] = Utf8(descriptor.ToString(SaveOptions.DisableFormatting));
        attachment["Doc_0/Metadata/retained.bin"] = [7, 8, 9];
        attachment["Doc_0/Extras/retained.bin"] = [10, 11, 12];
        package = await Read(Zip(attachment));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "opaque binary entry ownership");

        attachment.Remove("Doc_0/Extras/retained.bin");
        package = await Read(Zip(attachment));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        var (attachmentOutput, _) = await Write(package);
        Assert.Equal(attachment["Doc_0/Metadata/retained.bin"], Entries(attachmentOutput)["Doc_0/Metadata/retained.bin"]);
    }

    [Fact]
    public async Task TemplateWithoutReaderAcceptedIdFailsBeforeAdjacentResourceCanBeWritten()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<TemplatePage BaseLoc='Templates/Content.xml'/></CommonData>", StringComparison.Ordinal));
        entries["Doc_0/Templates/Content.xml"] = Utf8($"<Page xmlns='{Ns}'><Content><Layer ID='9'/></Content></Page>");
        entries["Doc_0/Templates/PageRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><MultiMedias><MultiMedia ID='10' MediaFile='orphan.bin'/></MultiMedias></Res>");
        entries["Doc_0/Templates/Assets/orphan.bin"] = [4, 5, 6];
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteWithResultAsync(package, output));
        Assert.Contains("Template declaration requires a positive numeric ID", error.Message);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnusedRetainedAttachmentFileLocCannotAuthenticateOpaqueBytes(bool duplicateElement)
    {
        var entries = Entries(Seed());
        var descriptor = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Metadata/List.xml"]));
        var retained = new XElement(XName.Get("Attachment", Ns), new XAttribute("ID", "8"),
            new XAttribute("Name", "retained.bin"),
            new XElement(XName.Get("FileLoc", Ns), "retained.bin"));
        if (duplicateElement)
            retained.Add(new XElement(XName.Get("FileLoc", Ns), "orphan.bin"));
        else
            retained.Add(new XAttribute("FileLoc", "orphan.bin"));
        descriptor.Root!.Add(retained);
        entries["Doc_0/Metadata/List.xml"] = Utf8(descriptor.ToString(SaveOptions.DisableFormatting));
        entries["Doc_0/Metadata/retained.bin"] = [7, 8, 9];
        entries["Doc_0/Metadata/orphan.bin"] = [10, 11, 12];
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "opaque binary entry ownership");

        entries.Remove("Doc_0/Metadata/orphan.bin");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.Equal(entries["Doc_0/Metadata/retained.bin"], Entries(output)["Doc_0/Metadata/retained.bin"]);
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

        var legacy = Entries(Seed());
        legacy["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(legacy["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<PublicRes>Res/PublicRes.xml</PublicRes></CommonData>", StringComparison.Ordinal));
        legacy["Doc_0/Res/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><MultiMedias><MultiMedia ID='9' MediaFile='image.bin'/></MultiMedias></Res>");
        legacy["Doc_0/Res/Assets/image.bin"] = [4, 5, 6];
        package = await Read(Zip(legacy));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        (output, _) = await Write(package);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.Equal(new byte[] { 4, 5, 6 }, Entries(output)["Doc_0/Res/Assets/image.bin"]);

        legacy["Doc_0/Res/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><MultiMedias><MultiMedia ID='9' MediaFile='image.bin'><MediaFile>orphan.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        legacy["Doc_0/Res/Assets/orphan.bin"] = [7, 8, 9];
        package = await Read(Zip(legacy));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        // The lexical MediaFile also resolves relative to the table itself,
        // but BaseLoc owns only the Assets path. It must not bless an orphan.
        entries["Doc_0/Res/image.bin"] = [7, 8, 9];
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task OpaqueXmlSuffixResourceIsOwnedRegardlessOfZipEntryOrder()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<PublicRes>Res/PublicRes.xml</PublicRes></CommonData>", StringComparison.Ordinal));
        entries["Doc_0/Res/Assets/media.xml"] = [4, 5, 6];
        entries["Doc_0/Res/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><MultiMedias><MultiMedia ID='9'><MediaFile>media.xml</MediaFile></MultiMedia></MultiMedias></Res>");

        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.Equal(entries["Doc_0/Res/Assets/media.xml"], Entries(output)["Doc_0/Res/Assets/media.xml"]);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
    }

    [Fact]
    public async Task PublicFontMetadataSurvivesPhysicalCleanupWithoutResourceRelocation()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<PublicRes>PublicRes.xml</PublicRes></CommonData>", StringComparison.Ordinal));
        entries["Doc_0/PublicRes.xml"] = Utf8($"<Res xmlns='{Ns}' BaseLoc='Assets'><Fonts><Font ID='3' FontName='R18' Serif='true' FixedWidth='false'><FontFile>font.bin</FontFile></Font></Fonts></Res>");
        entries["Doc_0/Assets/font.bin"] = [4, 5, 6];
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        var saved = Entries(output);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", saved.Keys);
        Assert.Equal(entries["Doc_0/Assets/font.bin"], saved["Doc_0/Assets/font.bin"]);
        var resource = XDocument.Parse(Encoding.UTF8.GetString(saved["Doc_0/PublicRes.xml"]).TrimStart('\uFEFF'));
        var font = resource.Descendants(XName.Get("Font", Ns)).Single();
        Assert.Equal("true", font.Attribute("Serif")!.Value);
        Assert.Equal("false", font.Attribute("FixedWidth")!.Value);
    }

    [Fact]
    public async Task WriterGeneratedPathAndClipRemainAfterPhysicalCleanup()
    {
        var package = await Read(Seed());
        var graphics = new OfdGraphics(package, Assert.Single(package.Pages));
        graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(1, 1, 20, 20));
        graphics.DrawRectangle(new OfdPen(OfdColor.Black), 2, 2, 10, 10);
        package.Pages[0].Elements.Add(new OfdImageElement
        {
            Data = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScLbtAAAAABJRU5ErkJggg=="),
            WidthMillimeters = 10, HeightMillimeters = 10
        });
        var (drawn, _) = await Write(package);
        var entries = Entries(drawn);
        var page = XDocument.Parse(Encoding.UTF8.GetString(entries["Doc_0/Pages/Content.xml"]).TrimStart('\uFEFF'));
        var ns = page.Root!.Name.Namespace;
        Assert.NotEmpty(page.Descendants(ns + "Path"));
        Assert.Contains("Rule=", page.ToString());
        page.Descendants(ns + "Clip").Descendants(ns + "Area").First().SetAttributeValue("Start", "0 0");
        var image = page.Descendants(ns + "ImageObject").Single();
        image.SetAttributeValue("Substitution", image.Attribute("ResourceID")!.Value);
        var imageHint = XName.Get("PdfInterpolateV1", "https://ofdrw.net/image-hints");
        image.SetAttributeValue(imageHint, "true");
        entries["Doc_0/Pages/Content.xml"] = Utf8(page.ToString(SaveOptions.DisableFormatting));

        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        var savedPage = XDocument.Parse(Encoding.UTF8.GetString(Entries(output)["Doc_0/Pages/Content.xml"]).TrimStart('\uFEFF'));
        Assert.Contains(savedPage.Descendants(ns + "Clip").Descendants(ns + "Area"),
            area => area.Attribute("Start")?.Value == "0 0");
        var savedImage = savedPage.Descendants(ns + "ImageObject").Single();
        Assert.Equal(savedImage.Attribute("ResourceID")!.Value, savedImage.Attribute("Substitution")!.Value);
        Assert.Equal("true", savedImage.Attribute(imageHint)!.Value);
        Assert.Single((await Read(output)).Pages);
    }

    [Fact]
    public async Task WriterOwnedItalicHintSurvivesPhysicalCleanup()
    {
        var package = await Read(Seed());
        package.Pages[0].Elements.Add(new OfdTextElement
        {
            Text = "italic", FontName = "SimSun", FontSizeMillimeters = 4,
            XMillimeters = 10, YMillimeters = 10, Italic = true
        });
        var (drawn, _) = await Write(package);
        var page = XDocument.Parse(Encoding.UTF8.GetString(Entries(drawn)["Doc_0/Pages/Content.xml"]).TrimStart('\uFEFF'));
        var italicHint = XName.Get("FauxItalicMatrixV1", "https://ofdrw.net/style-hints");
        Assert.NotNull(page.Descendants(XName.Get("TextObject", Ns)).Single().Attribute(italicHint));

        package = await Read(drawn);
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        var savedPage = XDocument.Parse(Encoding.UTF8.GetString(Entries(output)["Doc_0/Pages/Content.xml"]).TrimStart('\uFEFF'));
        Assert.NotNull(savedPage.Descendants(XName.Get("TextObject", Ns)).Single().Attribute(italicHint));

        var foreign = Entries(drawn);
        page.Descendants(XName.Get("TextObject", Ns)).Single()
            .SetAttributeValue(XName.Get("FauxItalicMatrixV1", "urn:vendor"), "1");
        foreign["Doc_0/Pages/Content.xml"] = Utf8(page.ToString(SaveOptions.DisableFormatting));
        package = await Read(Zip(foreign));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Fact]
    public async Task StandardTemplateAndWriterAnnotationSurvivePhysicalCleanup()
    {
        var fresh = new OfdDocumentPackage { DocumentEntryPath = "Doc_0/Document.xml" };
        fresh.Attachments.Add(new OfdAttachment { Name = "remove.bin", Data = Payload });
        fresh.Pages.Add(new OfdPage
        {
            WidthMillimeters = 210, HeightMillimeters = 297,
            Annotations = new OfdAnnotationCollection
            {
                Items = { new OfdAnnotation
                {
                    Type = OfdAnnotationType.Stamp, Creator = "R18", LastModifiedDate = new DateTime(2026, 10, 21),
                    Subtype = "Seal", Remark = "retain", Visible = true, Print = true, NoZoom = false,
                    NoRotate = true, ReadOnly = true,
                    Appearance = new() { XMillimeters = 10, YMillimeters = 15, WidthMillimeters = 20, HeightMillimeters = 20 }
                } }
            }
        });
        var (generated, _) = await Write(fresh);
        var generatedModel = await Read(generated);
        var entries = Entries(generated);
        var documentPath = generatedModel.DocumentEntryPath!;
        var pagePath = Assert.Single(generatedModel.Pages).SourceEntryPath!;
        var annotationPath = Assert.Single(generatedModel.Pages[0].Annotations!.Items).Source!.Entry;
        var document = XDocument.Parse(Encoding.UTF8.GetString(entries[documentPath]));
        var ns = document.Root!.Name.Namespace;
        var annotationIndexPath = "Doc_0/" + document.Root.Element(ns + "Annotations")!.Value;
        document.Root!.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage",
            new XAttribute("ID", "500"), new XAttribute("BaseLoc", "Templates/Content.xml")));
        entries[documentPath] = Utf8(document.ToString(SaveOptions.DisableFormatting));
        var page = XDocument.Parse(Encoding.UTF8.GetString(entries[pagePath]));
        page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "500"),
            new XAttribute("ZOrder", "Background")));
        entries[pagePath] = Utf8(page.ToString(SaveOptions.DisableFormatting));
        const string templatePath = "Doc_0/Templates/Content.xml";
        entries[templatePath] = Utf8($"<Page xmlns='{ns}'><Content><Layer ID='501'><TextObject ID='504' Font='502' Size='4' Boundary='0 0 20 10'><TextCode X='0' Y='4'>TEMPLATE</TextCode></TextObject><ImageObject ID='505' ResourceID='503' Boundary='0 0 10 10'/></Layer></Content></Page>");
        const string templateResourcePath = "Doc_0/Templates/PageRes.xml";
        const string templateFontPath = "Doc_0/Templates/Assets/font.bin";
        const string templateMediaPath = "Doc_0/Templates/Assets/image.bin";
        entries[templateResourcePath] = Utf8($"<Res xmlns='{ns}' BaseLoc='Assets'><Fonts><Font ID='502' FontName='R18'><FontFile>font.bin</FontFile></Font></Fonts><MultiMedias><MultiMedia ID='503' Format='PNG'><MediaFile>image.bin</MediaFile></MultiMedia></MultiMedias></Res>");
        entries[templateFontPath] = [4, 5, 6];
        entries[templateMediaPath] = [7, 8, 9];
        var originalAnnotation = entries[annotationPath].ToArray();
        var originalAnnotationIndex = entries[annotationIndexPath].ToArray();
        var originalTemplate = entries[templatePath].ToArray();

        var loaded = await Read(Zip(entries));
        Assert.Single(loaded.Pages[0].Templates);
        Assert.True(loaded.RemoveAttachment(Assert.Single(loaded.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(loaded);
        var saved = Entries(output);
        Assert.Contains(result.RemovedEntries, path => path.Contains("remove.bin", StringComparison.Ordinal));
        Assert.Equal(originalAnnotation, saved[annotationPath]);
        Assert.Equal(originalAnnotationIndex, saved[annotationIndexPath]);
        Assert.Equal(originalTemplate, saved[templatePath]);
        var savedTemplateResources = XDocument.Parse(Encoding.UTF8.GetString(saved[templateResourcePath]).TrimStart('\uFEFF'));
        Assert.Equal("Assets", savedTemplateResources.Root!.Attribute("BaseLoc")!.Value);
        Assert.Equal("image.bin", savedTemplateResources.Descendants(ns + "MediaFile").Single().Value);
        Assert.Equal(entries[templateMediaPath], saved[templateMediaPath]);
        Assert.Contains(saved.Values, bytes => bytes.SequenceEqual(entries[templateFontPath]));
        Assert.Equal("500", XDocument.Parse(Encoding.UTF8.GetString(saved[pagePath]))
            .Descendants(ns + "Template").Single().Attribute("TemplateID")!.Value);
        Assert.Equal("Templates/Content.xml", XDocument.Parse(Encoding.UTF8.GetString(saved[documentPath]))
            .Descendants(ns + "TemplatePage").Single().Attribute("BaseLoc")!.Value);
        var reread = await Read(output);
        var savedTemplate = Assert.Single(reread.Pages[0].Templates);
        Assert.Contains(savedTemplate.Elements, element => element is OfdTextElement { Text: "TEMPLATE" });
        Assert.Contains(savedTemplate.Elements, element => element is OfdImageElement image && image.Data.SequenceEqual(entries[templateMediaPath]));
        Assert.Contains(reread.Fonts, font => font.Data is not null && font.Data.SequenceEqual(entries[templateFontPath]));
        Assert.Equal("retain", Assert.Single(reread.Pages[0].Annotations!.Items).Remark);
    }

    [Fact]
    public async Task WriterNavigationAndEmrTagsSurvivePhysicalCleanup()
    {
        var fresh = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 };
        fresh.Pages.Add(page);
        fresh.Attachments.Add(new OfdAttachment { Name = "remove.bin", Data = Payload });
        var bookmark = new OfdBookmark { Name = "R18 landmark", Destination = new OfdDestination
        {
            Page = page, Type = OfdDestinationType.FitR, Left = 1, Top = 2, Right = 20, Bottom = 30, Zoom = 1.25
        } };
        fresh.Bookmarks = new OfdBookmarkCollection();
        fresh.Bookmarks.Items.Add(bookmark);
        fresh.Outlines = new OfdOutlineCollection();
        var outline = new OfdOutline { Title = "R18 outline", Expanded = true, Count = 1, Actions = new OfdActionCollection() };
        outline.Actions.Items.Add(new OfdGoToAction { Bookmark = bookmark, Event = "CLICK" });
        fresh.Outlines.Items.Add(outline);
        fresh.Actions = new OfdActionCollection();
        fresh.Actions.Items.Add(new OfdGoToAction { Destination = new OfdDestination { Page = page, Type = OfdDestinationType.Fit }, Event = "DO" });
        fresh.CustomTags.Add("patient", "保留 & 123");

        var (generated, _) = await Write(fresh);
        var source = Entries(generated);
        var loaded = await Read(generated);
        var documentPath = loaded.DocumentEntryPath!;
        var originalDocument = XDocument.Parse(Encoding.UTF8.GetString(source[documentPath]));
        var ns = originalDocument.Root!.Name.Namespace;
        var tagIndexPath = Assert.Single(source.Keys, path => path.EndsWith("/CustomTags.xml", StringComparison.Ordinal));
        var tagDetailPath = Assert.Single(source.Keys, path => path.EndsWith("/CustomTag_EMR.xml", StringComparison.Ordinal));
        Assert.True(loaded.RemoveAttachment(Assert.Single(loaded.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(loaded);
        var saved = Entries(output);
        Assert.Contains(result.RemovedEntries, path => path.Contains("remove.bin", StringComparison.Ordinal));
        Assert.Equal(source[tagIndexPath], saved[tagIndexPath]);
        Assert.Equal(source[tagDetailPath], saved[tagDetailPath]);
        var savedDocument = XDocument.Parse(Encoding.UTF8.GetString(saved[documentPath]));
        foreach (var name in new[] { "Outlines", "Bookmarks", "Actions", "CustomTags" })
            Assert.True(XNode.DeepEquals(originalDocument.Root.Element(ns + name), savedDocument.Root!.Element(ns + name)), name);
        var reread = await Read(output);
        Assert.Equal("R18 outline", Assert.Single(reread.Outlines!.Items).Title);
        Assert.Equal("R18 landmark", Assert.Single(reread.Bookmarks!.Items).Name);
        Assert.IsType<OfdGoToAction>(Assert.Single(reread.Actions!.Items));
        Assert.Equal("保留 & 123", reread.CustomTags["patient"]);
    }

    [Fact]
    public async Task ImportedTextDirectionsSurvivePhysicalCleanup()
    {
        var fresh = new OfdDocumentPackage();
        fresh.Attachments.Add(new OfdAttachment { Name = "remove.bin", Data = Payload });
        fresh.Pages.Add(new OfdPage
        {
            WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdTextElement { Text = "R18", FontName = "SimSun", FontSizeMillimeters = 4 } }
        });
        var (generated, _) = await Write(fresh);
        var entries = Entries(generated);
        var pagePath = Assert.Single((await Read(generated)).Pages).SourceEntryPath!;
        var page = XDocument.Parse(Encoding.UTF8.GetString(entries[pagePath]));
        var text = page.Descendants(page.Root!.Name.Namespace + "TextObject").Single();
        text.SetAttributeValue("ReadDirection", "90");
        text.SetAttributeValue("CharDirection", "90");
        entries[pagePath] = Utf8(page.ToString(SaveOptions.DisableFormatting));

        var loaded = await Read(Zip(entries));
        Assert.True(loaded.RemoveAttachment(Assert.Single(loaded.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(loaded);
        var savedText = XDocument.Parse(Encoding.UTF8.GetString(Entries(output)[pagePath]))
            .Descendants(page.Root.Name.Namespace + "TextObject").Single();
        Assert.Equal("90", savedText.Attribute("ReadDirection")!.Value);
        Assert.Equal("90", savedText.Attribute("CharDirection")!.Value);
        Assert.Equal("R18", Assert.IsType<OfdTextElement>(Assert.Single((await Read(output)).Pages[0].Elements)).Text);
    }

    [Fact]
    public async Task LiteralBackslashAndPercentTextDoNotBlockPhysicalCleanup()
    {
        const string literal = @"C:\Temp %2F literal";
        var source = await Read(Seed());
        source.Options.Metadata.Title = literal;
        source.CustomTags.Add("display", literal);
        source.Pages[0].Elements.Add(new OfdTextElement { Text = literal, FontName = "SimSun", FontSizeMillimeters = 4 });
        var (generated, _) = await Write(source);
        var loaded = await Read(generated);
        var originalText = Assert.IsType<OfdTextElement>(Assert.Single(loaded.Pages[0].Elements)).Text;
        Assert.True(loaded.RemoveAttachment(Assert.Single(loaded.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(loaded);
        Assert.Contains(result.RemovedEntries, path => path == "Doc_0/Payload/original.bin");
        var reread = await Read(output);
        Assert.Equal(literal, reread.Options.Metadata.Title);
        Assert.Equal(literal, reread.CustomTags["display"]);
        Assert.Equal(originalText, Assert.IsType<OfdTextElement>(Assert.Single(reread.Pages[0].Elements)).Text);
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

        entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures><Name>Doc_0/Payload/</Name><Title>original.bin</Title></Signatures></DocBody></OFD>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "mixed XML path content");
    }

    [Fact]
    public async Task WhitespaceAroundSignaturePathCannotHideAttachmentAlias()
    {
        var entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature BaseLoc='sig/Signature.xml'/></Signatures>");
        entries["Doc_0/Signs/sig/Signature.xml"] = Utf8($"<Signature xmlns='{Ns}'><SignedValue> ../../Payload/original.bin </SignedValue></Signature>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "surrounding whitespace");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ProcessingInstructionCannotHideAttachmentReference(bool documentLevel)
    {
        var entries = Entries(Seed());
        const string instruction = "<?xml-stylesheet href='../Payload/original.bin'?>";
        var document = Encoding.UTF8.GetString(entries["Doc_0/Document.xml"]);
        entries["Doc_0/Document.xml"] = Utf8(documentLevel
            ? instruction + document
            : document.Replace("</Document>", instruction + "</Document>", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "processing instruction reference closure");
    }

    [Fact]
    public async Task ImportedSealBaseLocElementIsScannedWhileInvalidatingSignature()
    {
        var entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody></OFD>");
        entries["Doc_0/Signs/Signatures.xml"] = Utf8($"<Signatures xmlns='{Ns}'><Signature BaseLoc='sig/Signature.xml'/></Signatures>");
        const string signaturePath = "Doc_0/Signs/sig/Signature.xml";
        entries[signaturePath] = Utf8($"<Signature xmlns='{Ns}'><SignedInfo><Seal><BaseLoc>Seal.esl</BaseLoc></Seal></SignedInfo><SignedValue>Value.dat</SignedValue></Signature>");
        entries["Doc_0/Signs/sig/Seal.esl"] = [4, 5, 6];
        entries["Doc_0/Signs/sig/Value.dat"] = [7, 8, 9];
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(package);
        Assert.True(result.SignaturesInvalidated);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);

        entries[signaturePath] = Utf8($"<Signature xmlns='{Ns}'><SignedInfo><Seal><BaseLoc>../../Payload/original%2Ebin</BaseLoc></Seal></SignedInfo><SignedValue>Value.dat</SignedValue></Signature>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);

        entries[signaturePath] = Utf8($"<Signature xmlns='{Ns}'><SignedInfo><Seal><BaseLoc><Title>../../Payload/</Title><Author>original.bin</Author></BaseLoc></Seal></SignedInfo><SignedValue>Value.dat</SignedValue></Signature>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task DocumentExtensionsPathCannotHideAttachmentAlias(bool split)
    {
        var entries = Entries(Seed());
        var extension = split
            ? "<Extensions><Title>Payload/</Title><Author>original.bin</Author></Extensions>"
            : "<Extensions>Payload/original%2Ebin</Extensions>";
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</Document>", extension + "</Document>", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, split ? "mixed XML path content" : "encoded or backslash XML reference");
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task MalformedDocBodyVersionsCannotHideAttachmentAlias(bool split)
    {
        var entries = Entries(Seed());
        var versions = split
            ? "<Versions><Title>Doc_0/Payload/</Title><Author>original.bin</Author></Versions>"
            : "<Versions>Doc_0/Payload/original%2Ebin</Versions>";
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot>{versions}</DocBody></OFD>");
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "malformed Versions wrapper");
    }

    [Fact]
    public async Task StandardVersionsWrapperPreservesVersionFileAndChecksItsReferences()
    {
        var entries = Entries(Seed());
        var versions = $"<Versions><Version ID='v1' Index='1' Current='true' BaseLoc='Doc_0/Versions/v1.xml'/></Versions>";
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot>{versions}</DocBody></OFD>");
        entries["Doc_0/Versions/v1.xml"] = Utf8($"<DocVersion xmlns='{Ns}' ID='v1'><FileList><File ID='f1'>../Pages/Content.xml</File></FileList><DocRoot>../Document.xml</DocRoot></DocVersion>");
        var package = await Read(Zip(entries));
        Assert.Single(package.Versions!.Items).Index = 2;
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.Equal(entries["Doc_0/Versions/v1.xml"], Entries(output)["Doc_0/Versions/v1.xml"]);
        Assert.Equal("2", XDocument.Parse(Encoding.UTF8.GetString(Entries(output)["OFD.xml"]))
            .Descendants(XName.Get("Versions", Ns)).Single().Element(XName.Get("Version", Ns))!.Attribute("Index")!.Value);

        entries["Doc_0/Versions/v1.xml"] = Utf8($"<DocVersion xmlns='{Ns}' ID='v1'><FileList><File ID='f1'>../Payload/original%2Ebin</File></FileList><DocRoot>../Document.xml</DocRoot></DocVersion>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "encoded or backslash XML reference");

        entries["Doc_0/Versions/v1.xml"] = Utf8($"<DocVersion xmlns='{Ns}' ID='v1'><FileList><File ID='f1'>../Pages/Content.xml</File></FileList><DocRoot>../Document.xml</DocRoot></DocVersion>");
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Versions><Version ID='v1' Index='1' BaseLoc='Doc_0/Payload/original%2Ebin'/></Versions></DocBody></OFD>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "encoded or backslash BaseLoc");
    }

    [Fact]
    public async Task R16TypedMetadataAndStandardObjectsSurvivePhysicalAttachmentRemoval()
    {
        var entries = Entries(Seed());
        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocInfo><DocID>r18r16</DocID><DocUsage>EBook</DocUsage><CreatorVersion>1.0</CreatorVersion><Cover>Doc_0/Cover.png</Cover><CustomDatas><CustomData Name='case'>old</CustomData></CustomDatas></DocInfo><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "</CommonData><Permissions><Edit>true</Edit><Print Printable='true' Copies='2'/></Permissions><VPreferences><PageMode>UseOutlines</PageMode><HideToolbar>false</HideToolbar></VPreferences>", StringComparison.Ordinal));
        entries["Doc_0/Cover.png"] = [7, 8, 9];
        var package = await Read(Zip(entries));
        package.Options.Metadata.CustomDatas!.Items[0].Value = "new";
        package.Permissions!.Edit = false;
        package.ViewingPreferences!.PageMode = OfdPageMode.UseBookmarks;
        package.Actions ??= new OfdActionCollection();
        package.Actions.Items.Add(new OfdUriAction { Event = "CLICK", Uri = "https://example.invalid/a%20b" });
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        Assert.Equal(entries["Doc_0/Cover.png"], Entries(output)["Doc_0/Cover.png"]);
        var saved = await Read(output);
        Assert.Empty(saved.Attachments);
        Assert.Equal("new", saved.Options.Metadata.CustomDatas!.Items[0].Value);
        Assert.False(saved.Permissions!.Edit);
        Assert.Equal(OfdPageMode.UseBookmarks, saved.ViewingPreferences!.PageMode);
        Assert.Equal("https://example.invalid/a%20b", Assert.IsType<OfdUriAction>(Assert.Single(saved.Actions!.Items)).Uri);

        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocInfo><Cover>Doc_0/Payload/original.bin</Cover></DocInfo><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "remains referenced");

        entries["OFD.xml"] = Utf8($"<OFD xmlns='{Ns}'><DocBody><DocInfo><Cover>Doc_0/Payload/original%2Ebin</Cover></DocInfo><DocRoot>Doc_0/Document.xml</DocRoot></DocBody></OFD>");
        package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "encoded or backslash XML reference");
    }

    [Fact]
    public async Task OpaqueAttachmentActionCannotOutlivePhysicalTarget()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</Document>", "<Actions><Action Event='INVALID'><GotoA AttachID='7'/></Action></Actions></Document>", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.IsType<OfdRawAction>(Assert.Single(package.Actions!.Items));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "retained action targets the removed attachment ID");

        package.Actions.Items.Clear();
        var (output, _) = await Write(package);
        Assert.Empty((await Read(output)).Attachments);
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.DoesNotContain("GotoA", Encoding.UTF8.GetString(Entries(output)["Doc_0/Document.xml"]));
    }

    [Fact]
    public async Task StandardActionRegionSurvivesPhysicalAttachmentRemoval()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</Document>", "<Actions><Action Event='CLICK'><Region><Area Start='0 0'><Move Point1='0 0'/><Line Point1='10 10'/><Close/></Area></Region><URI URI='https://example.invalid'/></Action></Actions></Document>", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.IsType<OfdUriAction>(Assert.Single(package.Actions!.Items));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        var (output, _) = await Write(package);
        var document = XDocument.Parse(Encoding.UTF8.GetString(Entries(output)["Doc_0/Document.xml"]));
        Assert.Equal("0 0", document.Descendants(XName.Get("Region", Ns)).Single()
            .Element(XName.Get("Area", Ns))!.Attribute("Start")!.Value);
        Assert.Empty((await Read(output)).Attachments);
    }

    [Fact]
    public async Task NestedDocumentAttachmentDeclarationCannotSplitItsPath()
    {
        var entries = Entries(Seed());
        entries["Doc_0/Document.xml"] = Utf8(Encoding.UTF8.GetString(entries["Doc_0/Document.xml"])
            .Replace("</CommonData>", "<Attachments>Payload/orig<Title>inal.bin</Title></Attachments></CommonData>", StringComparison.Ordinal));
        var package = await Read(Zip(entries));
        Assert.True(package.RemoveAttachment(Assert.Single(package.Attachments), OfdAttachmentRemovalMode.Physical));
        await RefusesWithoutOutput(package, "mixed XML path content");
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
        using var signed = new MemoryStream();
        await new OfdSignatureService().SignAsync(new MemoryStream(Seed()), signed, new TestSignatureProvider(),
            new OfdSignatureOptions { CheckMethod = OfdDigestAlgorithms.Sha256Oid });
        var package = await Read(signed.ToArray());
        Assert.True(package.RemoveAttachment(package.Attachments[0], OfdAttachmentRemovalMode.Physical));
        var (output, result) = await Write(package);
        Assert.True(result.SignaturesInvalidated);
        Assert.Contains(result.Diagnostics, item => item.Contains("sign", StringComparison.OrdinalIgnoreCase));
        Assert.DoesNotContain("Doc_0/Payload/original.bin", Entries(output).Keys);
        Assert.DoesNotContain(Entries(output).Keys, path => path.EndsWith("/SignedValue.dat", StringComparison.OrdinalIgnoreCase));
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

    private static async Task RefusesWithoutOutput(OfdDocumentPackage package, string? expectedMessage = null)
    {
        var preserved = package.PreservedEntries.ToDictionary(entry => entry.Key, entry => entry.Value.ToArray());
        using var output = new MemoryStream();
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdPackageWriter().WriteWithResultAsync(package, output));
        if (expectedMessage is not null) Assert.Contains(expectedMessage, error.Message);
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

    private sealed class TestSignatureProvider : IOfdSignatureProvider
    {
        public string ProviderName => "R18 test";
        public string Company => "Ofdrw.Net";
        public string Version => "1";
        public string SignatureMethod => "urn:ofdrw-net:r18-test";
        public string SignatureType => "Sign";
        public byte[]? SealData => null;
        public Task<byte[]> SignAsync(byte[] signatureXml, string propertyInformation,
            CancellationToken cancellationToken = default) => Task.FromResult(new byte[] { 7, 8, 9 });
    }
}
