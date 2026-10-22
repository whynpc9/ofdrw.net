using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class PdfMetadataMigrationTests
{
    private static byte[] Font() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));

    private static byte[] Source(PdfFixture fixture, string destination = "11 0 R /Fit", bool embeddedPageFile = true,
        string? rootAction = null, string firstPageExtra = "", bool embeddedGlobalFile = true, bool mixedAnnotations = false,
        int extraLinkAnnotations = 0, string? globalStreamObject = null, string childZoom = "null")
    {
        var linkCount = extraLinkAnnotations + (mixedAnnotations ? 1 : 0);
        var page = embeddedPageFile
            ? "<< /Type /Annot /Subtype /FileAttachment /Rect [20 30 40 50] /Contents (page note) /FS 20 0 R >>"
            : "<< /Type /Annot /Subtype /FileAttachment /Rect [20 30 40 50] /FS << /F (missing.bin) >> >>";
        return fixture.Create(new[] { fixture.Text("A中文", 30, 500), fixture.Text("B中文", 30, 500) },
            catalogEntries: "/Outlines 13 0 R /Names << /EmbeddedFiles 16 0 R >>",
            pageEntriesByIndex: new[] { "/Annots [19 0 R" + string.Concat(Enumerable.Range(0, linkCount).Select(index => $" {22 + index} 0 R")) + "] " + firstPageExtra, "" },
            trailingObjects: new[]
            {
                "<< /Type /Outlines /First 14 0 R /Last 14 0 R /Count 2 >>",
                "<< /Title (Chapter A) /Parent 13 0 R /First 15 0 R /Last 15 0 R /Count 1 " +
                    (rootAction ?? "/Dest [" + destination + "]") + " >>",
                "<< /Title (Child) /Parent 14 0 R /Dest [9 0 R /XYZ 30 500 " + childZoom + "] >>",
                "<< /Names [(same.bin) 17 0 R] >>",
                embeddedGlobalFile ? "<< /Type /Filespec /F (same.bin) /EF << /F 18 0 R >> >>" :
                    "<< /Type /Filespec /F (same.bin) >>",
                globalStreamObject ?? "<< /Type /EmbeddedFile /Length 6 >>\nstream\nGLOBAL\nendstream",
                page,
                "<< /Type /Filespec /F (same.bin) /EF << /F 21 0 R >> >>",
                "<< /Type /EmbeddedFile /Length 4 >>\nstream\nPAGE\nendstream"
            }.Concat(Enumerable.Repeat("<< /Type /Annot /Subtype /Link /Rect [60 60 80 80] /A << /S /URI /URI (https://example.invalid) >> >>", linkCount)).ToArray());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BothPdfImportersCopyClosedOutlineAndEmbeddedFiles(bool vector)
    {
        using var fixture = new PdfFixture(Font()); var pdf = Source(fixture);
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        if (vector)
        {
            var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertWithResultAsync(input, output, new[] { 1, 0, 1, 0 });
            Assert.All(result.Pages, page => Assert.True(page.IsNative));
        }
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 1, 0, 1, 0 });

        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(4, package.Pages.Count);
        var root = Assert.Single(package.Outlines!.Items);
        Assert.Equal("Chapter A", root.Title);
        Assert.Same(package.Pages[0], Assert.IsType<OfdGoToAction>(Assert.Single(root.Actions!.Items)).Destination!.Page);
        var child = Assert.Single(root.Children);
        Assert.Equal("Child", child.Title);
        Assert.Same(package.Pages[1], Assert.IsType<OfdGoToAction>(Assert.Single(child.Actions!.Items)).Destination!.Page);
        Assert.Equal(new[] { "same.bin", "same.bin (2)", "same.bin (3)" }, package.Attachments.Select(item => item.Name));
        Assert.Equal(SHA256.HashData("GLOBAL"u8.ToArray()), SHA256.HashData(package.Attachments[0].Data));
        Assert.Equal(SHA256.HashData("PAGE"u8.ToArray()), SHA256.HashData(package.Attachments[1].Data));
        Assert.Equal(SHA256.HashData("PAGE"u8.ToArray()), SHA256.HashData(package.Attachments[2].Data));
        Assert.Null(package.Pages[0].Annotations);
        var annotation = Assert.Single(package.Pages[1].Annotations!.Items);
        Assert.Equal(OfdAnnotationType.Link, annotation.Type);
        Assert.Equal("page note", annotation.Remark);
        Assert.Equal(package.Attachments[1].Id, annotation.AttachmentId);
        var pin = Assert.IsType<OfdPathElement>(annotation.Appearance.Elements.Single());
        Assert.Contains("GotoA", pin.SourceXml);
        Assert.Equal(vector, pin.Fill);
        Assert.Null(package.Pages[2].Annotations);
        Assert.Equal(package.Attachments[2].Id, Assert.Single(package.Pages[3].Annotations!.Items).AttachmentId);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BothPdfImportersRetainExplicitXyzZoom(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination: "11 0 R /XYZ 30 500 1.5", childZoom: "2"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector) await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options }).ConvertAsync(input, output);
        else await new PdfToOfdConverter(options).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var root = Assert.Single(package.Outlines!.Items);
        Assert.Equal(1.5, Assert.IsType<OfdGoToAction>(Assert.Single(root.Actions!.Items)).Destination!.Zoom);
        var child = Assert.Single(root.Children);
        var xyz = Assert.IsType<OfdGoToAction>(Assert.Single(child.Actions!.Items)).Destination!;
        Assert.Equal(OfdDestinationType.XYZ, xyz.Type);
        Assert.Equal(2, xyz.Zoom);
    }

    [Fact]
    public async Task UnrepresentableXyzZoomFailsBeforeOutput()
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, childZoom: "100"));
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter(new PdfToOfdOptions
            { CopyBookmarks = true, PreferExternalPdfToPpm = true }).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task DisabledSwitchesLeaveNoNavigationOrAttachmentReferences()
    {
        using var fixture = new PdfFixture(Font()); using var input = new MemoryStream(Source(fixture)); using var output = new MemoryStream();
        await new PdfToOfdConverter(new PdfToOfdOptions { PreferExternalPdfToPpm = true }).ConvertAsync(input, output);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Null(package.Outlines); Assert.Null(package.Bookmarks); Assert.Empty(package.Attachments);
        Assert.All(package.Pages, page => Assert.Null(page.Annotations));
    }

    [Theory]
    [InlineData(true, "99 0 R /Fit")]
    [InlineData(false, "11 0 R /Fit")]
    public async Task InvalidDestinationOrMissingPageFileFailsBeforeOutput(bool invalidDestination, string destination)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination, embeddedPageFile: invalidDestination));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        await Assert.ThrowsAnyAsync<Exception>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task ExternalOutlineActionFailsBeforeOutput()
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, rootAction: "/A << /S /URI /URI (https://example.invalid) >>"));
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter(new PdfToOfdOptions
            { CopyBookmarks = true, PreferExternalPdfToPpm = true }).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task GroupingOutlineRetainsItsChildAndSelectedPageTarget(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, rootAction: ""));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertAsync(input, output, new[] { 1, 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 1, 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var group = Assert.Single(package.Outlines!.Items);
        Assert.Equal("Chapter A", group.Title);
        Assert.Null(group.Actions);
        var child = Assert.Single(group.Children);
        Assert.Equal("Child", child.Title);
        Assert.Same(package.Pages[1], Assert.IsType<OfdGoToAction>(Assert.Single(child.Actions!.Items)).Destination!.Page);
    }

    [Fact]
    public async Task MissingGlobalEmbeddedPayloadFailsBeforeOutput()
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, embeddedGlobalFile: false));
        using var output = new MemoryStream();
        await Assert.ThrowsAnyAsync<Exception>(() => new PdfToOfdConverter(new PdfToOfdOptions
            { CopyAttachments = true, PreferExternalPdfToPpm = true }).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(90, 10, 10, 480, 20, false)]
    [InlineData(90, 10, 10, 480, 20, true)]
    [InlineData(270, 520, 370, 70, 380, false)]
    [InlineData(270, 520, 370, 70, 380, true)]
    public async Task RotatedPageMetadataUsesDisplayedCropCoordinates(int rotation, double pinX, double pinY, double targetX, double targetY, bool vector)
    {
        using var fixture = new PdfFixture(Font());
        var destination = rotation == 90 ? "9 0 R /FitR 20 30 40 50" : "9 0 R /FitV 20";
        using var input = new MemoryStream(Source(fixture, destination: destination,
            firstPageExtra: $"/CropBox [10 20 410 570] /Rotate {rotation}"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector) await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options }).ConvertAsync(input, output);
        else await new PdfToOfdConverter(options).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var annotation = Assert.Single(package.Pages[0].Annotations!.Items);
        var box = annotation.Appearance;
        Assert.Equal(vector, Assert.IsType<OfdPathElement>(annotation.Appearance.Elements.Single()).Fill);
        Assert.Equal(pinX * 25.4 / 72, box.XMillimeters, 5);
        Assert.Equal(pinY * 25.4 / 72, box.YMillimeters, 5);
        Assert.Equal(20 * 25.4 / 72, box.WidthMillimeters, 5);
        Assert.Equal(20 * 25.4 / 72, box.HeightMillimeters, 5);
        var root = Assert.Single(package.Outlines!.Items);
        var destinationAction = Assert.IsType<OfdGoToAction>(Assert.Single(root.Actions!.Items));
        Assert.Same(package.Pages[0], destinationAction.Destination!.Page);
        if (rotation == 90)
        {
            Assert.Equal(OfdDestinationType.FitR, destinationAction.Destination.Type);
            Assert.Equal(pinX * 25.4 / 72, destinationAction.Destination.Left!.Value, 5);
            Assert.Equal(pinY * 25.4 / 72, destinationAction.Destination.Top!.Value, 5);
        }
        else
        {
            Assert.Equal(OfdDestinationType.FitH, destinationAction.Destination.Type);
            Assert.Equal(390 * 25.4 / 72, destinationAction.Destination.Top!.Value, 5);
        }
        var child = Assert.Single(root.Children);
        var xyz = Assert.IsType<OfdGoToAction>(Assert.Single(child.Actions!.Items)).Destination!;
        Assert.Equal(OfdDestinationType.XYZ, xyz.Type);
        Assert.Equal(targetX * 25.4 / 72, xyz.Left!.Value, 5);
        Assert.Equal(targetY * 25.4 / 72, xyz.Top!.Value, 5);
    }

    [Fact]
    public async Task MixedAnnotationsUseWholePageFallbackWhileRetainingFileAction()
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, mixedAnnotations: true));
        using var output = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
        {
            UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            Compatibility = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true }
        }).ConvertWithResultAsync(input, output, new[] { 0 });
        Assert.False(Assert.Single(result.Pages).IsNative);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Attachments.Count);
        Assert.Equal(package.Attachments[1].Id, Assert.Single(package.Pages[0].Annotations!.Items).AttachmentId);
    }

    [Fact]
    public async Task OrdinaryPageAnnotationsDoNotConsumeAttachmentSlots()
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, extraLinkAnnotations: 3));
        using var output = new MemoryStream();
        await new PdfToOfdConverter(new PdfToOfdOptions
            { CopyAttachments = true, MaxAttachmentCount = 2, PreferExternalPdfToPpm = true })
            .ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Attachments.Count);
        Assert.Equal(package.Attachments[1].Id, Assert.Single(package.Pages[0].Annotations!.Items).AttachmentId);
    }

    [Fact]
    public async Task CompressedAttachmentMayExceedItsDecodedByteLimit()
    {
        using var compressed = new MemoryStream();
        using (var deflater = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            deflater.WriteByte((byte)'A');
        var encoded = compressed.ToArray();
        Assert.True(encoded.Length > 4);
        var placeholder = new string('X', encoded.Length);
        using var fixture = new PdfFixture(Font());
        var pdf = Source(fixture, globalStreamObject: $"<< /Type /EmbeddedFile /Filter /FlateDecode /Length {encoded.Length} >>\nstream\n{placeholder}\nendstream");
        var marker = Encoding.ASCII.GetBytes("stream\n" + placeholder + "\nendstream");
        var position = pdf.AsSpan().IndexOf(marker);
        Assert.True(position >= 0);
        encoded.AsSpan().CopyTo(pdf.AsSpan(position + "stream\n".Length));
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        await new PdfToOfdConverter(new PdfToOfdOptions
            { CopyAttachments = true, MaxAttachmentBytes = 4, MaxTotalAttachmentBytes = 5, PreferExternalPdfToPpm = true })
            .ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(new byte[] { (byte)'A' }, package.Attachments[0].Data);
        Assert.Equal("PAGE"u8.ToArray(), package.Attachments[1].Data);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("duplicate")]
    [InlineData("external")]
    public async Task ReaderLeavesUnclosedAttachmentActionsUntyped(string defect)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture)); using var output = new MemoryStream();
        await new PdfToOfdConverter(new PdfToOfdOptions
            { CopyAttachments = true, PreferExternalPdfToPpm = true }).ConvertAsync(input, output, new[] { 0 });
        using var bytes = new MemoryStream();
        bytes.Write(output.ToArray());
        bytes.Position = 0;
        using (var archive = new ZipArchive(bytes, ZipArchiveMode.Update, leaveOpen: true))
        {
            var annotationPath = archive.Entries.Single(entry => entry.FullName.Contains("/Annots/Page_")).FullName;
            var annotation = ReadXml(annotationPath);
            var action = annotation.Descendants().Single(node => node.Name.LocalName == "GotoA");
            var target = action.Attribute("AttachID")!.Value;
            if (defect == "missing")
            {
                action.SetAttributeValue("AttachID", "9999999");
                ReplaceXml(annotationPath, annotation);
            }
            else
            {
                var attachmentPath = archive.Entries.Single(entry => entry.FullName.EndsWith("/Attachments.xml")).FullName;
                var attachments = ReadXml(attachmentPath);
                var item = attachments.Descendants().Single(node => node.Name.LocalName == "Attachment" && node.Attribute("ID")?.Value == target);
                if (defect == "duplicate") item.AddAfterSelf(new XElement(item));
                else item.SetAttributeValue("External", "true");
                ReplaceXml(attachmentPath, attachments);
            }
            XDocument ReadXml(string path)
            {
                using var stream = archive.GetEntry(path)!.Open();
                return XDocument.Load(stream);
            }
            void ReplaceXml(string path, XDocument xml)
            {
                archive.GetEntry(path)!.Delete();
                using var stream = archive.CreateEntry(path).Open();
                var data = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                stream.Write(data);
            }
        }
        bytes.Position = 0;
        var loaded = await new OfdReader().ReadAsync(bytes);
        Assert.Null(loaded.Pages[0].Annotations);
    }
}
