using System.Security.Cryptography;
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
        string? rootAction = null, string firstPageExtra = "", bool embeddedGlobalFile = true, bool mixedAnnotations = false)
    {
        var page = embeddedPageFile
            ? "<< /Type /Annot /Subtype /FileAttachment /Rect [20 30 40 50] /Contents (page note) /FS 20 0 R >>"
            : "<< /Type /Annot /Subtype /FileAttachment /Rect [20 30 40 50] /FS << /F (missing.bin) >> >>";
        return fixture.Create(new[] { fixture.Text("A中文", 30, 500), fixture.Text("B中文", 30, 500) },
            catalogEntries: "/Outlines 13 0 R /Names << /EmbeddedFiles 16 0 R >>",
            pageEntriesByIndex: new[] { "/Annots [19 0 R" + (mixedAnnotations ? " 22 0 R" : "") + "] " + firstPageExtra, "" },
            trailingObjects: new[]
            {
                "<< /Type /Outlines /First 14 0 R /Last 14 0 R /Count 2 >>",
                "<< /Title (Chapter A) /Parent 13 0 R /First 15 0 R /Last 15 0 R /Count 1 " +
                    (rootAction ?? "/Dest [" + destination + "]") + " >>",
                "<< /Title (Child) /Parent 14 0 R /Dest [9 0 R /XYZ 30 500 null] >>",
                "<< /Names [(same.bin) 17 0 R] >>",
                embeddedGlobalFile ? "<< /Type /Filespec /F (same.bin) /EF << /F 18 0 R >> >>" :
                    "<< /Type /Filespec /F (same.bin) >>",
                "<< /Type /EmbeddedFile /Length 6 >>\nstream\nGLOBAL\nendstream",
                page,
                "<< /Type /Filespec /F (same.bin) /EF << /F 21 0 R >> >>",
                "<< /Type /EmbeddedFile /Length 4 >>\nstream\nPAGE\nendstream"
            }.Concat(mixedAnnotations ? new[] { "<< /Type /Annot /Subtype /Link /Rect [60 60 80 80] /A << /S /URI /URI (https://example.invalid) >> >>" } :
                Array.Empty<string>()).ToArray());
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
        Assert.Contains("GotoA", Assert.IsType<OfdPathElement>(annotation.Appearance.Elements.Single()).SourceXml);
        Assert.Null(package.Pages[2].Annotations);
        Assert.Equal(package.Attachments[2].Id, Assert.Single(package.Pages[3].Annotations!.Items).AttachmentId);
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

    [Fact]
    public async Task RotatedPageAttachmentUsesDisplayedCropCoordinates()
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, firstPageExtra: "/CropBox [10 20 410 570] /Rotate 90"));
        using var output = new MemoryStream();
        await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = new PdfToOfdOptions
            { CopyBookmarks = true, CopyAttachments = true } }).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var box = Assert.Single(package.Pages[0].Annotations!.Items).Appearance;
        Assert.Equal(10 * 25.4 / 72, box.XMillimeters, 5);
        Assert.Equal(370 * 25.4 / 72, box.YMillimeters, 5);
        Assert.Equal(20 * 25.4 / 72, box.WidthMillimeters, 5);
        Assert.Equal(20 * 25.4 / 72, box.HeightMillimeters, 5);
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
}
