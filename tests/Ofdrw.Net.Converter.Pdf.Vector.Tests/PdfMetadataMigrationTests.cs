using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Pdf.Vector.Tests;

public sealed class PdfMetadataMigrationTests
{
    private static byte[] Font() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "regular.ttf"));

    [Theory]
    [InlineData(false, false, true, false)]
    [InlineData(false, false, false, true)]
    [InlineData(false, false, true, true)]
    [InlineData(true, false, true, false)]
    [InlineData(true, false, false, true)]
    [InlineData(true, false, true, true)]
    [InlineData(false, true, true, false)]
    [InlineData(false, true, false, true)]
    [InlineData(false, true, true, true)]
    public async Task AcceptedIndirectOrTypelessPageKeepsMetadata(bool vector, bool typeless,
        bool copyBookmarks, bool copyAttachments)
    {
        using var fixture = new PdfFixture(Font());
        var pdf = typeless
            ? Source(fixture, pageTypeByIndex: new[] { "", "/Page" })
            : Source(fixture, extraObjects: new[] { "/Page" }, pageTypeByIndex: new[] { "22 0 R", "/Page" });
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf,
            new UglyToad.PdfPig.ParsingOptions { UseLenientParsing = !vector }))
        {
            Assert.Equal(2, parsed.NumberOfPages);
            var rawPages = new PdfImportMetadataMigrator.RawPageLookup(parsed, new PdfToOfdOptions());
            Assert.Equal(parsed.GetPage(1).Dictionary, rawPages.Get(0, CancellationToken.None));
        }
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = copyBookmarks, CopyAttachments = copyAttachments,
            PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
                .ConvertAsync(input, output, new[] { 1, 0, 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 1, 0, 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(3, package.Pages.Count);
        Assert.Equal(copyAttachments ? 3 : 0, package.Attachments.Count);
        if (copyAttachments)
        {
            Assert.Null(package.Pages[0].Annotations);
            Assert.Equal(package.Attachments[1].Id, Assert.Single(package.Pages[1].Annotations!.Items).AttachmentId);
            Assert.Equal(package.Attachments[2].Id, Assert.Single(package.Pages[2].Annotations!.Items).AttachmentId);
        }
        else Assert.All(package.Pages, page => Assert.Null(page.Annotations));
        if (copyBookmarks)
        {
            var root = Assert.Single(package.Outlines!.Items);
            Assert.Same(package.Pages[0], Assert.IsType<OfdGoToAction>(Assert.Single(root.Actions!.Items)).Destination!.Page);
            Assert.Same(package.Pages[1], Assert.IsType<OfdGoToAction>(Assert.Single(Assert.Single(root.Children).Actions!.Items)).Destination!.Page);
        }
        else Assert.Null(package.Outlines);
    }

    [Fact]
    public void StrictPdfPigRejectsTypelessPageBeforeVectorMetadata()
    {
        using var fixture = new PdfFixture(Font());
        var pdf = Source(fixture, pageTypeByIndex: new[] { "", "/Page" });
        Assert.Throws<UglyToad.PdfPig.Core.PdfDocumentFormatException>(() =>
            UglyToad.PdfPig.PdfDocument.Open(pdf,
                new UglyToad.PdfPig.ParsingOptions { UseLenientParsing = false }));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task MetadataRejectsPageLeafWithContradictoryKids(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        var pdf = Source(fixture, firstPageExtra: "/Kids []");
        using (var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf,
            new UglyToad.PdfPig.ParsingOptions { UseLenientParsing = !vector }))
            Assert.Equal(2, parsed.NumberOfPages);
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output, new[] { 1, 0 }));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options)
            .ConvertAsync(input, output, new[] { 1, 0 }));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public void RepeatedLatePageLookupVisitsRawPageTreeOnce()
    {
        using var fixture = new PdfFixture(Font());
        var pdf = fixture.Create(Enumerable.Repeat(string.Empty, 32).ToArray());
        using var source = UglyToad.PdfPig.PdfDocument.Open(pdf);
        var options = new PdfToOfdOptions { CopyAttachments = true };
        var pages = new PdfImportMetadataMigrator.RawPageLookup(source, options);
        foreach (var index in new[] { 31, 0, 31, 12, 31, 12 })
            Assert.True(PdfImportMetadataMigrator.HasOnlyFileAttachments(source, index, options, pages, CancellationToken.None));
        var last = pages.Get(31, CancellationToken.None);
        Assert.Equal(33, pages.VisitedNodes); // one Pages node and 32 Page nodes
        Assert.Same(last, pages.Get(31, CancellationToken.None));
        Assert.Same(last, pages.Get(31, CancellationToken.None));
        Assert.NotSame(last, pages.Get(0, CancellationToken.None));
        Assert.Equal(33, pages.VisitedNodes);
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => pages.Get(31, cancelled.Token));
        using var otherDocument = UglyToad.PdfPig.PdfDocument.Open(pdf);
        Assert.Throws<InvalidOperationException>(() => PdfImportMetadataMigrator.HasOnlyFileAttachments(
            otherDocument, 31, options, pages, CancellationToken.None));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task DuplicatePageAttachmentsStillChargeEachOutputOccurrence(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture)); using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, MaxAttachmentCount = 2, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output, new[] { 0, 0 }));
        else
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output,
                new[] { 0, 0 }));
    }

    private static byte[] Source(PdfFixture fixture, string destination = "11 0 R /Fit", bool embeddedPageFile = true,
        string? rootAction = null, string firstPageExtra = "", bool embeddedGlobalFile = true, bool mixedAnnotations = false,
        int extraLinkAnnotations = 0, string? globalStreamObject = null, string childZoom = "null", string? outlineRoot = null,
        string attachmentName = "same.bin", int? attachmentFlags = 4, string? rootOutlineItem = null,
        string annotationRect = "20 30 40 50", string[]? extraObjects = null, string? globalFileSpec = null,
        string? nameTreeKeyToken = null, string? pageFileSpec = null, string pageSubtype = "/FileAttachment",
        IReadOnlyList<string>? pageTypeByIndex = null)
    {
        var linkCount = extraLinkAnnotations + (mixedAnnotations ? 1 : 0);
        var flags = attachmentFlags is int value ? " /F " + value : string.Empty;
        var page = embeddedPageFile
            ? "<< /Type /Annot /Subtype " + pageSubtype + " /Rect [" + annotationRect + "] /Contents (page note) /FS 20 0 R" + flags + " >>"
            : "<< /Type /Annot /Subtype /FileAttachment /Rect [" + annotationRect + "] /FS << /F (missing.bin) >>" + flags + " >>";
        return fixture.Create(new[] { fixture.Text("A中文", 30, 500), fixture.Text("B中文", 30, 500) },
            catalogEntries: "/Outlines 13 0 R /Names << /EmbeddedFiles 16 0 R >>",
            pageEntriesByIndex: new[] { "/Annots [19 0 R" + string.Concat(Enumerable.Range(0, linkCount).Select(index => $" {22 + index} 0 R")) + "] " + firstPageExtra, "" },
            trailingObjects: new[]
            {
                outlineRoot ?? "<< /Type /Outlines /First 14 0 R /Last 14 0 R /Count 2 >>",
                rootOutlineItem ?? "<< /Title (Chapter A) /Parent 13 0 R /First 15 0 R /Last 15 0 R /Count 1 " +
                    (rootAction ?? "/Dest [" + destination + "]") + " >>",
                "<< /Title (Child) /Parent 14 0 R /Dest [9 0 R /XYZ 30 500 " + childZoom + "] >>",
                "<< /Names [" + (nameTreeKeyToken ?? "(" + attachmentName + ")") + " 17 0 R] >>",
                embeddedGlobalFile ? globalFileSpec ?? "<< /Type /Filespec /F (" + attachmentName + ") /EF << /F 18 0 R >> >>" :
                    "<< /Type /Filespec /F (" + attachmentName + ") >>",
                globalStreamObject ?? "<< /Type /EmbeddedFile /Length 6 >>\nstream\nGLOBAL\nendstream",
                page,
                pageFileSpec ?? "<< /Type /Filespec /F (" + attachmentName + ") /EF << /F 21 0 R >> >>",
                "<< /Type /EmbeddedFile /Length 4 >>\nstream\nPAGE\nendstream"
            }.Concat(Enumerable.Repeat("<< /Type /Annot /Subtype /Link /Rect [60 60 80 80] /A << /S /URI /URI (https://example.invalid) >> >>", linkCount))
            .Concat(extraObjects ?? Array.Empty<string>()).ToArray(), pageTypeByIndex: pageTypeByIndex);
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
        Assert.True(annotation.Visible);
        Assert.True(annotation.Print);
        Assert.Equal("page note", annotation.Remark);
        Assert.Equal(package.Attachments[1].Id, annotation.AttachmentId);
        var pin = Assert.IsType<OfdPathElement>(annotation.Appearance.Elements.Single());
        Assert.Contains("GotoA", pin.SourceXml);
        Assert.Equal(vector, pin.Fill);
        Assert.Null(package.Pages[2].Annotations);
        Assert.Equal(package.Attachments[2].Id, Assert.Single(package.Pages[3].Annotations!.Items).AttachmentId);

        // The reader-owned attachment action is metadata on a visible pin, not
        // an unresolved drawing resource in either renderer's shared preflight.
        using var renderedPdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(new MemoryStream(output.ToArray()), renderedPdf);
        Assert.True(renderedPdf.Length > 0);
        using var renderedSvg = new MemoryStream();
        await new OfdToSvgConverter().ConvertAsync(new MemoryStream(output.ToArray()), renderedSvg, pageIndex: 1);
        Assert.Contains("<path", Encoding.UTF8.GetString(renderedSvg.ToArray()));
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

    [Theory]
    [InlineData(false, "11 0 R /XYZ 30 500", true)]
    [InlineData(true, "11 0 R /XYZ 30 500", true)]
    [InlineData(false, "11 0 R /XYZ null null", false)]
    [InlineData(true, "11 0 R /XYZ null null", false)]
    [InlineData(false, "11 0 R /XYZ", false)]
    [InlineData(true, "11 0 R /XYZ", false)]
    public async Task MissingTrailingXyzOperandsRetainReaderValues(bool vector, string rawDestination, bool positioned)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination: rawDestination));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options }).ConvertAsync(input, output);
        else await new PdfToOfdConverter(options).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var target = Assert.IsType<OfdGoToAction>(Assert.Single(Assert.Single(package.Outlines!.Items).Actions!.Items)).Destination!;
        Assert.Equal(OfdDestinationType.XYZ, target.Type);
        Assert.Equal(positioned, target.Left.HasValue && target.Top.HasValue);
        Assert.Null(target.Zoom);
    }

    [Theory]
    [InlineData(false, "11 0 R /XYZ 30 500 null", "11 0 R /XYZ 22 0 R 23 0 R null")]
    [InlineData(true, "11 0 R /XYZ 30 500 null", "11 0 R /XYZ 22 0 R 23 0 R null")]
    [InlineData(false, "11 0 R /FitH 500", "11 0 R /FitH 23 0 R")]
    [InlineData(true, "11 0 R /FitH 500", "11 0 R /FitH 23 0 R")]
    [InlineData(false, "11 0 R /FitV 30", "11 0 R /FitV 22 0 R")]
    [InlineData(true, "11 0 R /FitV 30", "11 0 R /FitV 22 0 R")]
    [InlineData(false, "11 0 R /FitR 30 500 40 550", "11 0 R /FitR 22 0 R 23 0 R 24 0 R 25 0 R")]
    [InlineData(true, "11 0 R /FitR 30 500 40 550", "11 0 R /FitR 22 0 R 23 0 R 24 0 R 25 0 R")]
    public async Task IndirectBookmarkCoordinatesMatchDirectValues(bool vector, string direct, string indirect)
    {
        using var fixture = new PdfFixture(Font());
        async Task<OfdDestination> Convert(string destination)
        {
            using var input = new MemoryStream(Source(fixture, destination: destination, extraObjects: new[] { "30", "500", "40", "550" }));
            using var output = new MemoryStream();
            var options = new PdfToOfdOptions { CopyBookmarks = true, PreferExternalPdfToPpm = true };
            if (vector) await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
                .ConvertAsync(input, output);
            else await new PdfToOfdConverter(options).ConvertAsync(input, output);
            output.Position = 0;
            var package = await new OfdReader().ReadAsync(output);
            return Assert.IsType<OfdGoToAction>(Assert.Single(Assert.Single(package.Outlines!.Items).Actions!.Items)).Destination!;
        }
        var expected = await Convert(direct);
        var actual = await Convert(indirect);
        Assert.Equal(expected.Type, actual.Type);
        Assert.Equal(expected.Left, actual.Left);
        Assert.Equal(expected.Top, actual.Top);
        Assert.Equal(expected.Right, actual.Right);
        Assert.Equal(expected.Bottom, actual.Bottom);
        Assert.True(expected.Left.HasValue || expected.Top.HasValue);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NonnumericIndirectBookmarkCoordinateFailsBeforeOutput(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination: "11 0 R /FitV 22 0 R", extraObjects: new[] { "(bad)" }));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, PreferExternalPdfToPpm = true };
        if (vector) await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
            { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
            .ConvertAsync(input, output));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task BookmarkOnlyCropBoxRasterAlignsWithOutputPage(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, firstPageExtra: "/CropBox [10 20 410 570] /Rotate 90"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, PreferExternalPdfToPpm = true };
        if (vector)
        {
            var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
                .ConvertWithResultAsync(input, output, new[] { 0, 1 });
            Assert.False(result.Pages[0].IsNative);
        }
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0, 1 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Null(package.Pages[0].Annotations);
        Assert.Empty(package.Attachments);
        var image = Assert.Single(package.Pages[0].Elements.OfType<OfdImageElement>());
        Assert.Equal(1100, System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(image.Data.AsSpan(16, 4)));
        Assert.Equal(800, System.Buffers.Binary.BinaryPrimitives.ReadInt32BigEndian(image.Data.AsSpan(20, 4)));
        var child = Assert.Single(Assert.Single(package.Outlines!.Items).Children);
        var destination = Assert.IsType<OfdGoToAction>(Assert.Single(child.Actions!.Items)).Destination!;
        Assert.Same(package.Pages[0], destination.Page);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task UnknownExplicitOutlineDestinationFailsBeforeOutput(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination: "11 0 R /UnsupportedDestination"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output));
        else await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false, "/FitH", OfdDestinationType.FitV)]
    [InlineData(true, "/FitH", OfdDestinationType.FitV)]
    [InlineData(false, "/FitV", OfdDestinationType.FitH)]
    [InlineData(true, "/FitV", OfdDestinationType.FitH)]
    public async Task NullFitCoordinateRetainsReaderPositionAfterRotation(bool vector, string fit, OfdDestinationType expected)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination: $"9 0 R {fit} null",
            firstPageExtra: "/CropBox [10 20 410 570] /Rotate 90"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
                .ConvertAsync(input, output);
        else await new PdfToOfdConverter(options).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var root = Assert.Single(package.Outlines!.Items);
        var destination = Assert.IsType<OfdGoToAction>(Assert.Single(root.Actions!.Items)).Destination!;
        Assert.Same(package.Pages[0], destination.Page);
        Assert.Equal(expected, destination.Type);
        Assert.Null(destination.Left);
        Assert.Null(destination.Top);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task OutlineRootCountWithoutFirstFailsBeforeOutput(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, outlineRoot: "<< /Type /Outlines /Count 1 >>"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
                .ConvertAsync(input, output));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false, "missing-root-last")]
    [InlineData(true, "missing-root-last")]
    [InlineData(false, "unreachable-root-last")]
    [InlineData(true, "unreachable-root-last")]
    [InlineData(false, "missing-child-last")]
    [InlineData(true, "missing-child-last")]
    [InlineData(false, "unreachable-child-last")]
    [InlineData(true, "unreachable-child-last")]
    public async Task OutlineSiblingChainMustReachDeclaredLastBeforeOutput(bool vector, string defect)
    {
        using var fixture = new PdfFixture(Font());
        var root = defect switch
        {
            "missing-root-last" => "<< /Type /Outlines /First 14 0 R /Count 2 >>",
            "unreachable-root-last" => "<< /Type /Outlines /First 14 0 R /Last 15 0 R /Count 2 >>",
            _ => null
        };
        var rootItem = defect switch
        {
            "missing-child-last" => "<< /Title (Chapter A) /Parent 13 0 R /First 15 0 R /Count 1 /Dest [11 0 R /Fit] >>",
            "unreachable-child-last" => "<< /Title (Chapter A) /Parent 13 0 R /First 15 0 R /Last 14 0 R /Count 1 /Dest [11 0 R /Fit] >>",
            _ => null
        };
        using var input = new MemoryStream(Source(fixture, outlineRoot: root, rootOutlineItem: rootItem));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options, UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage })
                .ConvertAsync(input, output));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false, "11 0 R /FitB")]
    [InlineData(true, "11 0 R /FitB")]
    [InlineData(false, "11 0 R /FitBH 500")]
    [InlineData(true, "11 0 R /FitBH 500")]
    [InlineData(false, "11 0 R /FitBV 30")]
    [InlineData(true, "11 0 R /FitBV 30")]
    public async Task BoundingBoxFitFailsBeforeOutput(bool vector, string destination)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, destination: destination));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output));
        else await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
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
    public async Task ChainedLocalGoToFailsBeforeOutput(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture,
            rootAction: "/A << /S /GoTo /D [11 0 R /Fit] /Next << /S /URI /URI (https://example.invalid) >> >>"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyBookmarks = true, CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output));
        else await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task LongDuplicateAttachmentNamesFitWithSuffixAndExtension(bool vector)
    {
        var name = new string('a', 251) + ".bin";
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, attachmentName: name));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options }).ConvertAsync(input, output);
        else await new PdfToOfdConverter(options).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Attachments.Count);
        Assert.Equal(name, package.Attachments[0].Name);
        Assert.Equal(new string('a', 247) + " (2).bin", package.Attachments[1].Name);
        Assert.All(package.Attachments, attachment => Assert.True(attachment.Name.Length <= 255));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task CollisionTruncationKeepsSurrogatePairWhole(bool vector)
    {
        var name = new string('a', 246) + "😀bbb.bin";
        Assert.Equal(255, name.Length);
        var hexName = "<" + Convert.ToHexString(Encoding.BigEndianUnicode.GetBytes("\uFEFF" + name)) + ">";
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, nameTreeKeyToken: hexName,
            pageFileSpec: $"<< /Type /Filespec /UF {hexName} /EF << /F 21 0 R >> >>"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertAsync(input, output, new[] { 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(name, package.Attachments[0].Name);
        Assert.Equal(new string('a', 246) + " (2).bin", package.Attachments[1].Name);
        Assert.All(package.Attachments, attachment => Assert.True(attachment.Name.Length <= 255));
        Assert.Equal(SHA256.HashData("GLOBAL"u8.ToArray()), SHA256.HashData(package.Attachments[0].Data));
        Assert.Equal(SHA256.HashData("PAGE"u8.ToArray()), SHA256.HashData(package.Attachments[1].Data));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task IndirectFileAttachmentSubtypeIsCopied(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, pageSubtype: "22 0 R", extraObjects: new[] { "/FileAttachment" }));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertAsync(input, output, new[] { 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("PAGE"u8.ToArray(), package.Attachments[1].Data);
        Assert.Equal(package.Attachments[1].Id, Assert.Single(package.Pages[0].Annotations!.Items).AttachmentId);
    }

    [Theory]
    [InlineData(false, "(invalid)")]
    [InlineData(true, "(invalid)")]
    [InlineData(false, "22 0 R")]
    [InlineData(true, "22 0 R")]
    public async Task InvalidOrCyclicAnnotationSubtypeFailsBeforeOutput(bool vector, string target)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, pageSubtype: "22 0 R", extraObjects: new[] { target }));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output, new[] { 0 }));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 }));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData(true, null)]
    [InlineData(false, 0)]
    [InlineData(true, 0)]
    [InlineData(false, 2)]
    [InlineData(true, 2)]
    [InlineData(false, 32)]
    [InlineData(true, 32)]
    [InlineData(false, 36)]
    [InlineData(true, 36)]
    public async Task NonprintableOrHiddenAttachmentFailsBeforeOutput(bool vector, int? flags)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, attachmentFlags: flags));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<NotSupportedException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output));
        else await Assert.ThrowsAsync<NotSupportedException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ReversedFileAttachmentCornersRetainPinGeometry(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        async Task<OfdAnnotation> Convert(string rect)
        {
            using var input = new MemoryStream(Source(fixture, annotationRect: rect,
                firstPageExtra: "/CropBox [10 20 410 570] /Rotate 90"));
            using var output = new MemoryStream();
            var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
            if (vector) await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options }).ConvertAsync(input, output);
            else await new PdfToOfdConverter(options).ConvertAsync(input, output);
            output.Position = 0;
            var package = await new OfdReader().ReadAsync(output);
            return Assert.Single(package.Pages[0].Annotations!.Items);
        }
        var expected = (await Convert("20 30 40 50")).Appearance;
        var actual = (await Convert("40 50 20 30")).Appearance;
        Assert.Equal(expected.XMillimeters, actual.XMillimeters);
        Assert.Equal(expected.YMillimeters, actual.YMillimeters);
        Assert.Equal(expected.WidthMillimeters, actual.WidthMillimeters);
        Assert.Equal(expected.HeightMillimeters, actual.HeightMillimeters);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ZeroAreaFileAttachmentRectangleFailsBeforeOutput(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture, annotationRect: "20 30 20 50"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector) await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
            { Compatibility = options }).ConvertAsync(input, output));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output));
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

    [Theory]
    [InlineData("/FlateDecode", false)]
    [InlineData("[/FlateDecode]", false)]
    [InlineData("[/Fl]", true)]
    public async Task CompressedAttachmentMayExceedItsDecodedByteLimit(string filter, bool vector)
    {
        using var compressed = new MemoryStream();
        using (var deflater = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            deflater.WriteByte((byte)'A');
        var encoded = compressed.ToArray();
        Assert.True(encoded.Length > 4);
        var placeholder = new string('X', encoded.Length);
        using var fixture = new PdfFixture(Font());
        var pdf = Source(fixture, globalStreamObject: $"<< /Type /EmbeddedFile /Filter {filter} /Length {encoded.Length} >>\nstream\n{placeholder}\nendstream");
        var marker = Encoding.ASCII.GetBytes("stream\n" + placeholder + "\nendstream");
        var position = pdf.AsSpan().IndexOf(marker);
        Assert.True(position >= 0);
        encoded.AsSpan().CopyTo(pdf.AsSpan(position + "stream\n".Length));
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var options = new PdfToOfdOptions
            { CopyAttachments = true, MaxAttachmentBytes = 4, MaxTotalAttachmentBytes = 5, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertWithResultAsync(input, output, new[] { 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(new byte[] { (byte)'A' }, package.Attachments[0].Data);
        Assert.Equal("PAGE"u8.ToArray(), package.Attachments[1].Data);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NameTreeCopiesUnicodeOnlyEmbeddedFile(bool vector)
    {
        using var fixture = new PdfFixture(Font());
        using var input = new MemoryStream(Source(fixture,
            globalFileSpec: "<< /Type /Filespec /UF (unicode.bin) /EF << /UF 18 0 R >> >>"));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertAsync(input, output, new[] { 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("GLOBAL"u8.ToArray(), package.Attachments[0].Data);
        Assert.Equal("same.bin", package.Attachments[0].Name);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public async Task HexUnicodeNamesWorkForGlobalAndPageAttachments(bool vector, bool indirect)
    {
        const string encodedName = "<FEFF4E2D6587002E00620069006E>"; // 中文.bin
        using var fixture = new PdfFixture(Font());
        var nameToken = indirect ? "22 0 R" : encodedName;
        var pageNameToken = indirect ? "23 0 R" : encodedName;
        using var input = new MemoryStream(Source(fixture, nameTreeKeyToken: nameToken,
            pageFileSpec: $"<< /Type /Filespec /UF {pageNameToken} /EF << /F 21 0 R >> >>",
            extraObjects: indirect ? new[] { encodedName, encodedName } : null));
        using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, PreferExternalPdfToPpm = true };
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertAsync(input, output, new[] { 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(new[] { "中文.bin", "中文.bin (2)" }, package.Attachments.Select(item => item.Name));
        Assert.Equal(package.Attachments[1].Id, Assert.Single(package.Pages[0].Annotations!.Items).AttachmentId);
    }

    [Theory]
    [InlineData(false, 3, true)]
    [InlineData(true, 3, true)]
    [InlineData(false, 2, false)]
    [InlineData(true, 2, false)]
    public async Task AliasedCompressedStreamCountsEachOutputOccurrence(bool vector, int totalBytes, bool accepted)
    {
        using var compressed = new MemoryStream();
        using (var deflater = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            deflater.WriteByte((byte)'A');
        var encoded = compressed.ToArray();
        var placeholder = new string('X', encoded.Length);
        using var fixture = new PdfFixture(Font());
        var pdf = Source(fixture, globalStreamObject:
            $"<< /Type /EmbeddedFile /Filter /FlateDecode /Length {encoded.Length} >>\nstream\n{placeholder}\nendstream",
            pageFileSpec: "<< /Type /Filespec /F (page.bin) /EF << /F 18 0 R >> >>");
        var marker = Encoding.ASCII.GetBytes("stream\n" + placeholder + "\nendstream");
        var position = pdf.AsSpan().IndexOf(marker);
        Assert.True(position >= 0);
        encoded.AsSpan().CopyTo(pdf.AsSpan(position + "stream\n".Length));
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var options = new PdfToOfdOptions { CopyAttachments = true, MaxAttachmentBytes = 1,
            MaxTotalAttachmentBytes = totalBytes, PreferExternalPdfToPpm = true };
        if (!accepted)
        {
            if (vector)
                await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                    { Compatibility = options }).ConvertAsync(input, output, new[] { 0, 0 }));
            else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0, 0 }));
            Assert.Equal(0, output.Length);
            return;
        }
        if (vector)
            await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { Compatibility = options })
                .ConvertAsync(input, output, new[] { 0, 0 });
        else await new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0, 0 });
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(3, package.Attachments.Count);
        Assert.All(package.Attachments, item => Assert.Equal("A"u8.ToArray(), item.Data));
        Assert.Equal(2, package.Pages.Count);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NameTreeCompressedPayloadHonorsDecodedLimitBeforeOutput(bool vector)
    {
        using var compressed = new MemoryStream();
        using (var deflater = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true))
            deflater.Write(new byte[1024 * 1024]);
        var encoded = compressed.ToArray();
        var placeholder = new string('X', encoded.Length);
        using var fixture = new PdfFixture(Font());
        var pdf = Source(fixture, globalStreamObject:
            $"<< /Type /EmbeddedFile /Filter [/FlateDecode] /Length {encoded.Length} >>\nstream\n{placeholder}\nendstream");
        var marker = Encoding.ASCII.GetBytes("stream\n" + placeholder + "\nendstream");
        var position = pdf.AsSpan().IndexOf(marker);
        Assert.True(position >= 0);
        encoded.AsSpan().CopyTo(pdf.AsSpan(position + "stream\n".Length));
        using var input = new MemoryStream(pdf); using var output = new MemoryStream();
        var options = new PdfToOfdOptions
            { CopyAttachments = true, MaxAttachmentBytes = 1024, MaxTotalAttachmentBytes = 2048, PreferExternalPdfToPpm = true };
        if (vector)
            await Assert.ThrowsAsync<InvalidDataException>(() => new PdfVectorToOfdConverter(new PdfVectorToOfdOptions
                { Compatibility = options }).ConvertAsync(input, output, new[] { 0 }));
        else await Assert.ThrowsAsync<InvalidDataException>(() => new PdfToOfdConverter(options).ConvertAsync(input, output, new[] { 0 }));
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData("missing")]
    [InlineData("invalid-action-id")]
    [InlineData("duplicate")]
    [InlineData("duplicate-numeric")]
    [InlineData("duplicate-external")]
    [InlineData("duplicate-numeric-external")]
    [InlineData("external")]
    [InlineData("invalid-attachment-id")]
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
            else if (defect == "invalid-action-id")
            {
                action.SetAttributeValue("AttachID", "0");
                ReplaceXml(annotationPath, annotation);
            }
            else
            {
                var attachmentPath = archive.Entries.Single(entry => entry.FullName.EndsWith("/Attachments.xml")).FullName;
                var attachments = ReadXml(attachmentPath);
                var item = attachments.Descendants().Single(node => node.Name.LocalName == "Attachment" && node.Attribute("ID")?.Value == target);
                if (defect == "duplicate") item.AddAfterSelf(new XElement(item));
                else if (defect == "duplicate-numeric")
                {
                    var other = new XElement(item);
                    other.SetAttributeValue("ID", "0" + target);
                    item.AddAfterSelf(other);
                }
                else if (defect == "duplicate-external")
                {
                    var other = new XElement(item);
                    other.SetAttributeValue("External", "true");
                    other.Elements().Single(node => node.Name.LocalName == "FileLoc").Value = "external.bin";
                    item.AddAfterSelf(other);
                }
                else if (defect == "duplicate-numeric-external")
                {
                    var other = new XElement(item);
                    other.SetAttributeValue("ID", "0" + target);
                    other.SetAttributeValue("External", "true");
                    other.Elements().Single(node => node.Name.LocalName == "FileLoc").Value = "external.bin";
                    item.AddAfterSelf(other);
                }
                else if (defect == "invalid-attachment-id") item.SetAttributeValue("ID", "0");
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
