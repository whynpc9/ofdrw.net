using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit017RegressionTests
{
    [Fact]
    public void Audit017_SamePartRepeated()
    {
        using var source = new Source();
        var bytes = Png(255, 0, 0);
        source.Create(main =>
        {
            AddImage(main, "rIdImage", bytes);
            for (var i = 0; i < 12; i++) main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12));
        });
        var operations = new List<string>();
        var reader = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default, operations.Add);
        var images = Images(reader.Read(source.Path));
        Assert.Equal(new[] { "Open", "Identify" }, operations);
        Assert.Equal(12, images.Length);
        Assert.All(images, image => Assert.Equal(bytes, image.Data));
        Assert.All(images, image => Assert.Same(images[0].Data, image.Data));
    }

    [Fact]
    public void Audit017_DifferentExtents()
    {
        using var source = new Source();
        source.Create(main =>
        {
            AddImage(main, "rIdImage", Png(255, 0, 0));
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
        });
        var images = Images(Read(source));
        Assert.Equal(2, images.Length);
        Assert.NotSame(images[0], images[1]);
        Assert.Equal(24, images[0].WidthPoints); Assert.Equal(12, images[0].HeightPoints);
        Assert.Equal(48, images[1].WidthPoints); Assert.Equal(36, images[1].HeightPoints);
    }

    [Fact]
    public void Audit017_SameRelationIdDifferentPart()
    {
        using var source = new Source();
        var red = Png(255, 0, 0); var blue = Png(0, 0, 255);
        source.Create(main =>
        {
            AddImage(main, "rIdImage", red);
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12));
            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            AddImage(header, "rIdImage", blue);
            header.Header = new W.Header(Paragraph("rIdImage", 12, 6));
            main.Document!.Body.Append(new W.SectionProperties(new W.HeaderReference
                { Id = "rIdHeader", Type = W.HeaderFooterValues.Default }));
        });
        var model = Read(source); var body = Assert.Single(Images(model));
        var headerImage = Assert.Single(model.Sections[0].Headers.SelectMany(p => p.Inlines).OfType<BuiltInImageModel>());
        Assert.Equal(red, body.Data); Assert.Equal(blue, headerImage.Data);
        Assert.NotSame(body.Data, headerImage.Data);
    }

    [Fact]
    public void Audit017_DistinctAndFailedParts()
    {
        using var source = new Source();
        var red = Png(255, 0, 0); var blue = Png(0, 0, 255);
        source.Create(main =>
        {
            AddImage(main, "rIdRed", red); AddImage(main, "rIdBlue", blue);
            AddImage(main, "rIdBroken", new byte[] { 1, 2, 3, 4 });
            main.Document!.Body!.Append(Paragraph("rIdRed", 24, 12), Paragraph("rIdBroken", 24, 12),
                Paragraph("rIdBlue", 24, 12), Paragraph("rIdBroken", 24, 12));
        });
        var diagnostics = new List<DocxConversionDiagnostic>();
        var model = new DocxModelReader(new DocxConversionOptions(), diagnostics, default).Read(source.Path);
        var images = Images(model);
        Assert.Equal(2, images.Length); Assert.Equal(red, images[0].Data); Assert.Equal(blue, images[1].Data);
        Assert.NotSame(images[0].Data, images[1].Data);
        Assert.Equal(2, diagnostics.Count(d => d.Code == "DOCX_IMAGE_FORMAT_UNSUPPORTED"));
        Assert.Equal(2, model.Sections[0].Blocks.OfType<BuiltInParagraphModel>()
            .SelectMany(p => p.Inlines).OfType<BuiltInTextModel>().Count(t => t.Text.Contains("Unsupported")));
    }

    [Fact]
    public async Task Audit017_ConcurrentConversions()
    {
        using var first = new Source(); using var second = new Source();
        var red = Png(255, 0, 0); var blue = Png(0, 0, 255);
        first.Create(main => { AddImage(main, "rIdImage", red); main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12)); });
        second.Create(main => { AddImage(main, "rIdImage", blue); main.Document!.Body!.Append(Paragraph("rIdImage", 48, 36)); });
        var models = await Task.WhenAll(Task.Run(() => Read(first)), Task.Run(() => Read(second)), Task.Run(() => Read(first)));
        var images = models.Select(m => Assert.Single(Images(m))).ToArray();
        Assert.Equal(red, images[0].Data); Assert.Equal(blue, images[1].Data); Assert.Equal(red, images[2].Data);
        Assert.NotSame(images[0].Data, images[1].Data); Assert.NotSame(images[0].Data, images[2].Data);
        Assert.Equal(24, images[0].WidthPoints); Assert.Equal(48, images[1].WidthPoints);
    }

    [Fact]
    public void Audit017_AliasesAcrossOwningPartsShareActualPart()
    {
        using var source = new Source();
        source.Create(main =>
        {
            var image = AddImage(main, "rIdImage", Png(255, 0, 0));
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
            var header = main.AddNewPart<HeaderPart>("rIdHeader");
            header.AddPart(image, "rIdAlias");
            header.Header = new W.Header(Paragraph("rIdAlias", 12, 6));
            main.Document!.Body.Append(new W.SectionProperties(new W.HeaderReference
                { Id = "rIdHeader", Type = W.HeaderFooterValues.Default }));
        });
        var operations = new List<string>();
        var model = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default, operations.Add).Read(source.Path);
        var body = Images(model); var headerImage = Assert.Single(model.Sections[0].Headers.SelectMany(p => p.Inlines).OfType<BuiltInImageModel>());
        Assert.Equal(new[] { "Open", "Identify" }, operations);
        Assert.Same(body[0].Data, body[1].Data); Assert.Same(body[0].Data, headerImage.Data);
        Assert.Equal(12, headerImage.WidthPoints); Assert.Equal(6, headerImage.HeightPoints);
    }

    [Fact]
    public void Audit017_DistinctPartsWithIdenticalBytesRemainIndependent()
    {
        using var source = new Source(); var bytes = Png(255, 0, 0);
        source.Create(main =>
        {
            AddImage(main, "rIdFirst", bytes); AddImage(main, "rIdSecond", bytes);
            main.Document!.Body!.Append(Paragraph("rIdFirst", 24, 12), Paragraph("rIdSecond", 24, 12));
        });
        var operations = new List<string>();
        var images = Images(new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default, operations.Add).Read(source.Path));
        Assert.Equal(2, operations.Count(o => o == "Open")); Assert.Equal(2, operations.Count(o => o == "Identify"));
        Assert.Equal(bytes, images[0].Data); Assert.Equal(bytes, images[1].Data); Assert.NotSame(images[0].Data, images[1].Data);
    }

    [Theory]
    [InlineData(UnsupportedDocxFeatureBehavior.Placeholder)]
    [InlineData(UnsupportedDocxFeatureBehavior.BestEffort)]
    [InlineData(UnsupportedDocxFeatureBehavior.Throw)]
    public void Audit017_FailedPartsAreNeverCachedAsSuccess(UnsupportedDocxFeatureBehavior behavior)
    {
        using var source = new Source();
        source.Create(main =>
        {
            AddImage(main, "rIdBroken", new byte[] { 1, 2, 3, 4 });
            main.Document!.Body!.Append(Paragraph("rIdBroken", 24, 12), Paragraph("rIdBroken", 24, 12));
        });
        var operations = new List<string>(); var diagnostics = new List<DocxConversionDiagnostic>();
        var reader = new DocxModelReader(new DocxConversionOptions { UnsupportedFeatureBehavior = behavior }, diagnostics, default, operations.Add);
        if (behavior == UnsupportedDocxFeatureBehavior.Throw)
        {
            Assert.Throws<NotSupportedException>(() => reader.Read(source.Path));
            Assert.Equal(1, operations.Count(o => o == "Open"));
        }
        else
        {
            Assert.Empty(Images(reader.Read(source.Path)));
            Assert.Equal(2, operations.Count(o => o == "Open"));
            Assert.Equal(2, diagnostics.Count(d => d.Code == "DOCX_IMAGE_FORMAT_UNSUPPORTED"));
        }
        // Identify returns null for this tiny unknown payload; neither occurrence becomes a success.
        Assert.Equal(behavior == UnsupportedDocxFeatureBehavior.Throw ? 1 : 2, operations.Count(o => o == "Identify"));
    }

    [Fact]
    public void Audit017_TruncatedRecognizedImageRetainsDecoderFailure()
    {
        using var source = new Source(); var truncated = Png(255, 0, 0).Take(20).ToArray();
        source.Create(main =>
        {
            AddImage(main, "rIdImage", truncated);
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
        });
        var operations = new List<string>();
        var reader = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default, operations.Add);
        Assert.Throws<InvalidImageContentException>(() => reader.Read(source.Path));
        Assert.Equal(new[] { "Open" }, operations);
        // The failed conversion must retry the actual stream, never reuse partial bytes.
        Assert.Throws<InvalidImageContentException>(() => reader.Read(source.Path));
        Assert.Equal(new[] { "Open", "Open" }, operations);
    }

    [Theory]
    [InlineData("bytes")]
    [InlineData("pixels")]
    public void Audit017_AdmissionLimitsAndFailedReadRetry(string limit)
    {
        using var source = new Source(); var bytes = Png(255, 0, 0);
        source.Create(main =>
        {
            AddImage(main, "rIdImage", bytes);
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
        });
        var options = new DocxConversionOptions { MaxEmbeddedImageBytes = limit == "bytes" ? bytes.Length - 1 : bytes.Length,
            MaxEmbeddedImagePixels = limit == "pixels" ? 3 : 4 };
        var operations = new List<string>();
        var reader = new DocxModelReader(options, new List<DocxConversionDiagnostic>(), default, operations.Add);
        Assert.Throws<InvalidDataException>(() => reader.Read(source.Path));
        Assert.Equal(1, operations.Count(o => o == "Open"));
        Assert.Equal(limit == "pixels" ? 1 : 0, operations.Count(o => o == "Identify"));
        options.MaxEmbeddedImageBytes = bytes.Length; options.MaxEmbeddedImagePixels = 4; operations.Clear();
        var images = Images(reader.Read(source.Path));
        Assert.Equal(2, images.Length); Assert.Same(images[0].Data, images[1].Data);
        Assert.Equal(new[] { "Open", "Identify" }, operations);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void Audit017_BigTiffRejectedBeforeIdentify(bool littleEndian)
    {
        using var source = new Source();
        source.Create(main =>
        {
            AddImage(main, "rIdImage", littleEndian ? new byte[] { 0x49, 0x49, 0x2b, 0 } : new byte[] { 0x4d, 0x4d, 0, 0x2b });
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
        });
        var operations = new List<string>();
        var reader = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default, operations.Add);
        var error = Assert.Throws<InvalidDataException>(() => reader.Read(source.Path));
        Assert.Contains("BigTIFF", error.Message); Assert.Equal(new[] { "Open" }, operations);
    }

    [Theory]
    [InlineData("BeforeRead")]
    [InlineData("Open")]
    [InlineData("Identify")]
    public void Audit017_CancellationRemainsCancellation(string cancelAt)
    {
        using var source = new Source();
        source.Create(main =>
        {
            AddImage(main, "rIdImage", Png(255, 0, 0));
            main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
        });
        using var cancellation = new CancellationTokenSource(); var operations = new List<string>();
        if (cancelAt == "BeforeRead") cancellation.Cancel();
        var reader = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), cancellation.Token,
            operation => { operations.Add(operation); if (operation == cancelAt) cancellation.Cancel(); });
        Assert.Throws<OperationCanceledException>(() => reader.Read(source.Path));
        Assert.Equal(cancelAt == "BeforeRead" ? 0 : 1, operations.Count(o => o == "Open"));
        Assert.Equal(cancelAt == "Identify" ? 1 : 0, operations.Count(o => o == "Identify"));
    }

    [Fact]
    public void Audit017_ReaderDropsCacheBetweenPackagesAndKeepsReturnedBytes()
    {
        using var first = new Source(); using var second = new Source(); var red = Png(255, 0, 0); var blue = Png(0, 0, 255);
        first.Create(main => { AddImage(main, "rIdImage", red); main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12)); });
        second.Create(main => { AddImage(main, "rIdImage", blue); main.Document!.Body!.Append(Paragraph("rIdImage", 24, 12)); });
        var operations = new List<string>();
        var reader = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default, operations.Add);
        var one = Assert.Single(Images(reader.Read(first.Path))); var two = Assert.Single(Images(reader.Read(second.Path)));
        var three = Assert.Single(Images(reader.Read(first.Path)));
        Assert.Equal(red, one.Data); Assert.Equal(blue, two.Data); Assert.Equal(red, three.Data);
        Assert.NotSame(one.Data, three.Data); Assert.Equal(3, operations.Count(o => o == "Open"));
        Assert.Equal(3, operations.Count(o => o == "Identify"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit017_PublicNativeAndDefaultPreserveTextImagesExtentsAndSelectedPage(bool explicitNative)
    {
        using var source = new Source(); var red = Png(255, 0, 0); var blue = Png(0, 0, 255);
        source.Create(main =>
        {
            AddImage(main, "rIdImage", red);
            var body = main.Document!.Body!;
            body.Append(new W.Paragraph(new W.Run(new W.Text("017 Shared image 中文 A"))));
            body.Append(Paragraph("rIdImage", 24, 12), Paragraph("rIdImage", 48, 36));
            body.Append(new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })));
            body.Append(new W.Paragraph(new W.Run(new W.Text("017 Shared image 中文 B"))));
            body.Append(Paragraph("rIdImage", 36, 18));
            var header = main.AddNewPart<HeaderPart>("rIdHeader"); AddImage(header, "rIdImage", blue);
            header.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("Blue header"))), Paragraph("rIdImage", 12, 6));
            body.Append(new W.SectionProperties(new W.HeaderReference { Id = "rIdHeader", Type = W.HeaderFooterValues.Default }));
        });
        var options = new DocxConversionOptions(); if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        var converter = new DocxToOfdConverter(options);
        using var full = new MemoryStream();
        await using (var input = File.OpenRead(source.Path)) await converter.ConvertAsync(input, full);
        full.Position = 0; var model = await new OfdReader().ReadAsync(full);
        Assert.Equal(2, model.Pages.Count);
        var images = model.Pages.SelectMany(p => p.Elements).OfType<OfdImageElement>().ToArray();
        Assert.Equal(5, images.Length);
        Assert.Equal(3, images.Count(i => i.Data.SequenceEqual(red))); Assert.Equal(2, images.Count(i => i.Data.SequenceEqual(blue)));
        var bodyImages = images.Where(i => i.Data.SequenceEqual(red)).ToArray();
        var widths = new[] { 24d, 48d, 36d }; var heights = new[] { 12d, 36d, 18d };
        for (var i = 0; i < bodyImages.Length; i++)
        {
            // OFD serializes millimeter coordinates to three decimal places.
            Assert.Equal(widths[i] * 25.4 / 72, bodyImages[i].WidthMillimeters, 3);
            Assert.Equal(heights[i] * 25.4 / 72, bodyImages[i].HeightMillimeters, 3);
        }
        Assert.Contains("017 Shared image 中文 A", string.Concat(model.Pages[0].Elements.OfType<OfdTextElement>().Select(t => t.Text)));
        Assert.Contains("017 Shared image 中文 B", string.Concat(model.Pages[1].Elements.OfType<OfdTextElement>().Select(t => t.Text)));
        using var selected = new MemoryStream();
        await using (var input = File.OpenRead(source.Path)) await converter.ConvertAsync(input, selected, new[] { 1 }); // Public API page indexes are zero-based.
        selected.Position = 0; var selectedModel = await new OfdReader().ReadAsync(selected);
        var selectedPage = Assert.Single(selectedModel.Pages);
        Assert.Equal(2, selectedPage.Elements.OfType<OfdImageElement>().Count());
        var selectedText = string.Concat(selectedPage.Elements.OfType<OfdTextElement>().Select(t => t.Text));
        Assert.Contains("017 Shared image 中文 B", selectedText); Assert.DoesNotContain("017 Shared image 中文 A", selectedText);
        var artifacts = Environment.GetEnvironmentVariable("OFDRW_PLAN017_VISUAL");
        if (artifacts is not null)
        {
            Directory.CreateDirectory(artifacts); var mode = explicitNative ? "native" : "default";
            File.Copy(source.Path, Path.Combine(artifacts, "repeated-images.docx"), true);
            File.WriteAllBytes(Path.Combine(artifacts, $"repeated-images-{mode}.ofd"), full.ToArray());
            File.WriteAllBytes(Path.Combine(artifacts, $"repeated-images-{mode}-page2.ofd"), selected.ToArray());
            full.Position = 0; using var pdf = File.Create(Path.Combine(artifacts, $"repeated-images-{mode}.pdf"));
            await new OfdToPdfConverter().ConvertAsync(full, pdf);
        }
    }

    private static BuiltInDocumentModel Read(Source source) =>
        new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default).Read(source.Path);

    private static BuiltInImageModel[] Images(BuiltInDocumentModel model) => model.Sections
        .SelectMany(s => s.Blocks).OfType<BuiltInParagraphModel>().SelectMany(p => p.Inlines).OfType<BuiltInImageModel>().ToArray();

    private static byte[] Png(byte r, byte g, byte b)
    {
        using var image = new Image<Rgba32>(2, 2, new Rgba32(r, g, b));
        using var data = new MemoryStream(); image.SaveAsPng(data); return data.ToArray();
    }

    private static ImagePart AddImage(OpenXmlPartContainer owner, string id, byte[] bytes)
    {
        var part = owner switch
        {
            MainDocumentPart main => main.AddImagePart(ImagePartType.Png, id),
            HeaderPart header => header.AddImagePart(ImagePartType.Png, id),
            _ => throw new ArgumentException("Unsupported fixture owner.")
        };
        using var data = new MemoryStream(bytes); part.FeedData(data); return part;
    }

    private static W.Paragraph Paragraph(string id, int width, int height) => new($"""
        <w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
             xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
             xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
             xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
          <w:r><w:drawing><wp:inline><wp:extent cx="{width * 12700}" cy="{height * 12700}"/>
          <a:graphic><a:graphicData><a:blip r:embed="{id}"/></a:graphicData></a:graphic>
          </wp:inline></w:drawing></w:r>
        </w:p>
        """);

    private sealed class Source : IDisposable
    {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ofdrw-plan017-" + Guid.NewGuid() + ".docx");
        internal void Create(Action<MainDocumentPart> populate)
        {
            using var document = WordprocessingDocument.Create(Path, WordprocessingDocumentType.Document);
            var main = document.AddMainDocumentPart(); main.Document = new W.Document(new W.Body()); populate(main);
        }
        public void Dispose() => File.Delete(Path);
    }
}
