using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Tests;

/// <summary>
/// Covers the generated, non-sensitive DOCX conversion fixture.
/// </summary>
public sealed partial class DocxConversionTests
{
    /// <summary>
    /// Verifies the in-process renderer without invoking Word or LibreOffice.
    /// </summary>
    [Fact]
    public async Task GeneratedDocx_ShouldConvertWithBuiltIn()
    {
        var samplePath = ResolveGeneratedSample();
        var converter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn
        });

        await using var pdf = new MemoryStream();
        await using (var docx = File.OpenRead(samplePath))
        {
            var result = await converter.ConvertWithResultAsync(docx, pdf);
            Assert.Equal(DocxConversionEngine.BuiltIn, result.ActualEngine);
            Assert.Equal(new[] { DocxConversionEngine.BuiltIn }, result.AttemptedEngines);
        }

        AssertPdfHasTwoPages(pdf);
        pdf.Position = 0;
        using var document = PdfPigDocument.Open(pdf);
        var text = string.Join(" ", document.GetPages().Select(page => page.Text));
        var compactText = new string(text.Where(character => !char.IsWhiteSpace(character)).ToArray());
        Assert.Contains("GeneratedDOCX", compactText, StringComparison.Ordinal);
        Assert.Contains("Alpha", compactText, StringComparison.Ordinal);
    }

    /// <summary>
    /// Verifies Auto remains usable when neither desktop renderer is available.
    /// </summary>
    [Fact]
    public async Task Auto_ShouldUseBuiltInWhenDesktopEnginesAreUnavailable()
    {
        var converter = new DocxToPdfConverter(
            new DocxConversionOptions { Engine = DocxConversionEngine.Auto },
            canUseMicrosoftWord: () => false,
            canUseLibreOffice: () => false);

        await using var output = new MemoryStream();
        await using (var input = File.OpenRead(ResolveGeneratedSample()))
        {
            var result = await converter.ConvertWithResultAsync(input, output);
            Assert.Equal(DocxConversionEngine.BuiltIn, result.ActualEngine);
            Assert.Equal(new[] { DocxConversionEngine.BuiltIn }, result.AttemptedEngines);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == "DOCX_LIBREOFFICE_UNAVAILABLE");
        }

        AssertPdfHasTwoPages(output);
    }

    /// <summary>
    /// Verifies Auto isolates a failed external attempt before committing BuiltIn output.
    /// </summary>
    [Fact]
    public async Task Auto_ShouldFallbackWhenLibreOfficeFailsToStart()
    {
        var converter = new DocxToPdfConverter(
            new DocxConversionOptions
            {
                Engine = DocxConversionEngine.Auto,
                LibreOfficePath = Path.Combine(Path.GetTempPath(), "missing-ofdrw-soffice")
            },
            canUseMicrosoftWord: () => false,
            canUseLibreOffice: () => true);

        await using var output = new MemoryStream();
        await using (var input = File.OpenRead(ResolveGeneratedSample()))
        {
            var result = await converter.ConvertWithResultAsync(input, output);
            Assert.Equal(DocxConversionEngine.BuiltIn, result.ActualEngine);
            Assert.Equal(
                new[] { DocxConversionEngine.LibreOffice, DocxConversionEngine.BuiltIn },
                result.AttemptedEngines);
            Assert.Contains(result.Diagnostics, diagnostic =>
                diagnostic.Code == "DOCX_ENGINE_FAILED");
        }

        AssertPdfHasTwoPages(output);
    }

    /// <summary>
    /// Verifies native DOCX-to-OFD skips desktop/PDF rendering and preserves source text.
    /// </summary>
    [Fact]
    public async Task GeneratedDocx_ShouldConvertDirectlyToNativeOfdText()
    {
        var converter = new DocxToOfdConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.MicrosoftWord,
            OfdMode = DocxToOfdMode.Native
        });

        await using var output = new MemoryStream();
        await using (var input = File.OpenRead(ResolveGeneratedSample()))
        {
            await converter.ConvertAsync(input, output);
        }

        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.NotEmpty(package.Pages);
        Assert.All(package.Pages, page =>
        {
            Assert.Empty(page.Elements.OfType<OfdImageElement>());
            Assert.NotEmpty(page.Elements.OfType<OfdTextElement>());
            Assert.All(page.Elements.OfType<OfdTextElement>(), text =>
                Assert.Equal(255, text.FillColor.Alpha));
        });
        Assert.Equal("DOCX/OpenXML", package.CustomTags["source-text-origin"]);
        Assert.Equal("Native", package.CustomTags["docx-ofd-mode"]);
        Assert.Equal(ReadExpectedSourceText(), CompactExtractedText(package));
        Assert.Contains(package.Pages, page => page.Elements.OfType<OfdPathElement>().Any() ||
            page.Elements.OfType<OfdTextElement>().Select(text => text.XMillimeters).Distinct().Count() > 1);
    }

    private static (int, int, int) Rgb(OfdColor color) => (color.Red, color.Green, color.Blue);

    [Fact]
    public async Task Native_ShouldPreserveVisualStylesAndProportionalAdvances()
    {
        await using var input = File.OpenRead(ResolveGeneratedSample());
        await using var output = new MemoryStream();
        await new DocxToOfdConverter().ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Pages.Count);
        var text = package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().ToList();
        var normal = Assert.Single(text, t => t.Text == "第二页用于确认分页保持稳定。");
        var emphasized = Assert.Single(text, t => t.Text.Contains("这段文字应为红色粗体。"));
        Assert.Equal(OfdColor.Black, normal.FillColor);
        Assert.False(package.Fonts.Single(f => f.FontName == normal.FontName).Bold);
        Assert.Equal((192, 0, 0), Rgb(emphasized.FillColor));
        Assert.True(package.Fonts.Single(f => f.FontName == emphasized.FontName).Bold);
        Assert.Equal(normal.YMillimeters, emphasized.YMillimeters);
        Assert.True(emphasized.XMillimeters >= normal.XMillimeters + normal.WidthMillimeters - 0.002);
        var subtitle = Assert.Single(text, t => t.Text.Contains("Generated DOCX"));
        Assert.True(package.Fonts.Single(f => f.FontName == subtitle.FontName).Italic);
        Assert.All(package.Fonts, font => Assert.NotEmpty(font.Data));
        var paths = package.Pages[0].Elements.OfType<OfdPathElement>().ToList();
        Assert.Equal(3, paths.Count(p => p.Fill && p.FillColor is not null && Rgb(p.FillColor) == (217, 234, 247)));
        Assert.Contains(paths, p => p.Stroke && Rgb(p.StrokeColor) == (68, 114, 196));
        Assert.Contains(paths, p => p.Stroke && Rgb(p.StrokeColor) == (165, 165, 165));
        var alpha = Assert.Single(text, t => t.Text == "Alpha");
        var advances = alpha.Runs.Single().DeltaX!.Split(' ').Select(v => double.Parse(v, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
        Assert.True(advances[0] > advances[1], "Proportional A must be wider than l.");
        Assert.Contains(text, t => t.Text.Contains("information"));
        Assert.Equal(ReadExpectedSourceText(), CompactExtractedText(package));

        // Exercise the actual OFD -> PDF path, including embedded fallback face simulation.
        output.Position = 0;
        await using var pdfOutput = new MemoryStream();
        await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(output, pdfOutput);
        pdfOutput.Position = 0;
        using var pdf = PdfPigDocument.Open(pdfOutput);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(1, pdf.GetPage(1).Letters.Count(letter => letter.Value == "档"));
        Assert.Equal(1, pdf.GetPage(2).Letters.Count(letter => letter.Value == "红"));
        Assert.Equal(ReadExpectedSourceText(), new string(string.Concat(pdf.GetPages().Select(page => page.Text))
            .Where(character => !char.IsWhiteSpace(character)).ToArray()));
        Assert.Contains("Generated", pdf.GetPage(1).Text);
        var italicGlyph = pdf.GetPage(1).Letters.First(letter => letter.Value == "G").BoundingBox;
        Assert.True(italicGlyph.TopLeft.X > italicGlyph.BottomLeft.X + 0.1,
            "The italic subtitle must be visibly slanted after OFD -> PDF export.");
    }

    [Fact]
    public async Task Native_ShouldKeepMixedSizesAndCellBorderOverrides()
    {
        await using var input = CreateMinimalDocx("""
            <w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>Small</w:t></w:r><w:r><w:rPr><w:sz w:val="40"/><w:i/></w:rPr><w:t>Large</w:t></w:r><w:r><w:t>Normal</w:t></w:r></w:p>
            <w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="0000FF"/><w:bottom w:val="single" w:sz="8" w:color="0000FF"/><w:left w:val="single"/><w:right w:val="single"/></w:tblBorders></w:tblPr>
            <w:tblGrid><w:gridCol w:w="6000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:bottom w:val="single" w:sz="16" w:color="FF0000"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>Plain</w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="008000"/></w:rPr><w:t>Green</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
            """);
        await using var output = new MemoryStream();
        await new DocxToOfdConverter().ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var texts = package.Pages[0].Elements.OfType<OfdTextElement>().ToList();
        var small = texts.Single(t => t.Text == "Small");
        var large = texts.Single(t => t.Text == "Large");
        Assert.InRange(large.FontSizeMillimeters / small.FontSizeMillimeters, 1.99, 2.01);
        Assert.Equal(small.YMillimeters + small.Runs.Single().YMillimeters,
            large.YMillimeters + large.Runs.Single().YMillimeters);
        Assert.True(package.Fonts.Single(f => f.FontName == large.FontName).Italic);
        Assert.Equal(OfdColor.Black, texts.Single(t => t.Text == "Plain").FillColor);
        Assert.Equal((0, 128, 0), Rgb(texts.Single(t => t.Text == "Green").FillColor));
        var borders = package.Pages[0].Elements.OfType<OfdPathElement>().Where(p => p.Stroke).ToList();
        Assert.Equal(3, borders.Count);
        var red = Assert.Single(borders, p => Rgb(p.StrokeColor) == (255, 0, 0));
        Assert.InRange(red.LineWidthMillimeters, 0.705, 0.707);
    }

    /// <summary>
    /// Verifies the dual-layer route keeps rendered pages but uses DOCX source text.
    /// </summary>
    [Fact]
    public async Task GeneratedDocx_ShouldConvertToDualLayerOfdWithBuiltIn()
    {
        var converter = new DocxToOfdConverter(
            new DocxConversionOptions
            {
                Engine = DocxConversionEngine.BuiltIn,
                OfdMode = DocxToOfdMode.DualLayer
            },
            new PdfToOfdOptions
            {
                TextLayerMode = PdfTextLayerMode.None
            });

        await using var output = new MemoryStream();
        await using (var input = File.OpenRead(ResolveGeneratedSample()))
        {
            await converter.ConvertAsync(input, output);
        }

        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Pages.Count);
        Assert.All(package.Pages, page =>
        {
            Assert.Single(page.Elements.OfType<OfdImageElement>());
            Assert.Single(page.Elements.OfType<OfdTextElement>());
            Assert.All(page.Elements.OfType<OfdTextElement>(), text =>
                Assert.Equal(0, text.FillColor.Alpha));
        });
        Assert.Equal("DOCX/OpenXML", package.CustomTags["source-text-origin"]);
        Assert.Equal("machine-readable", package.CustomTags["source-text-kind"]);
        Assert.Equal("DualLayer", package.CustomTags["docx-ofd-mode"]);
        Assert.Equal(ReadExpectedSourceText(), CompactExtractedText(package));
    }

    /// <summary>
    /// Verifies the real LibreOffice PDF and OFD pipeline.
    /// </summary>
    [Fact]
    public async Task GeneratedDocx_ShouldConvertToPdfAndOfd()
    {
        var samplePath = ResolveGeneratedSample();

        var options = new DocxConversionOptions
        {
            Engine = DocxConversionEngine.LibreOffice,
            OfdMode = DocxToOfdMode.DualLayer
        };
        var docxToPdf = new DocxToPdfConverter(options);
        await using var pdf = new MemoryStream();
        await using (var docx = File.OpenRead(samplePath))
        {
            await docxToPdf.ConvertAsync(docx, pdf);
        }

        AssertPdfHasTwoPages(pdf);

        var docxToOfd = new DocxToOfdConverter(options);
        await using var ofd = new MemoryStream();
        await using (var docx = File.OpenRead(samplePath))
        {
            await docxToOfd.ConvertAsync(docx, ofd);
        }

        ofd.Position = 0;
        var package = await new OfdReader().ReadAsync(ofd);
        Assert.Equal(2, package.Pages.Count);
        Assert.All(package.Pages, page =>
        {
            var image = Assert.Single(page.Elements.OfType<OfdImageElement>());
            Assert.Equal("image/png", image.MediaType);
            Assert.NotEmpty(image.Data);
            Assert.NotEmpty(page.Elements.OfType<OfdTextElement>());
        });
    }

    /// <summary>
    /// Verifies process timeout validation.
    /// </summary>
    [Fact]
    public void DocxToPdf_ShouldRejectNonPositiveTimeout()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            new DocxToPdfConverter(new DocxConversionOptions
            {
                ProcessTimeout = TimeSpan.Zero
            }));
    }

    /// <summary>
    /// Verifies the BuiltIn input resource gate before Open XML parsing.
    /// </summary>
    [Fact]
    public async Task BuiltIn_ShouldRejectInputLargerThanConfiguredLimit()
    {
        var converter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn,
            MaxInputBytes = 8
        });

        await using var input = new MemoryStream(new byte[9]);
        await using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// Verifies ProcessTimeout also bounds the in-process renderer.
    /// </summary>
    [Theory]
    [InlineData(1)]
    [InlineData(9999)]
    public async Task BuiltIn_ShouldHonorProcessTimeout(long ticks)
    {
        var converter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn,
            ProcessTimeout = TimeSpan.FromTicks(ticks)
        });

        await using var input = File.OpenRead(ResolveGeneratedSample());
        await using var output = new MemoryStream();
        await Assert.ThrowsAsync<TimeoutException>(() => converter.ConvertAsync(input, output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// Verifies BuiltIn never dereferences external package resources.
    /// </summary>
    [Fact]
    public async Task BuiltIn_ShouldRejectExternalImageRelationship()
    {
        const string relationships = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="https://example.invalid/tracker.png" TargetMode="External" />
            </Relationships>
            """;
        await using var input = CreateMinimalDocx("<w:p />", relationships);
        await using var output = new MemoryStream();
        var converter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn
        });

        await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, output));
        Assert.Empty(output.ToArray());
    }

    /// <summary>
    /// Verifies unsupported-feature policy can fail closed or emit a visible placeholder.
    /// </summary>
    [Fact]
    public async Task BuiltIn_ShouldApplyUnsupportedFeaturePolicy()
    {
        const string body = """
            <w:p><w:r><w:footnoteReference w:id="1" /></w:r></w:p>
            """;
        await using (var throwingInput = CreateMinimalDocx(body))
        await using (var throwingOutput = new MemoryStream())
        {
            var throwingConverter = new DocxToPdfConverter(new DocxConversionOptions
            {
                Engine = DocxConversionEngine.BuiltIn,
                UnsupportedFeatureBehavior = UnsupportedDocxFeatureBehavior.Throw
            });
            await Assert.ThrowsAsync<NotSupportedException>(() =>
                throwingConverter.ConvertAsync(throwingInput, throwingOutput));
            Assert.Empty(throwingOutput.ToArray());
        }

        await using var placeholderInput = CreateMinimalDocx(body);
        await using var placeholderOutput = new MemoryStream();
        var placeholderConverter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn,
            UnsupportedFeatureBehavior = UnsupportedDocxFeatureBehavior.Placeholder
        });
        var result = await placeholderConverter.ConvertWithResultAsync(
            placeholderInput,
            placeholderOutput);

        Assert.Contains(result.Diagnostics, diagnostic =>
            diagnostic.Code == "DOCX_FOOTNOTE_UNSUPPORTED");
        placeholderOutput.Position = 0;
        using var pdf = PdfPigDocument.Open(placeholderOutput);
        var placeholderText = new string(
            pdf.GetPage(1).Text.Where(character => !char.IsWhiteSpace(character)).ToArray());
        Assert.Contains("UnsupportedDOCXfeature", placeholderText, StringComparison.Ordinal);
    }

    /// <summary>
    /// Verifies an embedded raster relationship is rendered without external processes.
    /// </summary>
    [Fact]
    public async Task BuiltIn_ShouldRenderEmbeddedImage()
    {
        const string body = """
            <w:p><w:r><w:drawing><wp:inline>
              <wp:extent cx="9525" cy="9525" /><wp:docPr id="1" name="pixel" />
              <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="pixel.png" /><pic:cNvPicPr /></pic:nvPicPr>
                  <pic:blipFill><a:blip r:embed="rId2" /><a:stretch><a:fillRect /></a:stretch></pic:blipFill>
                  <pic:spPr><a:xfrm><a:off x="0" y="0" /><a:ext cx="9525" cy="9525" /></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst /></a:prstGeom></pic:spPr>
                </pic:pic>
              </a:graphicData></a:graphic>
            </wp:inline></w:drawing></w:r></w:p>
            """;
        const string relationships = """
            <?xml version="1.0" encoding="UTF-8"?>
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId2"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png" />
            </Relationships>
            """;
        var image = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=");
        await using var input = CreateMinimalDocx(
            body,
            relationships,
            new Dictionary<string, byte[]> { ["word/media/image1.png"] = image });
        await using var output = new MemoryStream();
        var converter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn
        });

        var result = await converter.ConvertWithResultAsync(input, output);

        Assert.DoesNotContain(result.Diagnostics, diagnostic =>
            diagnostic.Code == "DOCX_IMAGE_FORMAT_UNSUPPORTED");
        output.Position = 0;
        using var pdf = PdfPigDocument.Open(output);
        Assert.NotEmpty(pdf.GetPage(1).GetImages());
    }

    /// <summary>
    /// Verifies section header/footer relationships and a page-number field.
    /// </summary>
    [Fact]
    public async Task BuiltIn_ShouldRenderHeaderFooterAndPageNumber()
    {
        await using var input = CreateDocxWithHeaderAndFooter();
        await using var output = new MemoryStream();
        var converter = new DocxToPdfConverter(new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn
        });

        await converter.ConvertAsync(input, output);

        output.Position = 0;
        using var pdf = PdfPigDocument.Open(output);
        var text = new string(
            pdf.GetPage(1).Text.Where(character => !char.IsWhiteSpace(character)).ToArray());
        Assert.Contains("FixtureHeader", text, StringComparison.Ordinal);
        Assert.Contains("BodyText", text, StringComparison.Ordinal);
        Assert.Contains("Page1", text, StringComparison.Ordinal);
    }

    /// <summary>
    /// Verifies independent in-process conversions do not share document state.
    /// </summary>
    [Fact]
    public async Task BuiltIn_ShouldSupportConcurrentConversions()
    {
        var conversions = Enumerable.Range(0, 4).Select(async _ =>
        {
            var converter = new DocxToPdfConverter(new DocxConversionOptions
            {
                Engine = DocxConversionEngine.BuiltIn
            });
            await using var input = File.OpenRead(ResolveGeneratedSample());
            await using var output = new MemoryStream();
            await converter.ConvertAsync(input, output);
            AssertPdfHasTwoPages(output);
        });

        await Task.WhenAll(conversions);
    }

    private static void AssertPdfHasTwoPages(MemoryStream stream)
    {
        stream.Position = 0;
        using var pdf = PdfPigDocument.Open(stream);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.All(Enumerable.Range(1, pdf.NumberOfPages), pageNumber =>
        {
            var page = pdf.GetPage(pageNumber);
            Assert.InRange(page.Width, 590, 600);
            Assert.InRange(page.Height, 840, 850);
        });
    }

    private static string CompactExtractedText(OfdDocumentPackage package)
    {
        return new string(new OfdTextExtractor()
            .Extract(package)
            .Where(character => !char.IsWhiteSpace(character))
            .ToArray());
    }

    private static string ReadExpectedSourceText()
    {
        using var source = WordprocessingDocument.Open(ResolveGeneratedSample(), false);
        var body = source.MainDocumentPart?.Document?.Body ??
            throw new InvalidDataException("The generated DOCX fixture has no document body.");
        return new string(string.Concat(body
                .Descendants<W.Text>()
                .Select(text => text.Text))
            .Where(character => !char.IsWhiteSpace(character))
            .ToArray());
    }

    private static string ResolveGeneratedSample()
    {
        return Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Docx.E2E",
            "testdata",
            "generated-layout.docx");
    }

    private static MemoryStream CreateMinimalDocx(
        string bodyContent,
        string? documentRelationships = null,
        IReadOnlyDictionary<string, byte[]>? binaryParts = null)
    {
        var stream = new MemoryStream();
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            WriteZipEntry(archive, "[Content_Types].xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" />
                  <Default Extension="xml" ContentType="application/xml" />
                  <Default Extension="png" ContentType="image/png" />
                  <Override PartName="/word/document.xml"
                    ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml" />
                </Types>
                """);
            WriteZipEntry(archive, "_rels/.rels", """
                <?xml version="1.0" encoding="UTF-8"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1"
                    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
                    Target="word/document.xml" />
                </Relationships>
                """);
            WriteZipEntry(archive, "word/document.xml", $$"""
                <?xml version="1.0" encoding="UTF-8"?>
                <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
                  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                  xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                  <w:body>{{bodyContent}}<w:sectPr /></w:body>
                </w:document>
                """);
            if (documentRelationships is not null)
            {
                WriteZipEntry(
                    archive,
                    "word/_rels/document.xml.rels",
                    documentRelationships);
            }

            if (binaryParts is not null)
            {
                foreach (var part in binaryParts)
                {
                    var entry = archive.CreateEntry(part.Key);
                    using var entryStream = entry.Open();
                    entryStream.Write(part.Value);
                }
            }
        }

        stream.Position = 0;
        return stream;
    }

    private static void WriteZipEntry(ZipArchive archive, string path, string content)
    {
        var entry = archive.CreateEntry(path);
        using var writer = new StreamWriter(entry.Open());
        writer.Write(content);
    }

    private static MemoryStream CreateDocxWithHeaderAndFooter()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream,
                   WordprocessingDocumentType.Document,
                   autoSave: true))
        {
            var mainPart = document.AddMainDocumentPart();
            var headerPart = mainPart.AddNewPart<HeaderPart>();
            headerPart.Header = new W.Header(
                new W.Paragraph(new W.Run(new W.Text("Fixture Header"))));

            var footerPart = mainPart.AddNewPart<FooterPart>();
            footerPart.Footer = new W.Footer(
                new W.Paragraph(
                    new W.Run(new W.Text("Page ")),
                    new W.SimpleField(new W.Run(new W.Text("1")))
                    {
                        Instruction = "PAGE"
                    }));

            mainPart.Document = new W.Document(
                new W.Body(
                    new W.Paragraph(new W.Run(new W.Text("Body Text"))),
                    new W.SectionProperties(
                        new W.HeaderReference
                        {
                            Id = mainPart.GetIdOfPart(headerPart),
                            Type = W.HeaderFooterValues.Default
                        },
                        new W.FooterReference
                        {
                            Id = mainPart.GetIdOfPart(footerPart),
                            Type = W.HeaderFooterValues.Default
                        })));
        }

        stream.Position = 0;
        return stream;
    }

    private static string ResolveRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate the repository root.");
    }
}
