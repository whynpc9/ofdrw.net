using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Constants;
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
    [Theory]
    [InlineData(true, 80, 80)]
    [InlineData(false, 80, 80)]
    [InlineData(true, 80, 8)]
    [InlineData(false, 80, 8)]
    [InlineData(true, 8, 80)]
    [InlineData(false, 8, 80)]
    public async Task Native_EmptyLines_ShouldUseTheirRunFontSize(bool explicitNative, int firstSize, int secondSize)
    {
        await using var input = CreateMinimalDocx($"""
            <w:p><w:r><w:t>A</w:t></w:r><w:r><w:rPr><w:sz w:val="{firstSize}"/></w:rPr><w:br/></w:r><w:r><w:rPr><w:sz w:val="{secondSize}"/></w:rPr><w:cr/></w:r></w:p>
            <w:p><w:r><w:t>B</w:t></w:r></w:p>
            """);
        await using var output = new MemoryStream();
        var converter = explicitNative ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
            : new DocxToOfdConverter();
        await converter.ConvertAsync(input, output);output.Position = 0;
        var text = Assert.Single((await new OfdReader().ReadAsync(output)).Pages).Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "A", "B" }, text.Select(value => value.Text));
        Assert.Equal((10.5 + (firstSize + secondSize) / 2d) * 25.4 / 72 * 1.3,
            text[1].YMillimeters - text[0].YMillimeters, 2);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Native_DecomposedLatinI_ShouldMatchCompositeAdvance(bool bold, bool italic)
    {
        var model = new BuiltInDocumentModel();var section = new BuiltInSectionModel();
        foreach (var value in new[] { "ìB", "i\u0300B", "íB", "i\u0301B", "îB", "i\u0302B", "ïB", "i\u0308B" })
        {
            var paragraph = new BuiltInParagraphModel();
            paragraph.Inlines.Add(new BuiltInTextModel { Text = value,
                Format = { Bold = bold, Italic = italic, FontSizePoints = 20 } });
            section.Blocks.Add(paragraph);
        }
        model.Sections.Add(section);
        var text = Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model).Pages).Elements.OfType<OfdTextElement>().ToArray();
        for (var i = 0; i < text.Length; i += 2)
        {
            Assert.Equal(text[i].WidthMillimeters, text[i + 1].WidthMillimeters, 5);
            Assert.Equal(text[i].Runs[0].DeltaX, text[i + 1].Runs[0].DeltaX);
        }
    }

    [Theory]
    [InlineData("café")]
    [InlineData("cafe\u0301")]
    public void Native_AccentedLatinWord_ShouldMoveIntactToFreshLine(string word)
    {
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel();
        var paragraph = new BuiltInParagraphModel();
        paragraph.Inlines.Add(new BuiltInTextModel { Text = "prefix " });
        paragraph.Inlines.Add(new BuiltInTextModel { Text = word });
        section.Blocks.Add(paragraph); model.Sections.Add(section);
        OfdTextElement[] Render() => Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
                new List<DocxConversionDiagnostic>(), default).Render(model).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        var measured = Render();
        section.MarginLeftPoints = section.MarginRightPoints = 10;
        section.PageWidthPoints = (measured[0].WidthMillimeters + measured[1].WidthMillimeters / 2) * 72 / 25.4 + 20;
        var actual = Render();
        Assert.Equal(new[] { "prefix ", word }, actual.Select(value => value.Text));
        Assert.True(actual[1].YMillimeters > actual[0].YMillimeters);
    }

    [Theory]
    [InlineData("\u00A0")]
    [InlineData("\u202F")]
    [InlineData("\u2007")]
    public void Native_NonbreakingSpaceWithCrossSpanMark_ShouldStayConnected(string separator)
    {
        var referenceModel = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel();
        var paragraph = new BuiltInParagraphModel();
        paragraph.Inlines.Add(new BuiltInTextModel { Text = "X " });
        paragraph.Inlines.Add(new BuiltInTextModel { Text = "A" + separator });
        paragraph.Inlines.Add(new BuiltInTextModel { Text = "\u0301", Format = { Bold = true } });
        paragraph.Inlines.Add(new BuiltInTextModel { Text = "B" });
        section.Blocks.Add(paragraph);
        referenceModel.Sections.Add(section);
        var measured = Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(referenceModel).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        var unitWidth = measured.Skip(1).Sum(value => value.WidthMillimeters);
        section.MarginLeftPoints = section.MarginRightPoints = 10;
        section.PageWidthPoints = (measured[0].WidthMillimeters + unitWidth / 2) * 72 / 25.4 + 20;
        var actual = Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(referenceModel).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "X ", "A" + separator + "\u0301", "B" }, actual.Select(value => value.Text));
        Assert.True(actual[1].YMillimeters > actual[0].YMillimeters);
        Assert.Equal(actual[1].YMillimeters, actual[2].YMillimeters, 5);
        Assert.False(actual[1].Weight == OfdTextElement.BoldWeight);
    }

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, false)]
    [InlineData(true, true)]
    [InlineData(false, true)]
    public async Task Native_ConsecutiveTerminalNewlines_ShouldDeferBlankTailPages(bool explicitNative, bool table)
    {
        async Task<OfdDocumentPackage> Convert(bool followingContent, bool shortPage)
        {
            var following = !followingContent ? "" : table
                ? "<w:tbl><w:tblGrid><w:gridCol w:w=\"5000\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>"
                : "<w:p><w:r><w:t>B</w:t></w:r></w:p>";
            var height = shortPage ? 1077 : 6000;
            await using var input = CreateMinimalDocx($"""
                <w:p><w:r><w:t>A</w:t><w:br/><w:br/></w:r></w:p>{following}
                <w:sectPr><w:pgSz w:w="6000" w:h="{height}"/><w:pgMar w:top="283" w:bottom="283" w:left="200" w:right="200"/></w:sectPr>
                """);
            await using var output = new MemoryStream();
            var converter = explicitNative ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
                : new DocxToOfdConverter();
            await converter.ConvertAsync(input, output); output.Position = 0;
            return await new OfdReader().ReadAsync(output);
        }
        var terminal = await Convert(false, true);
        Assert.Equal("A", Assert.Single(Assert.Single(terminal.Pages).Elements.OfType<OfdTextElement>()).Text);
        var paged = (await Convert(true, true)).Pages;
        Assert.Equal(4, paged.Count);
        Assert.Equal("A", Assert.Single(paged[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Empty(paged[1].Elements);
        Assert.Empty(paged[2].Elements);
        var last = Assert.Single(paged[3].Elements.OfType<OfdTextElement>());
        Assert.Equal("B", last.Text);
        Assert.Equal(283 * 25.4 / 1440 + (table ? 0.8 : 0), last.YMillimeters, 2);
        var followed = Assert.Single((await Convert(true, false)).Pages).Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "A", "B" }, followed.Select(value => value.Text));
        Assert.Equal(10.5 * 25.4 / 72 * 1.3 * 3 + (table ? 0.8 : 0),
            followed[1].YMillimeters - followed[0].YMillimeters, 2);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void Native_NonbreakingSpaceMetrics_ShouldUseNarrowAndDigitAdvances(bool bold, bool italic)
    {
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel();
        foreach (var value in new[] { "A B", "A\u00A0B", "A\u202FB", "0\u20070", "000" })
        {
            var paragraph = new BuiltInParagraphModel();
            paragraph.Inlines.Add(new BuiltInTextModel { Text = value,
                Format = { Bold = bold, Italic = italic, FontSizePoints = 20 } });
            section.Blocks.Add(paragraph);
        }
        model.Sections.Add(section);
        var text = Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        double SeparatorAdvance(OfdTextElement value) => double.Parse(value.Runs[0].DeltaX!.Split(' ')[1],
            System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal(SeparatorAdvance(text[0]), SeparatorAdvance(text[1]), 5);
        Assert.Equal(20 * 25.4 / 72 * 0.2, SeparatorAdvance(text[2]), 5);
        Assert.True(SeparatorAdvance(text[2]) < SeparatorAdvance(text[0]));
        Assert.Equal(text[4].WidthMillimeters, text[3].WidthMillimeters, 5);
        Assert.Equal(SeparatorAdvance(text[4]), SeparatorAdvance(text[3]), 5);
    }

    [Theory]
    [InlineData("\u00A0", true)]
    [InlineData("\u202F", true)]
    [InlineData("\u2007", true)]
    [InlineData("\u00A0", false)]
    [InlineData("\u202F", false)]
    [InlineData("\u2007", false)]
    public async Task Native_NonbreakingSpaces_ShouldKeepJoinedTextOnOneLine(string separator, bool explicitNative)
    {
        // Choose a narrow width from the current Native font's measured advances,
        // so each CI OS exercises the break instead of relying on installed fonts.
        var referenceModel = new BuiltInDocumentModel();
        var referenceSection = new BuiltInSectionModel();
        var referenceParagraph = new BuiltInParagraphModel();
        referenceParagraph.Inlines.Add(new BuiltInTextModel { Text = "prefix " });
        referenceParagraph.Inlines.Add(new BuiltInTextModel { Text = "A" + separator + "B" });
        referenceSection.Blocks.Add(referenceParagraph);
        referenceModel.Sections.Add(referenceSection);
        var measured = Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(referenceModel).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        var usableWidth = measured[0].WidthMillimeters + measured[1].WidthMillimeters / 2;
        var pageWidthTwips = (int)Math.Round(usableWidth * 72 / 25.4 * 20 + 400);
        await using var input = CreateMinimalDocx($"""
            <w:p><w:r><w:t xml:space="preserve">prefix A{separator}B</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="{pageWidthTwips}" w:h="6000"/><w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>
            """);
        await using var output = new MemoryStream();
        var converter = explicitNative
            ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
            : new DocxToOfdConverter();
        await converter.ConvertAsync(input, output);
        output.Position = 0;
        var text = Assert.Single((await new OfdReader().ReadAsync(output)).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "prefix ", "A" + separator + "B" }, text.Select(value => value.Text));
        Assert.True(text[1].YMillimeters > text[0].YMillimeters);
        Assert.Equal("prefix A" + separator + "B", string.Concat(text.Select(value => value.Text)));
    }

    [Theory]
    [InlineData("center", true)]
    [InlineData("right", true)]
    [InlineData("center", false)]
    [InlineData("right", false)]
    public async Task Native_WrappedSeparatorSpaces_ShouldNotShiftVisibleAlignment(string alignment, bool explicitNative)
    {
        await using var input = CreateMinimalDocx($"""
            <w:p><w:pPr><w:jc w:val="{alignment}"/><w:ind w:firstLine="120"/></w:pPr><w:r><w:t>Alpha</w:t></w:r></w:p>
            <w:p><w:pPr><w:jc w:val="{alignment}"/><w:ind w:firstLine="120"/></w:pPr><w:r><w:t xml:space="preserve">Alpha  information</w:t></w:r></w:p>
            <w:sectPr><w:pgSz w:w="1800" w:h="6000"/><w:pgMar w:top="200" w:bottom="200" w:left="200" w:right="200"/></w:sectPr>
            """);
        await using var output = new MemoryStream();
        var converter = explicitNative
            ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native })
            : new DocxToOfdConverter();
        await converter.ConvertAsync(input, output);
        output.Position = 0;
        var text = Assert.Single((await new OfdReader().ReadAsync(output)).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "Alpha", "Alpha  ", "information" }, text.Select(value => value.Text));
        Assert.Equal(text[0].XMillimeters, text[1].XMillimeters, 5);
        Assert.True(text[2].YMillimeters > text[1].YMillimeters);
        Assert.Equal("Alpha  information", string.Concat(text.Skip(1).Select(value => value.Text)));
        // OFD serializes X and width independently to 0.001 mm.
        Assert.All(text, value => Assert.True(value.XMillimeters + value.WidthMillimeters <= 80 * 25.4 / 72 + 0.001));
    }

    [Theory]
    [InlineData((int)BuiltInParagraphAlignment.Center)]
    [InlineData((int)BuiltInParagraphAlignment.Right)]
    public void Native_StyledWrappedSeparators_ShouldRemainInsideLine(int alignmentValue)
    {
        var alignment = (BuiltInParagraphAlignment)alignmentValue;
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel
        {
            PageWidthPoints = 90, MarginLeftPoints = 10, MarginRightPoints = 10
        };
        var reference = new BuiltInParagraphModel();
        reference.Format.Alignment = alignment;
        reference.Inlines.Add(new BuiltInTextModel { Text = "Alpha" });
        section.Blocks.Add(reference);
        var wrapped = new BuiltInParagraphModel();
        wrapped.Format.Alignment = alignment;
        wrapped.Inlines.Add(new BuiltInTextModel { Text = "Alpha" });
        wrapped.Inlines.Add(new BuiltInTextModel { Text = "  ", Format = { ColorHex = "C00000" } });
        wrapped.Inlines.Add(new BuiltInTextModel { Text = "information" });
        section.Blocks.Add(wrapped);
        model.Sections.Add(section);
        var text = Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model).Pages)
            .Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "Alpha", "Alpha", "  ", "information" }, text.Select(value => value.Text));
        Assert.Equal(text[0].XMillimeters, text[1].XMillimeters, 5);
        Assert.Equal("Alpha  information", string.Concat(text.Skip(1).Select(value => value.Text)));
        Assert.All(text, value => Assert.True(value.XMillimeters + value.WidthMillimeters <= 80 * 25.4 / 72 + 0.000001));
    }

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

    [Fact]
    public async Task Native_ShouldPreserveConsecutiveExplicitPageBreaks()
    {
        await using var input = CreateMinimalDocx("""
            <w:p><w:r><w:t>A</w:t></w:r><w:r><w:br w:type="page"/><w:br w:type="page"/></w:r><w:r><w:t>B</w:t></w:r></w:p>
            """);
        await using var output = new MemoryStream();
        await new DocxToOfdConverter().ConvertAsync(input, output);
        output.Position = 0;
        var pages = (await new OfdReader().ReadAsync(output)).Pages;
        Assert.Equal(3, pages.Count);
        Assert.Equal("A", Assert.Single(pages[0].Elements.OfType<OfdTextElement>()).Text);
        Assert.Empty(pages[1].Elements.OfType<OfdTextElement>());
        Assert.Equal("B", Assert.Single(pages[2].Elements.OfType<OfdTextElement>()).Text);
    }

    [Fact]
    public async Task Native_ShouldKeepOversizeGlyphInNarrowTableCell()
    {
        await using var input = CreateMinimalDocx("""
            <w:tbl><w:tblGrid><w:gridCol w:w="100"/><w:gridCol w:w="9000"/></w:tblGrid>
              <w:tr>
                <w:tc><w:p><w:r><w:rPr><w:sz w:val="80"/></w:rPr><w:t>W</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>OK</w:t></w:r></w:p></w:tc>
              </w:tr>
            </w:tbl>
            """);
        await using var output = new MemoryStream();
        await new DocxToOfdConverter().ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Contains(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(), text => text.Text == "W");
        Assert.Contains(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(), text => text.Text == "OK");
    }

    [Fact]
    public void Native_ShouldCarryParagraphSpacingWithoutCreatingBlankPages()
    {
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel
        {
            PageWidthPoints = 200, PageHeightPoints = 65,
            MarginTopPoints = 10, MarginBottomPoints = 10,
            MarginLeftPoints = 10, MarginRightPoints = 10
        };
        var first = new BuiltInParagraphModel();
        first.Inlines.Add(new BuiltInTextModel { Text = "A" });
        first.Format.SpaceAfterPoints = 30;
        section.Blocks.Add(first);
        var second = new BuiltInParagraphModel();
        second.Inlines.Add(new BuiltInTextModel { Text = "B" });
        section.Blocks.Add(second);
        var third = new BuiltInParagraphModel();
        third.Inlines.Add(new BuiltInTextModel { Text = "C" });
        third.Format.SpaceBeforePoints = 50;
        section.Blocks.Add(third);
        model.Sections.Add(section);

        var package = new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model);
        Assert.Equal(3, package.Pages.Count);
        Assert.Equal("A", Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>()).Text);
        var b = Assert.Single(package.Pages[1].Elements.OfType<OfdTextElement>());
        Assert.Equal("B", b.Text);
        Assert.Equal(40 * 25.4 / 72, b.YMillimeters, 4);
        var c = Assert.Single(package.Pages[2].Elements.OfType<OfdTextElement>());
        Assert.Equal("C", c.Text);
        Assert.True(c.YMillimeters + c.HeightMillimeters <= package.Pages[2].HeightMillimeters - 10 * 25.4 / 72 + 0.000001);
    }

    [Fact]
    public void Native_ShouldDiscardTailSpacingAfterExplicitBreak()
    {
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel();
        var first = new BuiltInParagraphModel();
        first.Inlines.Add(new BuiltInTextModel { Text = "A" });
        first.Inlines.Add(new BuiltInBreakModel());
        first.Inlines.Add(new BuiltInBreakModel());
        first.Format.SpaceAfterPoints = 30;
        section.Blocks.Add(first);
        var second = new BuiltInParagraphModel();
        second.Inlines.Add(new BuiltInTextModel { Text = "B" });
        second.Format.PageBreakBefore = true;
        section.Blocks.Add(second);
        model.Sections.Add(section);
        var package = new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model);
        Assert.Equal(2, package.Pages.Count);
        Assert.Equal(72 * 25.4 / 72,
            Assert.Single(package.Pages[1].Elements.OfType<OfdTextElement>()).YMillimeters, 4);
    }

    [Fact]
    public void Native_ShouldCommitTrailingBreakWithOwningSectionPageSize()
    {
        var model = new BuiltInDocumentModel();
        var firstSection = new BuiltInSectionModel { PageWidthPoints = 300 };
        var first = new BuiltInParagraphModel();
        first.Inlines.Add(new BuiltInTextModel { Text = "A" });
        first.Inlines.Add(new BuiltInBreakModel { IsPageBreak = true });
        firstSection.Blocks.Add(first);
        model.Sections.Add(firstSection);
        var secondSection = new BuiltInSectionModel { PageWidthPoints = 500 };
        var second = new BuiltInParagraphModel();
        second.Inlines.Add(new BuiltInTextModel { Text = "B" });
        secondSection.Blocks.Add(second);
        model.Sections.Add(secondSection);

        var package = new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model);
        Assert.Equal(3, package.Pages.Count);
        Assert.Equal(300 * 25.4 / 72, package.Pages[1].WidthMillimeters, 4);
        Assert.Empty(package.Pages[1].Elements.OfType<OfdTextElement>());
        Assert.Equal(500 * 25.4 / 72, package.Pages[2].WidthMillimeters, 4);
        Assert.Equal("B", Assert.Single(package.Pages[2].Elements.OfType<OfdTextElement>()).Text);
    }

    [Fact]
    public void Native_ShouldKeepTerminalLineBreakBeforeFollowingParagraph()
    {
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel();
        var first = new BuiltInParagraphModel();
        first.Inlines.Add(new BuiltInTextModel { Text = "A" });
        first.Inlines.Add(new BuiltInBreakModel());
        section.Blocks.Add(first);
        var second = new BuiltInParagraphModel();
        second.Inlines.Add(new BuiltInTextModel { Text = "B" });
        section.Blocks.Add(second);
        model.Sections.Add(section);
        var package = new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(model);
        var text = Assert.Single(package.Pages).Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "A", "B" }, text.Select(value => value.Text));
        Assert.Equal((10.5 * 25.4 / 72) * 1.3 * 2,
            text[1].YMillimeters - text[0].YMillimeters, 4);

        var terminal = new BuiltInDocumentModel();
        var shortSection = new BuiltInSectionModel
        {
            PageHeightPoints = 45, MarginTopPoints = 10, MarginBottomPoints = 10
        };
        var last = new BuiltInParagraphModel();
        last.Inlines.Add(new BuiltInTextModel { Text = "A" });
        last.Inlines.Add(new BuiltInBreakModel());
        shortSection.Blocks.Add(last);
        terminal.Sections.Add(shortSection);
        Assert.Single(new BuiltInOfdRenderer(new DocxConversionOptions(),
            new List<DocxConversionDiagnostic>(), default).Render(terminal).Pages);
    }

    private static (int, int, int) Rgb(OfdColor color) => (color.Red, color.Green, color.Blue);

    private static OfdFontResource ResolveFont(OfdDocumentPackage package, OfdTextElement text)
    {
        if (!string.IsNullOrEmpty(text.FontResourceId))
        {
            var byId = package.Fonts.FirstOrDefault(font => font.Id == text.FontResourceId);
            if (byId is not null)
                return byId;
        }

        return Assert.Single(package.Fonts, font => font.FontName == text.FontName);
    }

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
        Assert.False(ResolveFont(package, normal).Bold);
        Assert.Equal((192, 0, 0), Rgb(emphasized.FillColor));
        Assert.True(ResolveFont(package, emphasized).Bold);
        Assert.Equal(OfdTextElement.DefaultWeight, normal.Weight);
        Assert.Equal(OfdTextElement.BoldWeight, emphasized.Weight);
        Assert.Equal(normal.YMillimeters, emphasized.YMillimeters);
        Assert.True(emphasized.XMillimeters >= normal.XMillimeters + normal.WidthMillimeters - 0.002);
        var subtitle = Assert.Single(text, t => t.Text.Contains("Generated DOCX"));
        Assert.True(ResolveFont(package, subtitle).Italic);
        Assert.True(subtitle.Italic);
        Assert.All(package.Fonts, font =>
        {
            Assert.False(string.IsNullOrWhiteSpace(font.FontName));
            var baseName = font.FontName.Split('|')[0];
            if (DocxFontCatalog.IsViewerLocalCjkFamily(font.FontName) ||
                baseName.StartsWith("Noto Sans CJK", StringComparison.OrdinalIgnoreCase))
            {
                Assert.Empty(font.Data);
            }
            else
            {
                Assert.NotEmpty(font.Data);
            }
        });
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
    /// DualLayer output must open in the same readers as Native output: the visual
    /// stage may not leak the PDF converter's "/2016" namespace or PDF metadata.
    /// </summary>
    [Theory]
    [InlineData(null)]
    [InlineData(OfdConstants.StandardNamespace)]
    public async Task GeneratedDocx_DualLayerShouldUseSameOfdNamespaceAsNative(string? configuredNamespace)
    {
        var options = new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn, OfdMode = DocxToOfdMode.DualLayer
        };
        if (configuredNamespace is not null) options.OfdNamespace = configuredNamespace;
        var expected = configuredNamespace ?? OfdConstants.Namespace;
        var converter = new DocxToOfdConverter(options, new PdfToOfdOptions());

        await using var output = new MemoryStream();
        await using (var input = File.OpenRead(ResolveGeneratedSample()))
        {
            await converter.ConvertAsync(input, output);
        }

        output.Position = 0;
        AssertPackageNamespace(output, expected);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("OFD-H", package.Options.DocType);
        Assert.Equal("DOCX document", package.Options.Metadata.Title);
        Assert.Equal("DualLayer", package.CustomTags["docx-ofd-mode"]);
    }

    /// <summary>
    /// A custom visual converter is staged as an OFD and read back; its manifests
    /// and raw fragments must still be rewritten into the DOCX namespace.
    /// </summary>
    [Fact]
    public async Task GeneratedDocx_DualLayerShouldRemapNamespaceOfStagedVisualOfd()
    {
        var options = new DocxConversionOptions
        {
            Engine = DocxConversionEngine.BuiltIn, OfdMode = DocxToOfdMode.DualLayer
        };
        var converter = new DocxToOfdConverter(
            new DocxToPdfConverter(options),
            new StagedPdfToOfdConverter(new PdfToOfdConverter(new PdfToOfdOptions
            {
                TextLayerMode = PdfTextLayerMode.None, Namespace = OfdConstants.StandardNamespace
            })));

        await using var output = new MemoryStream();
        await using (var input = File.OpenRead(ResolveGeneratedSample()))
        {
            await converter.ConvertAsync(input, output);
        }

        output.Position = 0;
        AssertPackageNamespace(output, OfdConstants.Namespace);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Pages.Count);
        Assert.All(package.Pages, page => Assert.Single(page.Elements.OfType<OfdImageElement>()));
    }

    private static void AssertPackageNamespace(Stream ofd, string expectedNamespace)
    {
        using var zip = new ZipArchive(ofd, ZipArchiveMode.Read, leaveOpen: true);
        var xmlEntries = zip.Entries.Where(entry => entry.FullName.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase)).ToList();
        Assert.Contains(xmlEntries, entry => entry.FullName == "Doc_0/DocumentRes.xml");
        foreach (var entry in xmlEntries)
        {
            using var stream = entry.Open();
            var root = XDocument.Load(stream).Root!;
            Assert.True(root.Name.NamespaceName == expectedNamespace,
                $"{entry.FullName} uses namespace '{root.Name.NamespaceName}', expected '{expectedNamespace}'.");
            Assert.True(root.DescendantsAndSelf().All(node => node.Name.NamespaceName == expectedNamespace),
                $"{entry.FullName} mixes namespaces.");
            Assert.Equal("ofd", root.GetPrefixOfNamespace(root.Name.Namespace));
        }
    }

    private sealed class StagedPdfToOfdConverter : Ofdrw.Net.Converter.Abstractions.Interfaces.IPdfToOfdConverter
    {
        private readonly PdfToOfdConverter _inner;

        public StagedPdfToOfdConverter(PdfToOfdConverter inner) => _inner = inner;

        public Task ConvertAsync(Stream pdfInput, Stream ofdOutput, System.Collections.Generic.IReadOnlyList<int>? pages = null,
            System.Threading.CancellationToken cancellationToken = default)
            => _inner.ConvertAsync(pdfInput, ofdOutput, pages, cancellationToken);
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
