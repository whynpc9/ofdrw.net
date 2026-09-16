using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed partial class DocxConversionTests
{
    [Fact]
    public async Task Native_ShouldUseFirstAvailableConfiguredFallbackAndEmbedItsBytes()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-fallback-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var fontPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(ResolveGeneratedSample())!,
                "../../Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
            var expectedFont = File.ReadAllBytes(fontPath);
            File.WriteAllBytes(Path.Combine(directory, "fixture.ttf"), expectedFont);
            var options = new DocxConversionOptions();
            options.FontDirectories.Add(directory);
            options.FontFallbackFamilies.Clear();
            options.FontFallbackFamilies.Add("Missing-ofdrw-font-" + Guid.NewGuid().ToString("N"));
            options.FontFallbackFamilies.Add("Ofdrw Test Face");
            options.FontFallbackFamilies.Add("SimSun");
            using var input = CreateMinimalDocx("<w:p><w:r><w:t>中文字体测试 Alpha</w:t></w:r></w:p>");
            using var output = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(input, output);
            output.Position = 0;
            var package = await new OfdReader().ReadAsync(output);
            var font = Assert.Single(package.Fonts);
            Assert.Equal("Ofdrw Test Face|regular", font.FontName);
            Assert.Equal(expectedFont, font.Data);
            Assert.Contains("中文字体测试 Alpha", new OfdTextExtractor().Extract(package));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task Native_ShouldDeclareViewerLocalCjkNamesWithoutEmbeddingASubstituteFace()
    {
        var metricFont = TryFindCjkMetricFont();
        if (metricFont is null)
            return;

        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-cjk-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.Copy(metricFont, Path.Combine(directory, Path.GetFileName(metricFont)));
            var options = new DocxConversionOptions();
            options.FontDirectories.Add(directory);
            using var input = CreateMinimalDocx("""
                <w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:b/><w:sz w:val="28"/></w:rPr><w:t>病案出库清单</w:t></w:r></w:p>
                <w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:sz w:val="18"/></w:rPr><w:t>扫描条数:50</w:t></w:r></w:p>
                """);
            using var output = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(input, output);
            output.Position = 0;
            var package = await new OfdReader().ReadAsync(output);
            var fonts = package.Fonts.ToList();
            var boldSong = Assert.Single(fonts, font => font.FontName == "SimSun" && font.Bold);
            var song = Assert.Single(fonts, font => font.FontName == "SimSun" && !font.Bold);
            Assert.Empty(boldSong.Data);
            Assert.Empty(song.Data);
            Assert.True(boldSong.Bold);
            Assert.False(boldSong.Italic);
            Assert.False(song.Bold);
            Assert.False(song.Italic);
            var title = Assert.Single(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(),
                element => element.Text == "病案出库清单");
            Assert.Equal("SimSun", title.FontName);
            var body = Assert.Single(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(),
                element => element.Text == "扫描条数:50");
            Assert.Equal("SimSun", body.FontName);
            Assert.Equal(boldSong.Id, title.FontResourceId);
            Assert.Equal(song.Id, body.FontResourceId);
            var advances = title.Runs.Single().DeltaX!.Split(' ')
                .Select(value => double.Parse(value, System.Globalization.CultureInfo.InvariantCulture))
                .ToArray();
            Assert.Equal(5, advances.Length);
            Assert.All(advances, advance =>
                Assert.InRange(advance, title.FontSizeMillimeters * 0.85, title.FontSizeMillimeters * 1.15));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task Native_ShouldKeepDistinctViewerLocalCjkStyleResources()
    {
        var metricFont = TryFindCjkMetricFont();
        if (metricFont is null)
            return;

        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-cjk-style-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            File.Copy(metricFont, Path.Combine(directory, Path.GetFileName(metricFont)));
            var options = new DocxConversionOptions();
            options.FontDirectories.Add(directory);
            using var input = CreateMinimalDocx("""
                <w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:sz w:val="22"/></w:rPr><w:t>常规宋体</w:t></w:r></w:p>
                <w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:i/><w:sz w:val="22"/></w:rPr><w:t>斜体宋体</w:t></w:r></w:p>
                <w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>粗体宋体</w:t></w:r></w:p>
                <w:p><w:r><w:rPr><w:rFonts w:ascii="黑体" w:eastAsia="黑体" w:hAnsi="黑体"/><w:b/><w:sz w:val="22"/></w:rPr><w:t>粗体黑体</w:t></w:r></w:p>
                """);
            using var output = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(input, output);
            output.Position = 0;
            var package = await new OfdReader().ReadAsync(output);
            var texts = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToList();
            var regular = ResolveFont(package, Assert.Single(texts, element => element.Text == "常规宋体"));
            var italic = ResolveFont(package, Assert.Single(texts, element => element.Text == "斜体宋体"));
            var songBold = ResolveFont(package, Assert.Single(texts, element => element.Text == "粗体宋体"));
            var heiBold = ResolveFont(package, Assert.Single(texts, element => element.Text == "粗体黑体"));
            Assert.Equal("SimSun", regular.FontName);
            Assert.False(regular.Bold);
            Assert.False(regular.Italic);
            Assert.Equal("SimSun", italic.FontName);
            Assert.False(italic.Bold);
            Assert.True(italic.Italic);
            Assert.NotEqual(regular.Id, italic.Id);
            Assert.Equal("SimSun", songBold.FontName);
            Assert.True(songBold.Bold);
            Assert.False(songBold.Italic);
            Assert.Equal("SimHei", heiBold.FontName);
            Assert.True(heiBold.Bold);
            Assert.False(heiBold.Italic);
            Assert.NotEqual(songBold.Id, heiBold.Id);
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Theory]
    [InlineData("SimSun")]
    [InlineData("宋体")]
    [InlineData("SimHei")]
    public async Task Native_CjkBold_ShouldRemainVisibleWithARegularSubstitute(string family)
    {
        using var source = CreateMinimalDocx($"""
            <w:p><w:r><w:rPr><w:rFonts w:ascii="{family}" w:eastAsia="{family}" w:hAnsi="{family}"/><w:b/><w:sz w:val="40"/></w:rPr><w:t>中文字体</w:t></w:r></w:p>
            """);
        using var ofd = new MemoryStream();
        await new DocxToOfdConverter(new DocxConversionOptions()).ConvertAsync(source, ofd);
        ofd.Position = 0;
        var package = await new OfdReader().ReadAsync(ofd);
        var font = Assert.Single(package.Fonts);
        Assert.True(font.Bold);
        // Model the loader's regular CJK substitute with an original deterministic
        // face. This tests actual export ink without depending on host font names.
        font.Data = File.ReadAllBytes(Path.GetFullPath(Path.Combine(Path.GetDirectoryName(ResolveGeneratedSample())!,
            "../../Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/style-metrics.ttf")));
        async Task<byte[]> Export()
        {
            using var input = new MemoryStream();
            await new Ofdrw.Net.Packaging.OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
            using var pdf = new MemoryStream();
            await new Ofdrw.Net.Converter.Pdf.Converters.OfdToPdfConverter().ConvertAsync(input, pdf);
            return pdf.ToArray();
        }
        var bold = await Export();
        font.Bold = false;
        var regular = await Export();
        using var semantic = PdfPigDocument.Open(bold);
        Assert.Equal("中文字体", semantic.GetPage(1).Text);
        static int Ink(byte[] pdf)
        {
            using var reader = Docnet.Core.DocLib.Instance.GetDocReader(pdf, new Docnet.Core.Models.PageDimensions(3d));
            using var page = reader.GetPageReader(0);
            var pixels = page.GetImage(); var count = 0;
            for (var i = 0; i < pixels.Length; i += 4)
                if (pixels[i + 3] > 128 && pixels[i] < 128 && pixels[i + 1] < 128 && pixels[i + 2] < 128) count++;
            return count;
        }
        Assert.True(Ink(bold) > Ink(regular) * 1.025, "CJK emphasis must survive a regular substitute font.");
    }

    [Fact]
    public async Task Native_ShouldUseEmAdvanceForCjkWhenCatalogHasOnlyALatinFace()
    {
        var directory = Path.Combine(Path.GetTempPath(), "ofdrw-latin-cjk-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try
        {
            var fontPath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(ResolveGeneratedSample())!,
                "../../Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
            File.Copy(fontPath, Path.Combine(directory, "narrow.ttf"));
            var options = new DocxConversionOptions();
            options.FontDirectories.Add(directory);
            using var input = CreateMinimalDocx("""
                <w:p><w:r><w:rPr><w:rFonts w:ascii="宋体" w:eastAsia="宋体" w:hAnsi="宋体"/><w:sz w:val="28"/></w:rPr><w:t>病案出库清单</w:t></w:r></w:p>
                """);
            using var output = new MemoryStream();
            await new DocxToOfdConverter(options).ConvertAsync(input, output);
            output.Position = 0;
            var package = await new OfdReader().ReadAsync(output);
            var title = Assert.Single(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>(),
                element => element.Text == "病案出库清单");
            Assert.Equal("SimSun", title.FontName);
            Assert.Empty(Assert.Single(package.Fonts, font => font.FontName == "SimSun").Data);
            var advances = title.Runs.Single().DeltaX!.Split(' ')
                .Select(value => double.Parse(value, System.Globalization.CultureInfo.InvariantCulture))
                .ToArray();
            Assert.Equal(5, advances.Length);
            Assert.All(advances, advance =>
                Assert.Equal(title.FontSizeMillimeters, advance, 3));
        }
        finally { Directory.Delete(directory, recursive: true); }
    }

    [Fact]
    public async Task Native_ShouldKeepInlineJpegContentDimensionsAndMediaType()
    {
        const string body = """
            <w:p><w:r><w:t>Native JPEG image fixture</w:t></w:r></w:p>
            <w:p><w:r><w:drawing><wp:inline><wp:extent cx="1440000" cy="720000"/><wp:docPr id="1" name="image"/>
            <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
            <pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="image.jpg"/><pic:cNvPicPr/></pic:nvPicPr>
            <pic:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>
            <pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1440000" cy="720000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>
            </pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>
            """;
        using var pixels = new Image<Rgba32>(80, 40);
        for (var y = 0; y < 40; y++)
        for (var x = 0; x < 80; x++) pixels[x, y] = x < 40 ? new Rgba32(255, 0, 0) : new Rgba32(0, 0, 255);
        using var jpeg = new MemoryStream(); pixels.SaveAsJpeg(jpeg); jpeg.Position = 0;
        using var input = CreateMinimalDocx(body);
        using (var document = WordprocessingDocument.Open(input, true))
            document.MainDocumentPart!.AddImagePart(ImagePartType.Jpeg, "rId2").FeedData(jpeg);
        input.Position = 0;
        using var output = new MemoryStream();
        await new DocxToOfdConverter().ConvertAsync(input, output); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var image = Assert.Single(package.Pages.SelectMany(page => page.Elements).OfType<OfdImageElement>());
        Assert.Equal("image/jpeg", image.MediaType);
        Assert.Equal(jpeg.ToArray(), image.Data);
        Assert.Equal(40, image.WidthMillimeters, 3);
        Assert.Equal(20, image.HeightMillimeters, 3);
        SaveFixtureArtifact("image-native", input, output);
    }

    [Fact]
    public async Task DualLayer_ShouldRejectUnreliableBodyMappingBeforeWritingOutput()
    {
        using var input = CreateMinimalDocx("<w:p><w:r><w:t>BODY ORIGINAL</w:t></w:r></w:p>");
        using var output = new MemoryStream();
        var converter = new DocxToOfdConverter(new FixturePdfRenderer("BODY OTHER"),
            new Ofdrw.Net.Converter.Pdf.Converters.PdfToOfdConverter());
        var exception = await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, output));
        Assert.Contains("DOCX_TEXT_PAGE_MAPPING_FAILED", exception.Message);
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task DualLayer_ShouldReportDocumentScopedNotesAndRejectAmbiguousPageSelection()
    {
        using var input = CreateMinimalDocx("<w:p><w:r><w:t>BODY ORIGINAL</w:t></w:r></w:p>");
        using (var document = WordprocessingDocument.Open(input, true))
            document.MainDocumentPart!.AddNewPart<WordprocessingCommentsPart>().Comments = new W.Comments(
                new W.Comment(new W.Paragraph(new W.Run(new W.Text("UNPLACED COMMENT")))) { Id = "1", Author = "Fixture" });
        input.Position = 0;
        var converter = new DocxToOfdConverter(new FixturePdfRenderer("BODY ORIGINAL"),
            new Ofdrw.Net.Converter.Pdf.Converters.PdfToOfdConverter());
        using var output = new MemoryStream();
        var result = await converter.ConvertWithResultAsync(input, output);
        Assert.True(result.OriginalTextPreserved);
        Assert.False(result.PageTextMappingAccurate);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_SUPPLEMENTAL_DOCUMENT_SCOPE");
        output.Position = 0;
        Assert.Contains("UNPLACED COMMENT", new OfdTextExtractor().Extract(await new OfdReader().ReadAsync(output)));
        input.Position = 0;
        using var selected = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => converter.ConvertAsync(input, selected, [0]));
        Assert.Equal(0, selected.Length);
    }

    [Fact]
    public async Task NativeInputBudget_ShouldStopBeforeStagingTheWholeInput()
    {
        using var input = new Ofdrw.Net.TestSupport.NonSeekableInput(new byte[100]);
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new DocxToOfdConverter(new DocxConversionOptions { MaxInputBytes = 10 }).ConvertAsync(input, output));
        Assert.Equal(11, input.BytesRead);
        Assert.Equal(0, output.Length);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task NativeAndDefault_ShouldPreserveSupplementaryOriginalTextAndReturnDiagnostics(bool explicitNative)
    {
        using var source = CreateSupplementalDocx();
        using var output = new MemoryStream();
        var options = new DocxConversionOptions { Engine = DocxConversionEngine.BuiltIn };
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, output);
        Assert.True(result.OriginalTextPreserved);
        Assert.Null(result.ActualEngine);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_SUPPLEMENTAL_TEXT_APPENDED");
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var text = new OfdTextExtractor().Extract(package);
        foreach (var phrase in new[] { "BODY ORIGINAL", "脚注原文 Footnote body", "尾注原文 Endnote body", "批注原文 Comment body" })
            Assert.Contains(phrase.Replace(" ", ""), new string(text.Where(character => !char.IsWhiteSpace(character)).ToArray()));
        Assert.Equal("labeled-after-body", package.CustomTags["source-text-supplemental-layout"]);
        SaveFixtureArtifact(explicitNative ? "supplemental-native" : "supplemental-default", source, output);
    }

    [Fact]
    public async Task Native_ShouldRepeatHeadersFootersAndHonorSectionNumbering()
    {
        using var source = CreateMultiSectionDocx();
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter().ConvertWithResultAsync(source, output);
        Assert.Equal(5, result.SourcePageCount);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        var texts = new OfdTextExtractor().ExtractPages(package);
        var headers = new[] { "First Header", "Even Header", "Odd Header", "Even Header", "Odd Header" };
        var numbers = new[] { 1, 2, 3, 10, 11 };
        for (var index = 0; index < 5; index++)
        {
            Assert.Contains(headers[index], texts[index]);
            Assert.Contains($"Page{numbers[index]}/5", texts[index].Replace(" ", "").Replace("\n", ""));
            var header = package.Pages[index].Elements.OfType<OfdTextElement>().First(text => text.Text.Contains("Header"));
            var body = package.Pages[index].Elements.OfType<OfdTextElement>().First(text => text.Text.Contains("Body"));
            var footer = package.Pages[index].Elements.OfType<OfdTextElement>().First(text => text.Text.Contains("Page"));
            Assert.True(header.YMillimeters + header.HeightMillimeters < body.YMillimeters);
            Assert.True(footer.YMillimeters > 260);
            Assert.True(footer.YMillimeters + footer.HeightMillimeters < package.Pages[index].HeightMillimeters);
        }
        SaveFixtureArtifact("headers-native", source, output);
    }

    [Fact]
    public async Task Native_ShouldPreservePageOrderDuplicatesAndInvalidSelectionCompatibility()
    {
        using var source = CreateMultiSectionDocx();
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter().ConvertWithResultAsync(source, output, [2, 0, 2, -1, 99]);
        Assert.Equal(new[] { 2, 0, 2 }, result.SourcePages);
        output.Position = 0;
        var texts = new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(output));
        Assert.Contains("Body three", texts[0]);
        Assert.Contains("Body one", texts[1]);
        Assert.Equal(texts[0], texts[2]);
        source.Position = 0;
        using var all = new MemoryStream();
        var fallback = await new DocxToOfdConverter().ConvertWithResultAsync(source, all, [-1, 999]);
        Assert.Equal(new[] { 0, 1, 2, 3, 4 }, fallback.SourcePages);
        SaveFixtureArtifact("selected-native", source, output);
    }

    [Fact]
    public async Task DualLayer_ShouldMapAutomaticPaginationToOriginalTextBeforeSelectingPages()
    {
        using var source = CreateAutoPaginatedDocx();
        var options = new DocxConversionOptions { Engine = DocxConversionEngine.BuiltIn, OfdMode = DocxToOfdMode.DualLayer };
        using var pdfOutput = new MemoryStream();
        await new DocxToPdfConverter(options).ConvertAsync(source, pdfOutput);
        pdfOutput.Position = 0;
        using var pdf = PdfPigDocument.Open(pdfOutput);
        Assert.True(pdf.NumberOfPages >= 3);
        var expected = Enumerable.Range(1, pdf.NumberOfPages).Select(page => RowMarkers(pdf.GetPage(page).Text)).ToList();
        source.Position = 0;
        using var full = new MemoryStream();
        var fullResult = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, full);
        Assert.True(fullResult.OriginalTextPreserved);
        Assert.True(fullResult.PageTextMappingAccurate);
        full.Position = 0;
        var fullTexts = new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(full));
        for (var page = 0; page < expected.Count; page++) Assert.Equal(expected[page], RowMarkers(fullTexts[page]));
        source.Position = 0;
        using var selected = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(source, selected, [1, 0, 1]);
        Assert.Equal(new[] { 1, 0, 1 }, result.SourcePages);
        selected.Position = 0;
        var actual = new OfdTextExtractor().ExtractPages(await new OfdReader().ReadAsync(selected));
        Assert.Equal(expected[1], RowMarkers(actual[0]));
        Assert.Equal(expected[0], RowMarkers(actual[1]));
        Assert.Equal(actual[0], actual[2]);
        SaveFixtureArtifact("automatic-dual-layer", source, full);
        SaveFixtureArtifact("selected-dual-layer", source, selected);
    }

    [Fact]
    public async Task Native_ShouldExposeKnownTextLossUnderPlaceholderPolicy()
    {
        using var input = CreateMinimalDocx("<w:p><w:r><w:footnoteReference w:id='99'/></w:r></w:p>");
        using var output = new MemoryStream();
        var result = await new DocxToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.False(result.OriginalTextPreserved);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_FOOTNOTE_UNSUPPORTED");
    }

    private static string[] RowMarkers(string text) => Regex.Matches(text, "ROW-[0-9]{2}").Select(match => match.Value).ToArray();

    private static MemoryStream CreateAutoPaginatedDocx()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            var body = new W.Body();
            for (var index = 0; index < 45; index++) body.Append(new W.Paragraph(
                new W.ParagraphProperties(new W.SpacingBetweenLines { After = "0" }),
                new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Arial", HighAnsi = "Arial" }, new W.FontSize { Val = "24" }),
                    new W.Text($"ROW-{index:00} original paragraph."))));
            body.Append(new W.SectionProperties(new W.PageSize { Width = 6000, Height = 6000 },
                new W.PageMargin { Top = 400, Bottom = 400, Left = 400, Right = 400 }));
            main.Document = new W.Document(body);
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateSupplementalDocx()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<FootnotesPart>().Footnotes = new W.Footnotes(
                new W.Footnote(new W.Paragraph(new W.Run(new W.Text("脚注原文 Footnote body")))) { Id = 1 });
            main.AddNewPart<EndnotesPart>().Endnotes = new W.Endnotes(
                new W.Endnote(new W.Paragraph(new W.Run(new W.Text("尾注原文 Endnote body")))) { Id = 2 });
            main.AddNewPart<WordprocessingCommentsPart>().Comments = new W.Comments(
                new W.Comment(new W.Paragraph(new W.Run(new W.Text("批注原文 Comment body")))) { Id = "3", Author = "Fixture" });
            main.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text("BODY ORIGINAL "),
                new W.FootnoteReference { Id = 1 }, new W.EndnoteReference { Id = 2 }, new W.CommentReference { Id = "3" })), new W.SectionProperties()));
        }
        stream.Position = 0;
        return stream;
    }

    private static MemoryStream CreateMultiSectionDocx()
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var main = document.AddMainDocumentPart();
            main.AddNewPart<DocumentSettingsPart>().Settings = new W.Settings(new W.EvenAndOddHeaders());
            var odd = main.AddNewPart<HeaderPart>(); odd.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("Odd Header"))));
            var even = main.AddNewPart<HeaderPart>(); even.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("Even Header"))));
            var first = main.AddNewPart<HeaderPart>(); first.Header = new W.Header(new W.Paragraph(new W.Run(new W.Text("First Header"))));
            var footer = main.AddNewPart<FooterPart>();
            footer.Footer = new W.Footer(new W.Paragraph(new W.ParagraphProperties(new W.Justification { Val = W.JustificationValues.Center }),
                new W.Run(new W.Text("Page ")), new W.SimpleField(new W.Run(new W.Text("0"))) { Instruction = "PAGE" },
                new W.Run(new W.Text("/")), new W.SimpleField(new W.Run(new W.Text("0"))) { Instruction = "NUMPAGES" }));
            var firstSection = new W.SectionProperties(
                new W.HeaderReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(odd) },
                new W.HeaderReference { Type = W.HeaderFooterValues.Even, Id = main.GetIdOfPart(even) },
                new W.HeaderReference { Type = W.HeaderFooterValues.First, Id = main.GetIdOfPart(first) },
                new W.FooterReference { Type = W.HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) },
                new W.FooterReference { Type = W.HeaderFooterValues.Even, Id = main.GetIdOfPart(footer) },
                new W.FooterReference { Type = W.HeaderFooterValues.First, Id = main.GetIdOfPart(footer) },
                new W.PageNumberType { Start = 1 }, new W.TitlePage());
            var body = new W.Body();
            foreach (var (label, index) in new[] { ("one", 0), ("two", 1), ("three", 2) })
            {
                var properties = new W.ParagraphProperties();
                if (index > 0) properties.Append(new W.PageBreakBefore());
                if (index == 2) properties.Append(firstSection);
                body.Append(new W.Paragraph(properties, new W.Run(new W.Text("Body " + label))));
            }
            body.Append(new W.Paragraph(new W.Run(new W.Text("Body four"))),
                new W.Paragraph(new W.ParagraphProperties(new W.PageBreakBefore()), new W.Run(new W.Text("Body five"))),
                new W.SectionProperties(new W.PageNumberType { Start = 10 }));
            main.Document = new W.Document(body);
        }
        stream.Position = 0;
        return stream;
    }

    private static string? TryFindCjkMetricFont()
    {
        foreach (var directory in DocxFontCatalog.DefaultPlatformFontDirectories())
        {
            foreach (var name in new[]
                     {
                         "simhei.ttf",
                         "Ofdrw-CI-NotoSansCJKsc-Regular.ttf",
                         "NotoSansCJKsc-Regular.ttf"
                     })
            {
                var path = Path.Combine(directory, name);
                if (File.Exists(path)) return path;
            }

            try
            {
                var match = Directory.GetFiles(directory, "Ofdrw-CI-Noto*.ttf");
                if (match.Length > 0) return match[0];
            }
            catch (IOException)
            {
            }
        }

        return null;
    }

    private static void SaveFixtureArtifact(string name, MemoryStream docx, MemoryStream ofd)
    {
        var directory = Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS");
        if (string.IsNullOrEmpty(directory)) return;
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, name + ".docx"), docx.ToArray());
        File.WriteAllBytes(Path.Combine(directory, name + ".ofd"), ofd.ToArray());
    }

    private sealed class FixturePdfRenderer(string text) : Ofdrw.Net.Converter.Abstractions.Interfaces.IDocxToPdfConverter
    {
        public Task ConvertAsync(Stream input, Stream output, CancellationToken cancellationToken = default)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Ofdrw.Net.Converter.Pdf.PdfFontRegistry.EnsureInstalled();
            using var document = new PdfSharpCore.Pdf.PdfDocument();
            var page = document.AddPage(); page.Width = 595; page.Height = 842;
            using (var graphics = PdfSharpCore.Drawing.XGraphics.FromPdfPage(page))
                graphics.DrawString(text, new PdfSharpCore.Drawing.XFont("Arial", 12), PdfSharpCore.Drawing.XBrushes.Black, new PdfSharpCore.Drawing.XPoint(72, 90));
            document.Save(output, false);
            return Task.CompletedTask;
        }
    }
}
