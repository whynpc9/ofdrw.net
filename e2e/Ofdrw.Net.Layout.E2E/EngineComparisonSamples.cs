using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;
using Ofdrw.Net.Text.HarfBuzz;
using UglyToad.PdfPig;

// Issue #109 engine switch: run the existing visual-gate inputs through the compatibility layout and
// the opt-in paragraph layout engine, and record the page-by-page differences for review. The same
// DOCX, fonts and options are used for both; only the text engine changes.
internal static class EngineComparisonSamples
{
    internal static async Task Run(string directory, string baseline)
    {
        var fonts = PrepareFonts(directory, baseline);
        using (var sections = DocxReleaseFixtures.CreateMultiSectionDocx())
            File.WriteAllBytes(Path.Combine(directory, "headers-source.docx"), sections.ToArray());
        var image = DocxReleaseFixtures.CreateInlineJpegDocx("内嵌图片 / Native JPEG image fixture");
        using (image.Docx) File.WriteAllBytes(Path.Combine(directory, "image-source.docx"), image.Docx.ToArray());
        var records = new List<object>();
        foreach (var (scenario, source) in new[] { ("layout", baseline), ("headers", Path.Combine(directory, "headers-source.docx")),
                     ("image", Path.Combine(directory, "image-source.docx")) })
        {
            var outputs = new Dictionary<string, (OfdDocumentPackage Package, DocxToOfdConversionResult Result, string Ofd, string Pdf)>();
            foreach (var engine in new[] { "legacy", "engine" })
            {
                var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false, OfdMode = DocxToOfdMode.Native };
                options.FontDirectories.Add(fonts);
                if (engine == "engine") options.NativeParagraphLayoutEngine = HarfBuzzTextPrelayout.CreateParagraphLayout();
                var name = $"{scenario}-native-{engine}";
                var ofd = Path.Combine(directory, name + ".ofd");
                DocxToOfdConversionResult result;
                using (var input = File.OpenRead(source))
                using (var output = File.Create(ofd))
                    result = await new DocxToOfdConverter(options).ConvertWithResultAsync(input, output);
                OfdDocumentPackage package;
                using (var read = File.OpenRead(ofd)) package = await new OfdReader().ReadAsync(read);
                var pdf = Path.Combine(directory, name + ".pdf");
                using (var input = File.OpenRead(ofd))
                using (var output = File.Create(pdf))
                    await new OfdToPdfConverter().ConvertAsync(input, output);
                CheckPdfText(name, package, pdf);
                outputs[engine] = (package, result, ofd, pdf);
                Console.WriteLine($"{name}: {package.Pages.Count} pages");
            }
            records.Add(Compare(scenario, source, outputs["legacy"], outputs["engine"]));
        }
        foreach (var (scenario, flow) in new (string, Func<FlowTextEngine?, OfdDocumentPackage>)[] { ("flow", FlowSample), ("plain-text", PlainTextSample) })
        {
            var outputs = new Dictionary<string, (OfdDocumentPackage Package, DocxToOfdConversionResult? Result, string Ofd, string Pdf)>();
            foreach (var engine in new[] { "legacy", "engine" })
            {
                var package = flow(engine == "engine" ? FlowEngine(fonts) : null);
                var name = $"{scenario}-{engine}";
                var ofd = Path.Combine(directory, name + ".ofd");
                await using (var stream = File.Create(ofd)) await new OfdPackageWriter().WriteAsync(package, stream);
                using (var read = File.OpenRead(ofd)) package = await new OfdReader().ReadAsync(read);
                var pdf = Path.Combine(directory, name + ".pdf");
                using (var input = File.OpenRead(ofd))
                using (var output = File.Create(pdf))
                    await new OfdToPdfConverter().ConvertAsync(input, output);
                CheckPdfText(name, package, pdf);
                outputs[engine] = (package, null, ofd, pdf);
                Console.WriteLine($"{name}: {package.Pages.Count} pages");
            }
            records.Add(Compare(scenario, null, outputs["legacy"], outputs["engine"]));
        }
        var manifest = new
        {
            schemaVersion = 1, sourceSha = Environment.GetEnvironmentVariable("OFDRW_SOURCE_SHA"), generatedAtUtc = DateTimeOffset.UtcNow,
            runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
            os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
            chain = "DOCX → native OFD (legacy | ParagraphLayout) → PDF; Flow/plain text → OFD → PDF",
            fontFiles = Directory.GetFiles(fonts).OrderBy(path => path).Select(Identity).ToArray(),
            records,
        };
        File.WriteAllText(Path.Combine(directory, "engine-comparison-manifest.json"),
            JsonSerializer.Serialize(manifest, new JsonSerializerOptions { WriteIndented = true }));
    }

    private static object Compare<T>(string scenario, string? source, (OfdDocumentPackage Package, T Result, string Ofd, string Pdf) legacy,
        (OfdDocumentPackage Package, T Result, string Ofd, string Pdf) engine)
    {
        string[] PageText(OfdDocumentPackage package) => package.Pages.Select(page =>
            Compact(string.Concat(page.Elements.OfType<OfdTextElement>().Select(text => text.Text)))).ToArray();
        var legacyText = PageText(legacy.Package); var engineText = PageText(engine.Package);
        var sameText = string.Concat(legacyText) == string.Concat(engineText);
        if (!sameText) throw new InvalidDataException(scenario + ": the engines produced different logical text.");
        var diagnostics = engine.Result is DocxToOfdConversionResult result
            ? result.Diagnostics.Select(diagnostic => diagnostic.Code).Distinct().ToArray() : Array.Empty<string>();
        return new
        {
            scenario, source = source is null ? null : Identity(source),
            legacy = new { pages = legacy.Package.Pages.Count, ofd = Identity(legacy.Ofd), pdf = Identity(legacy.Pdf), textByPage = legacyText,
                fonts = legacy.Package.Fonts.Select(font => new { font.FontName, bytes = font.Data.Length }).ToArray() },
            engine = new { pages = engine.Package.Pages.Count, ofd = Identity(engine.Ofd), pdf = Identity(engine.Pdf), textByPage = engineText,
                fonts = engine.Package.Fonts.Select(font => new { font.FontName, bytes = font.Data.Length }).ToArray(), diagnostics },
            sameLogicalText = sameText, samePageCount = legacy.Package.Pages.Count == engine.Package.Pages.Count,
            samePageBreaks = legacyText.SequenceEqual(engineText),
        };
    }

    private static void CheckPdfText(string name, OfdDocumentPackage package, string pdf)
    {
        using var document = PdfDocument.Open(pdf);
        if (document.NumberOfPages != package.Pages.Count) throw new InvalidDataException(name + ": PDF page count differs.");
        for (var index = 0; index < package.Pages.Count; index++)
        {
            var expected = Compact(string.Concat(package.Pages[index].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
            if (Compact(document.GetPage(index + 1).Text) != expected)
                throw new InvalidDataException(name + ": PDF lost, reordered or repeated source text on page " + (index + 1));
        }
    }

    private static FlowTextEngine FlowEngine(string fonts)
    {
        TextFontFace Face(string name, string file) => new(name, File.ReadAllBytes(Path.Combine(fonts, file)));
        return new FlowTextEngine(HarfBuzzTextPrelayout.CreateParagraphLayout(), new[]
        {
            new TextFontFamily("Noto Sans", Face("Noto Sans", "NotoSans-Regular.ttf")),
            new TextFontFamily("Noto Sans CJK SC", Face("Noto Sans CJK SC", "Ofdrw-CI-NotoSansCJKsc-Regular.ttf")),
        });
    }

    // The default Flow generator's styled opening, shortened: title, local bold/italic/color, indented paragraphs.
    private static OfdDocumentPackage FlowSample(FlowTextEngine? engine)
    {
        var flow = new FlowDocument();
        flow.Options.MarginLeftMillimeters = 22; flow.Options.MarginRightMillimeters = 22;
        flow.Options.MarginTopMillimeters = 20; flow.Options.MarginBottomMillimeters = 20;
        flow.Options.TextEngine = engine;
        var title = new Paragraph("公开流式布局 / Public Flow Layout") { Alignment = ParagraphAlignment.Center, SpaceAfterMillimeters = 6 };
        title.Spans[0].FontSizeMillimeters = 6; title.Spans[0].Bold = true;
        flow.Blocks.Add(title);
        var styled = new Paragraph { SpaceAfterMillimeters = 4 };
        styled.Spans.Add(new Span("中文与 English typography 混排，普通正文。"));
        styled.Spans.Add(new Span("局部粗体 Bold") { Bold = true });
        styled.Spans.Add(new Span("；"));
        styled.Spans.Add(new Span("局部斜体 Italic") { Italic = true });
        styled.Spans.Add(new Span("；"));
        styled.Spans.Add(new Span("局部红色 Red") { Color = new OfdColor(192, 0, 0) });
        styled.Spans.Add(new Span("。后续文字恢复普通样式。"));
        flow.Blocks.Add(styled);
        for (var index = 1; index <= 30; index++)
            flow.Blocks.Add(new Paragraph($"第 {index:00} 段：病案示例不含真实患者信息。The proportional words information and minimum must wrap at natural line boundaries. 自动折行与分页保持完整文字。")
                { SpaceAfterMillimeters = 2, FirstLineIndentMillimeters = 4 });
        return flow.Render();
    }

    private static OfdDocumentPackage PlainTextSample(FlowTextEngine? engine)
    {
        var options = new PlainTextOfdOptions();
        options.Layout.TextEngine = engine;
        var text = "纯文本转换 / Plain text conversion\r\n\r\n" +
            string.Join("\n", Enumerable.Range(1, 40).Select(line => $"{line:00}\t项目 Item {line}\t数量 {line * 3}\t说明：中英文混排的纯文本行。")) +
            "\fSecond page after a form feed.\n";
        return new PlainTextOfdConverter(options).Render(text);
    }

    private static string PrepareFonts(string directory, string baseline)
    {
        const string archiveHash = "40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59";
        var archivePath = Path.GetFullPath(Path.Combine(Path.GetDirectoryName(baseline)!,
            "../../../tests/fixtures/trusted-fonts/sha256-" + archiveHash + ".zip"));
        if (Hash(File.ReadAllBytes(archivePath)) != archiveHash) throw new InvalidDataException("Trusted font bundle hash differs.");
        var target = Path.Combine(directory, "regular-fonts");
        Directory.CreateDirectory(target);
        using var archive = ZipFile.OpenRead(archivePath);
        foreach (var member in new[] { "fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf", "fonts/NotoSans-Regular.ttf" })
        {
            using var input = archive.GetEntry(member)!.Open(); using var bytes = new MemoryStream(); input.CopyTo(bytes);
            File.WriteAllBytes(Path.Combine(target, Path.GetFileName(member)), bytes.ToArray());
        }
        return target;
    }

    private static string Compact(string value) => new(value.Where(character => !char.IsWhiteSpace(character)).ToArray());
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
    private static object Identity(string path) => new { path = Path.GetFileName(path), bytes = new FileInfo(path).Length, sha256 = Hash(File.ReadAllBytes(path)) };
}
