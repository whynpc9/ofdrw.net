using System.Security.Cryptography;
using System.IO.Compression;
using System.Text;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

var outputDirectory = args.Length > 0 ? Path.GetFullPath(args[0]) :
    Path.GetFullPath("artifacts/flow-layout");
var docxPath = args.Length > 1 ? Path.GetFullPath(args[1]) :
    Path.GetFullPath("e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx");
Directory.CreateDirectory(outputDirectory);

var flow = new FlowDocument();
flow.Options.MarginLeftMillimeters = 22;
flow.Options.MarginRightMillimeters = 22;
flow.Options.MarginTopMillimeters = 20;
flow.Options.MarginBottomMillimeters = 20;
var title = new Paragraph("公开流式布局 / Public Flow Layout")
{
    Alignment = ParagraphAlignment.Center,
    SpaceAfterMillimeters = 6
};
title.Spans[0].FontSizeMillimeters = 6;
title.Spans[0].Bold = true;
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
for (var index = 1; index <= 55; index++)
{
    flow.Blocks.Add(new Paragraph($"第 {index:00} 段：病案示例不含真实患者信息。The proportional words information and minimum must wrap at natural line boundaries. 自动折行与分页保持完整文字。")
    {
        SpaceAfterMillimeters = 2,
        FirstLineIndentMillimeters = 4
    });
}

var flowPackage = flow.Render();
var expectedFlowText = Compact(string.Concat(flow.Blocks.Cast<Paragraph>()
    .SelectMany(paragraph => paragraph.Spans).Select(span => span.Text)));
await using (var stream = File.Create(Path.Combine(outputDirectory, "flow-public.ofd")))
    await new OfdPackageWriter().WriteAsync(flowPackage, stream);

// Larger text makes a separator-space alignment regression visible. Each Alpha
// reference must share its horizontal position with Alpha on the following wrapped line.
var alignmentFlow = new FlowDocument();
alignmentFlow.Options.PageWidthMillimeters = 120;
alignmentFlow.Options.MarginLeftMillimeters = alignmentFlow.Options.MarginRightMillimeters = 20;
alignmentFlow.Options.DefaultFontSizeMillimeters = 6;
foreach (var alignment in new[] { ParagraphAlignment.Center, ParagraphAlignment.Right })
{
    alignmentFlow.Blocks.Add(new Paragraph(alignment + " / reference then wrapped") { SpaceBeforeMillimeters = 8 });
    foreach (var value in new[] { "Alpha", "Alpha  information" })
        alignmentFlow.Blocks.Add(new Paragraph(value)
        {
            Alignment = alignment, LeftIndentMillimeters = 25, RightIndentMillimeters = 20,
            FirstLineIndentMillimeters = 2
        });
}
alignmentFlow.Blocks.Add(new Paragraph("NBSP / joined A B") { SpaceBeforeMillimeters = 8 });
foreach (var separator in new[] { "\u00A0", "\u202F", "\u2007" })
    alignmentFlow.Blocks.Add(new Paragraph("prefix A" + separator + "B")
    { LeftIndentMillimeters = 25, RightIndentMillimeters = 30 });
await using (var stream = File.Create(Path.Combine(outputDirectory, "alignment-public.ofd")))
    await new OfdPackageWriter().WriteAsync(alignmentFlow.Render(), stream);

var alignmentDocxPath = Path.Combine(outputDirectory, "alignment.docx");
using (var docx = File.Create(alignmentDocxPath))
using (var zip = new ZipArchive(docx, ZipArchiveMode.Create))
{
    void WritePart(string path, string xml)
    {
        using var writer = new StreamWriter(zip.CreateEntry(path).Open(), new UTF8Encoding(false));
        writer.Write(xml);
    }
    WritePart("[Content_Types].xml", """
        <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
        """);
    WritePart("_rels/.rels", """
        <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
        """);
    var body = new StringBuilder();
    foreach (var alignment in new[] { "center", "right" })
    {
        body.Append($"<w:p><w:r><w:t>{alignment} / reference then wrapped</w:t></w:r></w:p>");
        foreach (var value in new[] { "Alpha", "Alpha  information" })
            body.Append($"<w:p><w:pPr><w:jc w:val=\"{alignment}\"/><w:ind w:left=\"1417\" w:right=\"1134\" w:firstLine=\"113\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:sz w:val=\"34\"/></w:rPr><w:t xml:space=\"preserve\">{value}</w:t></w:r></w:p>");
    }
    body.Append("<w:p><w:r><w:t>NBSP / joined A B</w:t></w:r></w:p>");
    foreach (var separator in new[] { "\u00A0", "\u202F", "\u2007" })
        body.Append($"<w:p><w:pPr><w:ind w:left=\"1417\" w:right=\"1701\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:sz w:val=\"34\"/></w:rPr><w:t xml:space=\"preserve\">prefix A{separator}B</w:t></w:r></w:p>");
    WritePart("word/document.xml", $"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"6803\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:bottom=\"1134\" w:left=\"1134\" w:right=\"1134\"/></w:sectPr></w:body></w:document>");
}

async Task ConvertDocx(string name, DocxConversionOptions? options, string? source = null)
{
    await using var input = File.OpenRead(source ?? docxPath);
    await using var output = File.Create(Path.Combine(outputDirectory, name + ".ofd"));
    var converter = options is null ? new DocxToOfdConverter() : new DocxToOfdConverter(options);
    await converter.ConvertAsync(input, output);
}

await ConvertDocx("docx-native", new DocxConversionOptions { OfdMode = DocxToOfdMode.Native });
await ConvertDocx("docx-default", null);
await ConvertDocx("alignment-native", new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }, alignmentDocxPath);
await ConvertDocx("alignment-default", null, alignmentDocxPath);

string? nativeText = null;
string? alignmentNativeText = null;
foreach (var name in new[] { "flow-public", "docx-native", "docx-default", "alignment-public", "alignment-native", "alignment-default" })
{
    var ofdPath = Path.Combine(outputDirectory, name + ".ofd");
    var pdfPath = Path.Combine(outputDirectory, name + ".pdf");
    await using (var input = File.OpenRead(ofdPath))
    await using (var output = File.Create(pdfPath))
        await new OfdToPdfConverter().ConvertAsync(input, output);
    await using var read = File.OpenRead(ofdPath);
    var package = await new OfdReader().ReadAsync(read);
    var compactText = Compact(string.Concat(package.Pages.SelectMany(page => page.Elements)
        .OfType<OfdTextElement>().Select(element => element.Text)));
    if (name == "flow-public" && compactText != expectedFlowText)
        throw new InvalidOperationException("Public Flow lost or duplicated source text.");
    if (name == "docx-native") nativeText = compactText;
    if (name == "docx-default" && compactText != nativeText)
        throw new InvalidOperationException("Default DOCX output differs from explicit Native text.");
    if (name.StartsWith("alignment-"))
    {
        var alpha = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
            .Where(element => element.Text.StartsWith("Alpha")).ToArray();
        if (alpha.Length != 4 || alpha[1].Text != "Alpha  " || alpha[3].Text != "Alpha  " ||
            Math.Abs(alpha[0].XMillimeters - alpha[1].XMillimeters) > 0.00001 ||
            Math.Abs(alpha[2].XMillimeters - alpha[3].XMillimeters) > 0.00001)
            throw new InvalidOperationException("Wrapped separators changed visible alignment.");
        foreach (var separator in new[] { "\u00A0", "\u202F", "\u2007" })
        {
            var joined = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
                .Single(element => element.Text == "A" + separator + "B");
            var advances = joined.Runs[0].DeltaX!.Split(' ');
            if (double.Parse(advances[1], System.Globalization.CultureInfo.InvariantCulture) <= 0)
                throw new InvalidOperationException("A nonbreaking separator lost its advance.");
        }
        if (name == "alignment-native") alignmentNativeText = compactText;
        if (name == "alignment-default" && compactText != alignmentNativeText)
            throw new InvalidOperationException("Default alignment text differs from explicit Native.");
    }
    Console.WriteLine($"{name}: pages={package.Pages.Count}, chars={compactText.Length}, ofdBytes={new FileInfo(ofdPath).Length}, pdfBytes={new FileInfo(pdfPath).Length}, sha256={Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ofdPath)))}");
}

static string Compact(string text) => new(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
