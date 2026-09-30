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
alignmentFlow.Blocks.Add(new Paragraph("Latin word / café") { SpaceBeforeMillimeters = 8 });
foreach (var word in new[] { "café", "cafe\u0301" })
    alignmentFlow.Blocks.Add(new Paragraph("prefix " + word)
    { LeftIndentMillimeters = 25, RightIndentMillimeters = 30 });
var iLabel = new Paragraph("Latin i / mínimo") { SpaceBeforeMillimeters = 4 };
iLabel.Spans[0].FontSizeMillimeters = 4;
alignmentFlow.Blocks.Add(iLabel);
foreach (var word in new[] { "mínimo", "mi\u0301nimo" })
{
    var paragraph = new Paragraph("prefix " + word) { LeftIndentMillimeters = 25, RightIndentMillimeters = 30 };
    paragraph.Spans[0].FontSizeMillimeters = 4.5;
    alignmentFlow.Blocks.Add(paragraph);
}
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
    body.Append("<w:p><w:r><w:t>Latin word / café</w:t></w:r></w:p>");
    foreach (var word in new[] { "café", "cafe\u0301" })
        body.Append($"<w:p><w:pPr><w:ind w:left=\"1417\" w:right=\"1701\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:sz w:val=\"34\"/></w:rPr><w:t xml:space=\"preserve\">prefix {word}</w:t></w:r></w:p>");
    body.Append("<w:p><w:r><w:t>Latin i / mínimo</w:t></w:r></w:p>");
    foreach (var word in new[] { "mínimo", "mi\u0301nimo" })
        body.Append($"<w:p><w:pPr><w:ind w:left=\"1417\" w:right=\"1701\"/></w:pPr><w:r><w:rPr><w:rFonts w:ascii=\"Arial\"/><w:sz w:val=\"25\"/></w:rPr><w:t xml:space=\"preserve\">prefix {word}</w:t></w:r></w:p>");
    WritePart("word/document.xml", $"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{body}<w:sectPr><w:pgSz w:w=\"6803\" w:h=\"16838\"/><w:pgMar w:top=\"1134\" w:bottom=\"1134\" w:left=\"1134\" w:right=\"1134\"/></w:sectPr></w:body></w:document>");
}

var terminalFlow = new FlowDocument();
terminalFlow.Options.PageWidthMillimeters = 106;
terminalFlow.Options.PageHeightMillimeters = 19;
terminalFlow.Options.MarginTopMillimeters = terminalFlow.Options.MarginBottomMillimeters = 5;
terminalFlow.Options.MarginLeftMillimeters = terminalFlow.Options.MarginRightMillimeters = 5;
terminalFlow.Blocks.Add(new Paragraph("A\n\n"));
await using (var stream = File.Create(Path.Combine(outputDirectory, "terminal-public.ofd")))
    await new OfdPackageWriter().WriteAsync(terminalFlow.Render(), stream);
terminalFlow.Blocks.Add(new Paragraph("B"));
await using (var stream = File.Create(Path.Combine(outputDirectory, "terminal-followed-public.ofd")))
    await new OfdPackageWriter().WriteAsync(terminalFlow.Render(), stream);

// Reuse the minimal OPC shell for a one-line page with two trailing Word breaks.
var terminalDocxPath = Path.Combine(outputDirectory, "terminal.docx");
File.Copy(alignmentDocxPath, terminalDocxPath, overwrite: true);
using (var zip = ZipFile.Open(terminalDocxPath, ZipArchiveMode.Update))
{
    zip.GetEntry("word/document.xml")!.Delete();
    using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
    writer.Write("""
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>A</w:t><w:br/><w:br/></w:r></w:p><w:sectPr><w:pgSz w:w="6000" w:h="1077"/><w:pgMar w:top="283" w:bottom="283" w:left="283" w:right="283"/></w:sectPr></w:body></w:document>
        """);
}

var followedDocxPath = Path.Combine(outputDirectory, "terminal-followed.docx");
File.Copy(terminalDocxPath, followedDocxPath, overwrite: true);
using (var zip = ZipFile.Open(followedDocxPath, ZipArchiveMode.Update))
{
    var entry = zip.GetEntry("word/document.xml")!;
    string xml;
    using (var reader = new StreamReader(entry.Open())) xml = reader.ReadToEnd();
    entry.Delete();
    using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
    writer.Write(xml.Replace("<w:sectPr>", "<w:p><w:r><w:t>B</w:t></w:r></w:p><w:sectPr>"));
}

var styledNewlineFlow = new FlowDocument();
styledNewlineFlow.Options.PageWidthMillimeters = 120;
var large = new Paragraph();
large.Spans.Add(new Span("A\n\n") { FontSizeMillimeters = 20 });
styledNewlineFlow.Blocks.Add(large);
styledNewlineFlow.Blocks.Add(new Paragraph("B"));
styledNewlineFlow.Blocks.Add(new Paragraph("Mixed newline / 20 mm then 2 mm") { SpaceBeforeMillimeters = 8 });
var mixed = new Paragraph("C");
mixed.Spans.Add(new Span("\n") { FontSizeMillimeters = 20 });
mixed.Spans.Add(new Span("\n") { FontSizeMillimeters = 2 });
styledNewlineFlow.Blocks.Add(mixed);
styledNewlineFlow.Blocks.Add(new Paragraph("D"));
await using (var stream = File.Create(Path.Combine(outputDirectory, "styled-newline-public.ofd")))
    await new OfdPackageWriter().WriteAsync(styledNewlineFlow.Render(), stream);
var styledDocxPath = Path.Combine(outputDirectory, "styled-newline.docx");
File.Copy(terminalDocxPath, styledDocxPath, overwrite: true);
using (var zip = ZipFile.Open(styledDocxPath, ZipArchiveMode.Update))
{
    zip.GetEntry("word/document.xml")!.Delete();
    using var writer = new StreamWriter(zip.CreateEntry("word/document.xml").Open(), new UTF8Encoding(false));
    writer.Write("""
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:sz w:val="113"/></w:rPr><w:t>A</w:t><w:br/><w:br/></w:r></w:p><w:p><w:r><w:t>B</w:t></w:r></w:p><w:p><w:r><w:t>Mixed newline / 40 pt then 4 pt</w:t></w:r></w:p><w:p><w:r><w:t>C</w:t></w:r><w:r><w:rPr><w:sz w:val="80"/></w:rPr><w:br/></w:r><w:r><w:rPr><w:sz w:val="8"/></w:rPr><w:cr/></w:r></w:p><w:p><w:r><w:t>D</w:t></w:r></w:p><w:sectPr><w:pgSz w:w="6803" w:h="16838"/><w:pgMar w:top="1440" w:bottom="1440" w:left="1440" w:right="1440"/></w:sectPr></w:body></w:document>
        """);
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
await ConvertDocx("terminal-native", new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }, terminalDocxPath);
await ConvertDocx("terminal-default", null, terminalDocxPath);
await ConvertDocx("terminal-followed-native", new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }, followedDocxPath);
await ConvertDocx("terminal-followed-default", null, followedDocxPath);
await ConvertDocx("styled-newline-native", new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }, styledDocxPath);
await ConvertDocx("styled-newline-default", null, styledDocxPath);

string? nativeText = null;
string? alignmentNativeText = null;
foreach (var name in new[] { "flow-public", "docx-native", "docx-default", "alignment-public", "alignment-native", "alignment-default",
    "terminal-public", "terminal-native", "terminal-default",
    "terminal-followed-public", "terminal-followed-native", "terminal-followed-default",
    "styled-newline-public", "styled-newline-native", "styled-newline-default" })
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
    if (name.StartsWith("terminal-followed-") && (package.Pages.Count != 4 || compactText != "AB" ||
        package.Pages[1].Elements.Count != 0 || package.Pages[2].Elements.Count != 0))
        throw new InvalidOperationException("Consecutive newlines lost their individual page positions before following text.");
    if (name.StartsWith("terminal-") && !name.StartsWith("terminal-followed-") && (package.Pages.Count != 1 || compactText != "A"))
        throw new InvalidOperationException("Consecutive terminal newlines introduced a blank tail page.");
    if (name.StartsWith("styled-newline-"))
    {
        var text = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
            .Where(value => new[] { "A", "B", "C", "D" }.Contains(value.Text)).ToArray();
        var mixedSizes = name == "styled-newline-public" ? 22d : 44 * 25.4 / 72;
        if (package.Pages.Count != 1 || text.Length != 4 || string.Concat(text.Select(value => value.Text)) != "ABCD" ||
            Math.Abs(text[1].YMillimeters - text[0].YMillimeters - text[0].FontSizeMillimeters * 1.3 * 3) > 0.005)
            throw new InvalidOperationException("Empty lines lost their control run font size.");
        if (Math.Abs(text[3].YMillimeters - text[2].YMillimeters - (text[2].FontSizeMillimeters + mixedSizes) * 1.3) > 0.005)
            throw new InvalidOperationException("A following newline overwrote the preceding newline size.");
    }
    if (name.StartsWith("alignment-"))
    {
        var alpha = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
            .Where(element => element.Text.StartsWith("Alpha")).ToArray();
        if (alpha.Length != 4 || alpha[1].Text != "Alpha  " || alpha[3].Text != "Alpha  " ||
            Math.Abs(alpha[0].XMillimeters - alpha[1].XMillimeters) > 0.00001 ||
            Math.Abs(alpha[2].XMillimeters - alpha[3].XMillimeters) > 0.00001)
            throw new InvalidOperationException("Wrapped separators changed visible alignment.");
        foreach (var word in new[] { "café", "cafe\u0301", "mínimo", "mi\u0301nimo" })
            if (package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
                .Count(element => element.Text == word) != 1)
                throw new InvalidOperationException("An accented Latin word was split within its fresh-line width.");
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
