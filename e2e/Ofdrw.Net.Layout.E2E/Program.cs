using System.Security.Cryptography;
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

async Task ConvertDocx(string name, DocxConversionOptions? options)
{
    await using var input = File.OpenRead(docxPath);
    await using var output = File.Create(Path.Combine(outputDirectory, name + ".ofd"));
    var converter = options is null ? new DocxToOfdConverter() : new DocxToOfdConverter(options);
    await converter.ConvertAsync(input, output);
}

await ConvertDocx("docx-native", new DocxConversionOptions { OfdMode = DocxToOfdMode.Native });
await ConvertDocx("docx-default", null);

string? nativeText = null;
foreach (var name in new[] { "flow-public", "docx-native", "docx-default" })
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
    Console.WriteLine($"{name}: pages={package.Pages.Count}, chars={compactText.Length}, ofdBytes={new FileInfo(ofdPath).Length}, pdfBytes={new FileInfo(pdfPath).Length}, sha256={Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ofdPath)))}");
}

static string Compact(string text) => new(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
