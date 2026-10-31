using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Layout.Editing;

using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Signatures.Verification;
using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using System.Text;
using System.Xml.Linq;
using PdfSharpCore.Drawing;

var repoRoot = ResolveRepoRoot();
// Exercise host resolver composition before PDFsharp initializes its cache.
PdfSharpCore.Fonts.GlobalFontSettings.FontResolver = Ofdrw.Net.Converter.Pdf.PdfFontRegistry.CreateResolver(
    new HostFailureResolver(new PdfSharpCore.Utils.FontResolver(), repoRoot));
var outputDir = Environment.GetEnvironmentVariable("OFDRW_E2E_OUTPUT_DIR") ?? Path.Combine(repoRoot, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "output");
var testDataDir = Path.Combine(repoRoot, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "testdata", "upstream-ofdrw");
Directory.CreateDirectory(outputDir);
var visualOnly = args.Contains("--visual-corpus-only");
var captureBaseline = args.Contains("--capture-visual-baseline");
if (args.Any(arg => arg is not ("--visual-corpus-only" or "--capture-visual-baseline")) ||
    (captureBaseline && !visualOnly))
    throw new ArgumentException("Supported options: --visual-corpus-only [--capture-visual-baseline].");
if (visualOnly)
{
    await ValidateVisualCorpusAsync(repoRoot, outputDir, captureBaseline);
    Console.WriteLine(captureBaseline ? "[E2E] CAPTURE_ONLY_NOT_VALIDATED: review before adding golden images."
        : "[E2E] Visual corpus gate passed.");
    return;
}

// Consume both plain-text entrypoints from the installed Layout candidate.
const string plainSource = "  中英 iii WWW\tX\u00A0Y e\u0301  \r\n\r\n尾行\n";
var plainConverter = new PlainTextOfdConverter();
var plainString = plainConverter.Render(plainSource);
using var plainInput = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("\uFEFF" + plainSource));
var plainUtf8 = await plainConverter.RenderUtf8Async(plainInput);
foreach (var (label, plainPackage) in new[] { ("string", plainString), ("utf8", plainUtf8) })
{
    var path = Path.Combine(outputDir, "plain-text-" + label + ".ofd");
    await using (var output = File.Create(path)) await new OfdPackageWriter().WriteAsync(plainPackage, output);
    await using var input = File.OpenRead(path);
    var parsed = await new OfdReader().ReadAsync(input);
    var actual = string.Concat(parsed.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(t => t.Text));
    if (actual != plainSource.Replace("\r", "").Replace("\n", "") || parsed.Pages.SelectMany(p => p.Elements).Any(e => e is not OfdTextElement))
        throw new InvalidOperationException("Plain text package lost whitespace/Unicode or used non-native content.");
}
Console.WriteLine("[E2E] Installed Layout plain-text string/strict UTF-8 APIs passed exact native-text roundtrip.");

// Consume the public Layout package through the built nupkg, with no project reference.
var publicFlow = new FlowDocument();
publicFlow.Options.PageHeightMillimeters = 60;
publicFlow.Options.MarginTopMillimeters = 8;
publicFlow.Options.MarginBottomMillimeters = 8;
for (var i = 0; i < 30; i++)
    publicFlow.Blocks.Add(new Paragraph($"第{i + 1}段 Public flow words information preserve text."));
var flowPackage = publicFlow.Render();
if (flowPackage.Pages.Count < 2) throw new InvalidOperationException("Public flow did not paginate.");
var flowPath = Path.Combine(outputDir, "public-flow.ofd");
await using (var stream = File.Create(flowPath)) await new OfdPackageWriter().WriteAsync(flowPackage, stream);
await using (var stream = File.OpenRead(flowPath))
{
    var parsed = await new OfdReader().ReadAsync(stream);
    if (parsed.Pages.Count != flowPackage.Pages.Count ||
        !string.Concat(parsed.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.Text)).Contains("information"))
        throw new InvalidOperationException("Public flow package lost text during round trip.");
}
Console.WriteLine($"[E2E] Public Flow package generated {flowPackage.Pages.Count} pages from Paragraph/Span.");

var publicTables = new FlowDocument();
publicTables.Options.PageHeightMillimeters = 60;
publicTables.Options.MarginTopMillimeters = publicTables.Options.MarginBottomMillimeters = 8;
var consumerTable = new Table();
for (var i = 0; i < 5; i++)
{
    var row = new Row { MinimumHeightMillimeters = 18 };
    row.Cells.Add(new Cell($"合并{i}") { ColumnSpan = 2, BackgroundColor = new OfdColor(220, 230, 250) });
    var right = new Cell($"Right{i}") { VerticalAlignment = CellVerticalAlignment.Bottom };
    right.Paragraphs[0].Alignment = ParagraphAlignment.Right;
    row.Cells.Add(right); consumerTable.Rows.Add(row);
}
publicTables.Blocks.Add(consumerTable);
var tablePackage = publicTables.Render();
if (tablePackage.Pages.Count != 3) throw new InvalidOperationException("Public table did not paginate whole rows.");
var tablePath = Path.Combine(outputDir, "public-tables.ofd");
await using (var stream = File.Create(tablePath)) await new OfdPackageWriter().WriteAsync(tablePackage, stream);
await using (var stream = File.OpenRead(tablePath))
{
    var parsed = await new OfdReader().ReadAsync(stream);
    var expected = string.Concat(Enumerable.Range(0, 5).Select(i => $"合并{i}Right{i}"));
    if (string.Concat(parsed.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(t => t.Text)) != expected)
        throw new InvalidOperationException("Public table package lost or repeated text.");
}
await using (var input = File.OpenRead(tablePath))
await using (var output = File.Create(Path.Combine(outputDir, "public-tables.pdf")))
    await new OfdToPdfConverter().ConvertAsync(input, output);
Console.WriteLine("[E2E] Public Table/Row/Cell consumed from nupkg: horizontal merge, alignment, shading, 3 whole-row pages.");
// The package-consumer copy of this program must resolve the new Layout surface
// from the local NuGet feed, including its internal Core numeric dependency.
var graphicsPackage = new OfdDocumentPackage();
var graphicsPage = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 };
graphicsPackage.Pages.Add(graphicsPage);
var graphics = new OfdGraphics(graphicsPackage, graphicsPage);
graphics.Scale(0.9996, 1);
graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(0, 0, 100, 80));
graphics.DrawRectangle(new OfdPen(OfdColor.Black, 0.5), 5, 5, 90, 60);
graphics.DrawString("Graphics 包消费", new OfdFont("Graphics fixture", 3.3334, italic: true), new OfdBrush(OfdColor.Black), 10, 20);
using (var encoded = new MemoryStream())
{
    await new OfdPackageWriter().WriteAsync(graphicsPackage, encoded);
    File.WriteAllBytes(Path.Combine(outputDir, "native-graphics-package.ofd"), encoded.ToArray());
    encoded.Position = 0;
    var decoded = await new OfdReader().ReadAsync(encoded);
    if (decoded.Pages[0].Elements.Count != 2 || decoded.Pages[0].Elements.OfType<OfdImageElement>().Any()) throw new InvalidOperationException("Package graphics lost native objects.");
    if (decoded.Pages[0].Elements.OfType<OfdPathElement>().Single().Transform![0] != 0.9996) throw new InvalidOperationException("Package graphics lost CTM precision.");
    var text = decoded.Pages[0].Elements.OfType<OfdTextElement>().Single();
    if (text.Text != "Graphics 包消费" || !text.SourceXml!.Contains("FauxItalicMatrixV1", StringComparison.Ordinal)) throw new InvalidOperationException("Package graphics lost text/native italic.");
}
Console.WriteLine("[E2E] Installed Layout native graphics, exact CTM and name-only italic package consumption passed.");

// Keep R09's public entrypoints in the isolated NuGet consumer as well as the
// project-reference visual E2E. The two-color gradient is explicitly rasterized.
var r09Pixels = Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAIUlEQVR4nGP4z8DwHx8mAo4aMAwMIKAfCBjw4lEDhoUBAFfoPd+1WG54AAAAAElFTkSuQmCC");
var r09Package = new OfdDocumentPackage();
var r09Page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 70 };
r09Package.Pages.Add(r09Page);
var r09 = new OfdGraphics(r09Package, r09Page);
r09.DrawImage(r09Pixels, 5, 5, 20, 20, 8, 0, 8, 8, alpha: 180);
r09.FillLinearGradientRectangle(new OfdColor(255, 0, 0), new OfdColor(0, 0, 255), 30, 5, 30, 20);
r09.FillImagePatternRectangle(r09Pixels, 5, 30, 25, 15, 10, 10);
r09.DrawLine(new OfdPen(OfdColor.Black, 1, OfdLineCap.Round, OfdLineJoin.Bevel,
    dashPattern: new[] { 4d, 2d }, dashOffset: 1), 35, 40, 85, 40);
var r09Path = Path.Combine(outputDir, "r09-package-graphics.ofd");
await using (var stream = File.Create(r09Path)) await new OfdPackageWriter().WriteAsync(r09Package, stream);
await using (var stream = File.OpenRead(r09Path))
{
    var parsed = await new OfdReader().ReadAsync(stream);
    if (parsed.Pages[0].Elements.OfType<OfdImageElement>().Count() != 8 ||
        parsed.Pages[0].Elements.OfType<OfdPathElement>().Count() != 1)
        throw new InvalidOperationException("Installed R09 graphics lost raster fills, tile placements or pen.");
}
await using (var input = File.OpenRead(r09Path))
await using (var pdf = File.Create(Path.Combine(outputDir, "r09-package-graphics.pdf")))
    await new OfdToPdfConverter().ConvertAsync(input, pdf);
await using (var input = File.OpenRead(r09Path))
await using (var svg = File.Create(Path.Combine(outputDir, "r09-package-graphics.svg")))
    await new OfdToSvgConverter().ConvertAsync(input, svg);
await using (var input = File.OpenRead(r09Path))
await using (var html = File.Create(Path.Combine(outputDir, "r09-package-graphics.html")))
    await new OfdToHtmlConverter().ConvertAsync(input, html);
Console.WriteLine("[E2E] Installed R09 DrawImage, raster fills, pen and three exports passed.");

// Used and unused name-only resources must survive unavailable/unsupported host
// fonts. Each used text must obtain a registered physical face with verified
// coverage, including default/Arial candidates after optional host failures.
foreach (var failure in new[] { "Resolve", "Read", "Ttc", "Broken" })
{
    foreach (var used in new[] { false, true })
    {
        var family = "Ofdrw-HostFailure-" + failure;
        var probe = new OfdDocumentPackage();
        probe.Fonts.Add(new OfdFontResource { Id = "probe", FontName = family, Bold = true });
        var page = new OfdPage { WidthMillimeters = 80, HeightMillimeters = 50 };
        page.Elements.Add(new OfdTextElement { Text = "ABCD", FontName = used ? family : "Arial",
            FontResourceId = used ? "probe" : null, FontSizeMillimeters = 5, XMillimeters = 10, YMillimeters = 10 });
        probe.Pages.Add(page);
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(probe, ofd); ofd.Position = 0;
        using var pdf = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var readback = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        if (readback.GetPage(1).Text != "ABCD") throw new InvalidOperationException("Host font fallback lost text.");
    }
}
Console.WriteLine("[E2E] Host resolve/read failures, TTC and malformed fonts preserved used/unused resource fallback.");

var sourceOfdPath = Path.Combine(outputDir, "source.ofd");
var convertedPdfPath = Path.Combine(outputDir, "converted.pdf");
var roundtripOfdPath = Path.Combine(outputDir, "roundtrip.ofd");

Console.WriteLine($"[E2E] Output directory: {outputDir}");

var builder = new OfdDocumentBuilder();
builder.SetOptions(new OfdDocumentOptions
{
    DocType = "OFD-H",
    DocumentId = "Doc_0",
    Metadata = new OfdMetadata
    {
        Title = "Ofdrw.Net Package E2E",
        Creator = "Ofdrw.Net.Converter.Pdf.E2E",
        CreationDate = DateTimeOffset.UtcNow,
        ModificationDate = DateTimeOffset.UtcNow
    }
});

builder.AddPage(new OfdPage
{
    Index = 0,
    WidthMillimeters = 210,
    HeightMillimeters = 297,
    Elements =
    {
        new OfdTextElement
        {
            Text = "Package install E2E: OFD -> PDF -> OFD",
            FontName = "SimSun",
            FontSizeMillimeters = 4,
            XMillimeters = 10,
            YMillimeters = 12,
            WidthMillimeters = 160,
            HeightMillimeters = 12
        }
    }
});

var package = builder.Build();
var writer = new OfdPackageWriter();
await using (var sourceOfdStream = File.Create(sourceOfdPath))
{
    await writer.WriteAsync(package, sourceOfdStream);
}

var ofdToPdf = new OfdToPdfConverter();
await using (var ofdInput = File.OpenRead(sourceOfdPath))
await using (var pdfOutput = File.Create(convertedPdfPath))
{
    await ofdToPdf.ConvertAsync(ofdInput, pdfOutput);
}

var pdfToOfd = new PdfToOfdConverter();
var d3Rows = new List<object>();
foreach (var (variant, expectedHash) in new[] {
    ("cell", "7040cf0315281ffe97853b7ac9b8373135312400ba9ea7fd7b69a8966e5c5ea8"),
    ("fixed", "7b69e87771891ba440bde051af33284f3e008b376744751bc27064391c366ea5") })
{
    var fixture = Path.Combine(repoRoot, "tests/Ofdrw.Net.Converter.Pdf.Tests/Fixtures/D3", $"fpdf2-two-word-{variant}.pdf");
    var bytes = File.ReadAllBytes(fixture);
    string Hash(byte[] data) => Convert.ToHexString(SHA256.HashData(data)).ToLowerInvariant();
    if (Hash(bytes) != expectedHash) throw new InvalidDataException("D3 input identity changed.");
    using var rejected = new MemoryStream();
    string rejection;
    try
    {
        await pdfToOfd.ConvertAsync(new MemoryStream(bytes), rejected);
        throw new InvalidOperationException("D3 semantic import accepted unverified reading order.");
    }
    catch (NotSupportedException error) when (error.Message.Contains("PDF_TEXT_READING_ORDER_UNVERIFIED"))
    {
        rejection = error.Message;
        if (rejected.Length != 0) throw new InvalidOperationException("D3 rejection wrote output bytes.");
    }
    var imageOfd = Path.Combine(outputDir, $"d3-{variant}-image-only.ofd");
    using (var target = File.Create(imageOfd))
        await new PdfToOfdConverter(new Ofdrw.Net.Converter.Pdf.PdfToOfdOptions {
            TextLayerMode = Ofdrw.Net.Converter.Pdf.PdfTextLayerMode.None, RequireRasterization = true
        }).ConvertAsync(new MemoryStream(bytes), target);
    using (var input = File.OpenRead(imageOfd))
    {
        var imported = await new OfdReader().ReadAsync(input);
        if (imported.Pages.Count != 1 || imported.Pages[0].Elements.OfType<OfdImageElement>().Count() != 1 ||
            imported.Pages[0].Elements.OfType<OfdTextElement>().Any())
            throw new InvalidOperationException("D3 raster-only output claims text or loses its image.");
    }
    var imagePdf = Path.Combine(outputDir, $"d3-{variant}-image-only.pdf");
    using (var input = File.OpenRead(imageOfd))
    using (var target = File.Create(imagePdf)) await ofdToPdf.ConvertAsync(input, target);
    d3Rows.Add(new { variant, fixture, inputSha256 = expectedHash, rejection,
        rejectedOutputBytes = rejected.Length, imageOfd, ofdSha256 = Hash(File.ReadAllBytes(imageOfd)),
        imagePdf, pdfSha256 = Hash(File.ReadAllBytes(imagePdf)) });
}
var d3Assembly = typeof(PdfToOfdConverter).Assembly;
File.WriteAllText(Path.Combine(outputDir, "d3-import-receipt.json"), JsonSerializer.Serialize(new {
    assemblyPath = d3Assembly.Location,
    assemblySha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(d3Assembly.Location))).ToLowerInvariant(),
    rows = d3Rows
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("[E2E] D3 unverified RTL rejection and explicit raster-only import passed from installed packages.");
await using (var pdfInput = File.OpenRead(convertedPdfPath))
await using (var ofdOutput = File.Create(roundtripOfdPath))
{
    await pdfToOfd.ConvertAsync(pdfInput, ofdOutput);
}

var reader = new OfdReader();
await using var roundtripStream = File.OpenRead(roundtripOfdPath);
var roundtrip = await reader.ReadAsync(roundtripStream);

if (roundtrip.Pages.Count == 0)
{
    throw new InvalidOperationException("Roundtrip OFD has no pages.");
}

Console.WriteLine($"[E2E] Source OFD:    {sourceOfdPath}");
Console.WriteLine($"[E2E] Converted PDF: {convertedPdfPath}");
Console.WriteLine($"[E2E] Roundtrip OFD: {roundtripOfdPath}");
Console.WriteLine($"[E2E] Roundtrip page count: {roundtrip.Pages.Count}");
// Consume the new public image APIs from the packed assembly, not a project reference.
var imageExportPath = Path.Combine(outputDir, "image-api.png");
await using (var input = File.OpenRead(sourceOfdPath))
await using (var output = File.Create(imageExportPath))
    await new OfdToImageConverter(new Ofdrw.Net.Converter.Pdf.OfdToImageOptions { PixelsPerMillimeter = 2 }).ConvertAsync(input, output);
await AssertImageHasContentAsync(imageExportPath);
var imageImportPath = Path.Combine(outputDir, "image-api.ofd");
await using (var input = File.OpenRead(imageExportPath))
await using (var output = File.Create(imageImportPath))
    await new ImageToOfdConverter(new Ofdrw.Net.Converter.Pdf.ImageToOfdOptions
    { PixelsPerMillimeter = 2, PageSize = new OfdPageSize { WidthMillimeters = 220, HeightMillimeters = 307 } }).ConvertAsync(input, output);
await using (var input = File.OpenRead(imageImportPath))
{
    var imported = await reader.ReadAsync(input);
    var imagePage = imported.Pages.Single(); var image = imagePage.Elements.OfType<OfdImageElement>().Single();
    if (image.XMillimeters != 5 || image.YMillimeters != 5 || image.WidthMillimeters != 210 || image.HeightMillimeters != 297)
        throw new InvalidOperationException("Packed image API lost natural-size centering geometry.");
}
Console.WriteLine("[E2E] Packed PNG export and image import retained 210×297 mm content centered on 220×307 mm page.");
// The installed Layout/Reader/Packaging assemblies must compose with the new image APIs.
var imageToolsPath = Path.Combine(outputDir, "image-tools.ofd");
await using (var input = File.OpenRead(imageImportPath))
{
    var imported = await reader.ReadAsync(input);
    OfdWatermark.AddImage(imported, [0], File.ReadAllBytes(Path.Combine(repoRoot, "e2e/Ofdrw.Net.DocumentTools.E2E/mark.png")),
        "image/png", new OfdWatermarkOptions { XMillimeters = 1, YMillimeters = 1, WidthMillimeters = 8, HeightMillimeters = 4 }, alpha: 255);
    var split = OfdDocumentSplitter.Split(imported, [0]);
    if (split.Pages.Single().Elements.OfType<OfdImageElement>().Count() != 2)
        throw new InvalidOperationException("Packed watermark/split lost image resource ownership.");
    await using var target = File.Create(imageToolsPath); await writer.WriteAsync(split, target);
}
var imageToolsPng = Path.Combine(outputDir, "image-tools.png");
await using (var input = File.OpenRead(imageToolsPath))
await using (var output = File.Create(imageToolsPng))
    await new OfdToImageConverter(new Ofdrw.Net.Converter.Pdf.OfdToImageOptions { PixelsPerMillimeter = 2 }).ConvertAsync(input, output);
await AssertImageHasContentAsync(imageToolsPng);
Console.WriteLine("[E2E] Packed image import -> image watermark -> split -> PNG retained both image objects.");
Console.WriteLine("[E2E] Running upstream sample validation...");

foreach (var sampleOfd in Directory.EnumerateFiles(testDataDir, "*.ofd").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
{
    await using var stream = File.OpenRead(sampleOfd);
    var parsed = await reader.ReadAsync(stream);
    if (parsed.Pages.Count == 0)
    {
        throw new InvalidOperationException($"Sample OFD has no pages: {sampleOfd}");
    }

    if (string.Equals(Path.GetFileName(sampleOfd), "999.ofd", StringComparison.OrdinalIgnoreCase))
    {
        AssertComplexUpstreamSample(parsed, sampleOfd);
        await ValidateComplexSampleArtifactsAsync(sampleOfd, outputDir);
    }

    Console.WriteLine($"[E2E] Upstream OFD parsed: {Path.GetFileName(sampleOfd)} ({parsed.Pages.Count} page(s))");
}

foreach (var samplePdf in Directory.EnumerateFiles(testDataDir, "*.pdf").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
{
    await ValidatePdfSampleAsync(samplePdf, outputDir);
}

var docxSamplePath = Path.Combine(
    repoRoot,
    "e2e",
    "Ofdrw.Net.Converter.Docx.E2E",
    "testdata",
    "generated-layout.docx");
await ValidateDocxSampleAsync(docxSamplePath, outputDir);
await ValidateVisualCorpusAsync(repoRoot, outputDir);

Console.WriteLine("[E2E] Success: package installation and conversion flow is working.");

static async Task ValidateVisualCorpusAsync(string repoRoot, string outputDir, bool captureBaseline = false)
{
    var platform = OperatingSystem.IsLinux() ? "linux" : OperatingSystem.IsMacOS() ? "macos" : OperatingSystem.IsWindows() ? "windows" :
        throw new PlatformNotSupportedException("Visual corpus requires a reviewed platform baseline.");
    var goldenDir = Path.Combine(repoRoot, "e2e", "visual-corpus", "golden", platform);
    var actualDir = Path.Combine(outputDir, "visual-corpus");
    Directory.CreateDirectory(actualDir);
    // Bind every platform to the same reviewed font bytes. A host fallback
    // changed macOS CI antialiasing while leaving the geometric page intact.
    var visualFont = ReadVisualCorpusFont(repoRoot);
    var visualFamily = Ofdrw.Net.Converter.Pdf.PdfFontRegistry.RegisterFontFace(visualFont);
    var captures = new List<object>();
    var missingGoldens = new List<string>();
    var comparisonFailures = new List<string>();
    var metrics = new List<string> { "sample,page,region,normalized_rmse,threshold" };
    foreach (var kind in new[] { "ticket", "invoice", "template", "edge-layout" })
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { Index = 0, WidthMillimeters = 120, HeightMillimeters = 80 };
        package.Pages.Add(page);
        static OfdPathElement Box(double x, double y, double w, double h, OfdColor color) => new()
        {
            XMillimeters = x, YMillimeters = y, WidthMillimeters = w, HeightMillimeters = h,
            AbbreviatedData = FormattableString.Invariant($"M 0 0 L {w} 0 L {w} {h} L 0 {h} C"),
            Stroke = false, Fill = true, FillColor = color
        };
        static OfdPathElement Rule(double x, double y, double w, double h) => new()
        {
            XMillimeters = x, YMillimeters = y, WidthMillimeters = Math.Max(0.4, w),
            HeightMillimeters = Math.Max(0.4, h),
            AbbreviatedData = FormattableString.Invariant($"M 0 0 L {w} {h}"),
            Stroke = true, Fill = false, LineWidthMillimeters = 0.35,
            StrokeColor = new OfdColor(45, 68, 88)
        };
        page.Elements.Add(Box(5, 5, 110, 70, new OfdColor(238, 244, 250)));
        page.Elements.Add(Box(8, 8, 104, 12, new OfdColor(28, 75, 126)));
        page.Elements.Add(new OfdTextElement { Text = kind.ToUpperInvariant(), FontName = "Noto Sans CJK SC",
            FontSizeMillimeters = 5, XMillimeters = 12, YMillimeters = 10,
            WidthMillimeters = 85, HeightMillimeters = 8, FillColor = new OfdColor(255, 255, 255) });
        var caption = kind switch
        {
            "ticket" => "票据编号 0001 / Ticket No. 0001",
            "invoice" => "发票明细 / Invoice items",
            "template" => "模板字段：姓名、日期 / Template fields",
            _ => "异常布局：边界与叠放 / Edge layout"
        };
        page.Elements.Add(new OfdTextElement { Text = caption, FontName = "Noto Sans CJK SC",
            FontSizeMillimeters = 3.2, XMillimeters = 12, YMillimeters = 22,
            WidthMillimeters = 100, HeightMillimeters = 6,
            FillColor = new OfdColor(35, 55, 75) });
        switch (kind)
        {
            case "ticket":
                page.Elements.Add(Box(12, 29, 72, 3, new OfdColor(35, 35, 35)));
                page.Elements.Add(Box(12, 38, 48, 3, new OfdColor(35, 35, 35)));
                page.Elements.Add(Box(92, 28, 12, 25, new OfdColor(12, 12, 12)));
                break;
            case "invoice":
                for (var row = 0; row < 4; row++)
                    page.Elements.Add(Rule(10, 31 + row * 9, 100, 0));
                page.Elements.Add(Rule(10, 31, 0, 27));
                page.Elements.Add(Rule(70, 31, 0, 27));
                page.Elements.Add(Rule(110, 31, 0, 27));
                page.Elements.Add(new OfdTextElement { Text = "Item A  12.50", FontName = "Noto Sans CJK SC",
                    FontSizeMillimeters = 3.2, XMillimeters = 14, YMillimeters = 33,
                    WidthMillimeters = 74, HeightMillimeters = 5 });
                page.Elements.Add(Box(86, 62, 24, 5, new OfdColor(196, 42, 42)));
                break;
            case "template":
                page.Elements.Add(Box(12, 28, 96, 40, new OfdColor(214, 231, 221)));
                page.Elements.Add(Box(18, 34, 70, 3, new OfdColor(52, 96, 68)));
                page.Elements.Add(Box(18, 46, 50, 3, new OfdColor(52, 96, 68)));
                break;
            default:
                // A valid, intentionally awkward layout: touching and clipped shapes.
                page.Elements.Add(Box(16, 30, 76, 18, new OfdColor(245, 178, 42)));
                page.Elements.Add(Box(78, 40, 33, 25, new OfdColor(96, 45, 145)));
                page.Elements.Add(Box(105, 55, 25, 10, new OfdColor(210, 58, 58)));
                break;
        }

        if (visualFont is not null)
        {
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans CJK SC",
                FamilyName = "Noto Sans CJK SC", Data = visualFont });
            using var measure = XGraphics.CreateMeasureContext(new XSize(1000, 1000), XGraphicsUnit.Point, XPageDirection.Downwards);
            foreach (var text in page.Elements.OfType<OfdTextElement>())
            {
                text.FontResourceId = "10";
                var font = new XFont(visualFamily, text.FontSizeMillimeters * 72 / 25.4, XFontStyle.Regular);
                // Real glyph advances make the fixture's placement explicit in
                // direct OFD readers, including readers that omit implicit advances.
                text.Runs.Add(new OfdTextRun { Text = text.Text, YMillimeters = text.FontSizeMillimeters,
                    DeltaX = string.Join(" ", text.Text.EnumerateRunes().SkipLast(1).Select(rune =>
                        (measure.MeasureString(rune.ToString(), font).Width * 25.4 / 72)
                            .ToString("0.######", CultureInfo.InvariantCulture))) });
            }
        }
        var ofdPath = Path.Combine(actualDir, kind + ".ofd");
        var pdfPath = Path.Combine(actualDir, kind + ".pdf");
        await using (var ofd = File.Create(ofdPath))
            await new OfdPackageWriter().WriteAsync(package, ofd);
        await using (var ofd = File.OpenRead(ofdPath))
        await using (var pdf = File.Create(pdfPath))
            await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        var imagePath = Path.Combine(actualDir, kind + ".png");
        var render = await RunProcessAsync("pdftoppm", $"-f 1 -l 1 -r 144 -png -singlefile \"{pdfPath}\" \"{Path.Combine(actualDir, kind)}\"");
        if (render.ExitCode != 0 || !File.Exists(imagePath))
            throw new InvalidOperationException($"Visual corpus render failed for {kind}: {render.Error}");
        captures.Add(new { sample = kind, page = 1, image = Path.GetFileName(imagePath),
            sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(imagePath))).ToLowerInvariant() });
        if (captureBaseline) continue;
        var golden = Path.Combine(goldenDir, kind + ".png");
        if (!File.Exists(golden))
            missingGoldens.Add(golden);
        else
        {
            try
            {
                var error = await AssertImagesCloseAsync(golden, imagePath, 0.035,
                    Path.Combine(actualDir, kind + "-diff.png"));
                metrics.Add(FormattableString.Invariant($"{kind},1,full,{error:R},0.035"));
                foreach (var (region, geometry) in new[] { ("title", "400x90+40+30"), ("body", "560x260+55+145") })
                {
                    var expectedCrop = Path.Combine(actualDir, kind + "-" + region + "-expected.png");
                    var actualCrop = Path.Combine(actualDir, kind + "-" + region + "-actual.png");
                    foreach (var (source, target) in new[] { (golden, expectedCrop), (imagePath, actualCrop) })
                    {
                        var crop = await RunProcessAsync("convert", $"\"{source}\" -crop {geometry} +repage \"{target}\"");
                        if (crop.ExitCode != 0)
                            throw new InvalidOperationException($"Visual corpus crop failed: {crop.Error}");
                    }
                    var regionError = await AssertImagesCloseAsync(expectedCrop, actualCrop, 0.035,
                        Path.Combine(actualDir, kind + "-" + region + "-diff.png"));
                    metrics.Add(FormattableString.Invariant($"{kind},1,{region},{regionError:R},0.035"));
                }
                Console.WriteLine(FormattableString.Invariant($"[E2E] Visual corpus {kind}: page 1 RMSE {error:R} <= 0.035."));
            }
            catch (InvalidOperationException failure)
            {
                comparisonFailures.Add(failure.Message);
            }
        }
    }
    File.WriteAllText(Path.Combine(actualDir, "visual-run.json"), JsonSerializer.Serialize(new {
        status = captureBaseline ? "CAPTURE_ONLY_NOT_VALIDATED" :
            missingGoldens.Count == 0 && comparisonFailures.Count == 0 ? "PASS" : "FAIL",
        platform, dpi = 144, threshold = 0.035,
        sourceSha = Environment.GetEnvironmentVariable("OFDRW_SOURCE_SHA"),
        fontSha256 = Convert.ToHexString(SHA256.HashData(visualFont)).ToLowerInvariant(),
        captures, missingGoldens, comparisonFailures
    }, new JsonSerializerOptions { WriteIndented = true }));
    File.WriteAllLines(Path.Combine(actualDir, "visual-metrics.csv"), metrics);
    if (missingGoldens.Count != 0 || comparisonFailures.Count != 0)
        throw new InvalidOperationException("Visual corpus gate failed: " +
            string.Join("; ", missingGoldens.Select(path => "Missing golden " + path).Concat(comparisonFailures)));
}

static byte[] ReadVisualCorpusFont(string repoRoot)
{
    const string digest = "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa";
    var archive = Path.Combine(repoRoot, "tests", "fixtures", "trusted-fonts",
        "sha256-40c6444465a1149b2ef8b7a47c5bae3f9d3d47b701e9a5b60bcfadcae4df6f59.zip");
    using var zip = ZipFile.OpenRead(archive);
    using var input = (zip.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf") ??
        throw new InvalidDataException("Trusted visual font is absent; materialize test fixtures with Git LFS.")).Open();
    using var output = new MemoryStream();
    input.CopyTo(output);
    var bytes = output.ToArray();
    if (bytes.Length != 21659544 || Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() != digest)
        throw new InvalidDataException("Trusted visual font hash mismatch.");
    return bytes;
}

static DocxConversionOptions CreateDocxFixtureOptions(string outputDir)
{
    var options = new DocxConversionOptions();
    if (OperatingSystem.IsWindows())
    {
        // The DOCX requests Noto explicitly. Isolated Windows validation must
        // supply its real bytes instead of relying on installed host fonts.
        var directory = Path.Combine(outputDir, "fixture-fonts");
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, "NotoSansCJKsc-Regular.ttf"), ReadVisualCorpusFont(ResolveRepoRoot()));
        options.FontDirectories.Add(directory);
        options.UseInstalledMicrosoftOfficeFonts = false;
    }
    return options;
}

static async Task ValidateDocxSampleAsync(string samplePath, string outputDir)
{
    var docxOptions = CreateDocxFixtureOptions(outputDir);
    docxOptions.Engine = DocxConversionEngine.BuiltIn;
    docxOptions.OfdMode = DocxToOfdMode.DualLayer;
    var pdfPath = Path.Combine(outputDir, "generated-docx.pdf");
    var nativeOfdPath = Path.Combine(outputDir, "generated-docx-native.ofd");
    var defaultOfdPath = Path.Combine(outputDir, "generated-docx-default.ofd");
    var directOfdPath = Path.Combine(outputDir, "generated-docx.ofd");
    var pdfOfdPath = Path.Combine(outputDir, "generated-docx-pdf-stage.ofd");

    await using (var docxInput = File.OpenRead(samplePath))
    await using (var pdfOutput = File.Create(pdfPath))
    {
        await new DocxToPdfConverter(docxOptions).ConvertAsync(docxInput, pdfOutput);
    }

    await using (var docxInput = File.OpenRead(samplePath))
    await using (var ofdOutput = File.Create(directOfdPath))
    {
        await new DocxToOfdConverter(docxOptions).ConvertAsync(docxInput, ofdOutput);
    }

    await using (var docxInput = File.OpenRead(samplePath))
    await using (var ofdOutput = File.Create(nativeOfdPath))
    {
        var nativeOptions = CreateDocxFixtureOptions(outputDir);
        nativeOptions.OfdMode = DocxToOfdMode.Native;
        await new DocxToOfdConverter(nativeOptions).ConvertAsync(docxInput, ofdOutput);
    }

    await using (var pdfInput = File.OpenRead(pdfPath))
    await using (var ofdOutput = File.Create(pdfOfdPath))
    {
        await new PdfToOfdConverter().ConvertAsync(pdfInput, ofdOutput);
    }

    await using (var docxInput = File.OpenRead(samplePath))
    await using (var ofdOutput = File.Create(defaultOfdPath))
        await new DocxToOfdConverter(CreateDocxFixtureOptions(outputDir)).ConvertAsync(docxInput, ofdOutput);

    foreach (var ofdPath in new[] { directOfdPath, pdfOfdPath })
    {
        await using var ofdInput = File.OpenRead(ofdPath);
        var package = await new OfdReader().ReadAsync(ofdInput);
        if (package.Pages.Count != 2 ||
            package.Pages.Any(page =>
                page.Elements.OfType<OfdImageElement>().All(image => image.Data.Length == 0) ||
                !page.Elements.OfType<OfdTextElement>().Any()))
        {
            throw new InvalidOperationException($"DOCX conversion output is incomplete: {ofdPath}");
        }
    }

    using var sourceDocument = WordprocessingDocument.Open(samplePath, false);
    var original = string.Concat(sourceDocument.MainDocumentPart!.Document.Body!.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>().Select(text => text.Text));
    static string Compact(string value) => string.Concat(value.Where(character => !char.IsWhiteSpace(character)));
    foreach (var path in new[] { nativeOfdPath, defaultOfdPath, Path.Combine(outputDir, "cli-native.ofd") })
    {
        if (!File.Exists(path)) throw new InvalidOperationException($"Required native output is missing: {path}");
        await using (var input = File.OpenRead(path))
        {
            var native = await new OfdReader().ReadAsync(input);
            if (Compact(new OfdTextExtractor().Extract(native)) != Compact(original))
                throw new InvalidOperationException($"Native output did not preserve original DOCX text: {path}");
        }
        var rendered = Path.ChangeExtension(path, ".pdf");
        await using (var input = File.OpenRead(path))
        await using (var output = File.Create(rendered))
            await new OfdToPdfConverter().ConvertAsync(input, output);
        var render = await RunProcessAsync("pdftoppm", $"-png -r 144 \"{rendered}\" \"{Path.Combine(outputDir, Path.GetFileNameWithoutExtension(path))}\"");
        if (render.ExitCode != 0) throw new InvalidOperationException($"Native OFD PDF rendering failed: {render.Error}");
        var images = Directory.EnumerateFiles(outputDir, Path.GetFileNameWithoutExtension(path) + "-*.png").OrderBy(value => value, StringComparer.Ordinal).ToArray();
        if (images.Length != 2) throw new InvalidOperationException($"Expected two freshly rendered native pages: {path}");
        foreach (var image in images) await AssertImageHasContentAsync(image);
        if (path != nativeOfdPath)
        {
            var expectedImages = Directory.EnumerateFiles(outputDir, "generated-docx-native-*.png").OrderBy(value => value, StringComparer.Ordinal).ToArray();
            for (var index = 0; index < images.Length; index++) await AssertImagesCloseAsync(expectedImages[index], images[index], 0.001);
        }
    }

    await using (var nativeInput = File.OpenRead(nativeOfdPath))
    {
        var native = await new OfdReader().ReadAsync(nativeInput);
        if (native.Pages.Count != 2 ||
            native.Pages.Any(page => page.Elements.OfType<OfdImageElement>().Any()) ||
            native.Pages.Any(page => !page.Elements.OfType<OfdTextElement>().Any()) ||
            native.CustomTags.GetValueOrDefault("source-text-origin") != "DOCX/OpenXML")
        {
            throw new InvalidOperationException(
                $"Native DOCX conversion output is incomplete: {nativeOfdPath}");
        }
    }

    Console.WriteLine("[E2E] Native and dual-layer DOCX -> OFD flows passed with 2 pages.");
}

static void AssertComplexUpstreamSample(OfdDocumentPackage parsed, string samplePath)
{
    var pageTextObjects = parsed.Pages.Sum(page => page.Elements.OfType<OfdTextElement>().Count());
    var templateTextObjects = parsed.Pages.Sum(page => page.Templates.Sum(template =>
        template.Elements.OfType<OfdTextElement>().Count()));
    var pathObjects = parsed.Pages.Sum(page => page.Templates.Sum(template => template.Elements
        .OfType<OfdPathElement>()
        .Count()));
    var layers = parsed.Pages.Sum(page =>
        page.Elements.Concat(page.Templates.SelectMany(template => template.Elements))
            .Where(element => !string.IsNullOrWhiteSpace(element.LayerId))
            .Select(element => element.LayerId)
            .Distinct(StringComparer.Ordinal)
            .Count());
    var templates = parsed.Pages.Sum(page => page.PreservedPageElements.Count(xml =>
        xml.Contains("Template", StringComparison.Ordinal)));

    if (pageTextObjects != 560 || templateTextObjects != 121 || pathObjects != 79 || layers != 10 || templates != 5)
    {
        throw new InvalidOperationException(
            $"Complex upstream sample was only partially parsed: {samplePath}. " +
            $"pageText={pageTextObjects}, templateText={templateTextObjects}, " +
            $"paths={pathObjects}, layers={layers}, templates={templates}");
    }

    if (!parsed.PreservedEntries.Keys.Any(path => path.Contains("/Annots/", StringComparison.OrdinalIgnoreCase)))
    {
        throw new InvalidOperationException($"Annotation resources were not preserved: {samplePath}");
    }
}

static async Task ValidateComplexSampleArtifactsAsync(
    string samplePath,
    string outputDir)
{
    await using (var signatureInput = File.OpenRead(samplePath))
    {
        var signatureReport = await new OfdSignatureVerifier()
            .VerifyAsync(signatureInput);
        if (!signatureReport.ReferenceIntegrityValid)
        {
            throw new InvalidOperationException(
                $"Upstream signature reference integrity failed: {samplePath}");
        }
    }

    await using (var invalidInput = File.OpenRead(samplePath))
    await using (var discarded = new MemoryStream())
    {
        try
        {
            await new OfdToPdfConverter().ConvertAsync(invalidInput, discarded);
            throw new InvalidOperationException("Unresolved signed-seal ColorSpace was silently accepted.");
        }
        catch (NotSupportedException error) when (error.Message.Contains("ColorSpace '4'", StringComparison.Ordinal))
        {
            // Reference integrity was checked on the original bytes above.
        }
    }

    // Make a render-only copy with the seal's two mistyped color references
    // corrected. This copy is not a valid signed document.
    var renderCopy = Path.Combine(outputDir, "upstream-999-render-only.ofd");
    await CreateRenderableUpstreamCopyAsync(samplePath, renderCopy);
    var pdfPath = Path.Combine(outputDir, "upstream-999-render-only.pdf");
    await using (var pdfInput = File.OpenRead(renderCopy))
    await using (var pdfOutput = File.Create(pdfPath))
    {
        await new OfdToPdfConverter().ConvertAsync(pdfInput, pdfOutput);
    }

    var pdfPngPath = await RenderFirstPageAsync(
        pdfPath,
        outputDir,
        "upstream-999-page1-pdf");
    await AssertSignedSealRegionAsync(pdfPngPath);

    var svgPath = Path.Combine(outputDir, "upstream-999-page1.svg");
    await using (var svgInput = File.OpenRead(samplePath))
    await using (var svgOutput = File.Create(svgPath))
    {
        await new OfdToSvgConverter().ConvertAsync(svgInput, svgOutput);
    }

    var pngPath = Path.Combine(outputDir, "upstream-999-page1-svg.png");
    var render = await RunProcessAsync(
        "rsvg-convert",
        $"--background-color white --output \"{pngPath}\" \"{svgPath}\"");
    if (render.ExitCode != 0)
    {
        throw new InvalidOperationException(
            $"rsvg-convert failed for {svgPath}: {render.Error}");
    }

    await AssertImageHasContentAsync(pngPath);
    Console.WriteLine(
        "[E2E] Original SM3 references and invalid seal-color rejection passed; repaired render-only seal PDF and original SVG passed visual checks.");
}

static async Task CreateRenderableUpstreamCopyAsync(string sourcePath, string destinationPath)
{
    const string signaturePath = "Doc_0/Signs/Sign_0/Signature.xml";
    const string signedValuePath = "Doc_0/Signs/Sign_0/SignedValue.dat";
    using var source = ZipFile.OpenRead(sourcePath);
    byte[] signedValue;
    using (var input = source.GetEntry(signedValuePath)!.Open())
    using (var buffer = new MemoryStream()) { await input.CopyToAsync(buffer); signedValue = buffer.ToArray(); }
    var start = -1; var end = -1;
    for (var i = 0; i <= signedValue.Length - 4; i++)
    {
        if (start < 0 && signedValue.AsSpan(i, 4).SequenceEqual(new byte[] { 0x50, 0x4b, 0x03, 0x04 })) start = i;
        if (signedValue.AsSpan(i, 4).SequenceEqual(new byte[] { 0x50, 0x4b, 0x05, 0x06 })) end = i;
    }
    if (start < 0 || end < start || end + 22 > signedValue.Length)
        throw new InvalidDataException("Expected an embedded OFD seal ZIP.");
    var sealLength = checked(end + 22 + BitConverter.ToUInt16(signedValue, end + 20) - start);
    if (start + sealLength > signedValue.Length) throw new InvalidDataException("Embedded seal ZIP is truncated.");
    byte[] repairedSeal;
    using (var sealInput = new ZipArchive(new MemoryStream(signedValue, start, sealLength), ZipArchiveMode.Read))
    using (var sealOutput = new MemoryStream())
    {
        using (var copy = new ZipArchive(sealOutput, ZipArchiveMode.Create, leaveOpen: true))
        {
            var changed = 0;
            foreach (var entry in sealInput.Entries)
            {
                byte[] bytes;
                using (var input = entry.Open())
                using (var buffer = new MemoryStream()) { await input.CopyToAsync(buffer); bytes = buffer.ToArray(); }
                if (entry.FullName == "Doc_0/Pages/Page_0/Content.xml")
                {
                    var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
                    foreach (var color in xml.Descendants().Where(node => node.Name.LocalName == "StrokeColor" && node.Attribute("ColorSpace")?.Value == "4"))
                    { color.SetAttributeValue("ColorSpace", "5"); changed++; }
                    bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
                }
                using var target = copy.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
            }
            if (changed != 2) throw new InvalidDataException($"Expected two invalid seal color references, found {changed}.");
        }
        repairedSeal = sealOutput.ToArray();
    }
    await using var destination = File.Create(destinationPath);
    using var output = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);
    foreach (var entry in source.Entries)
    {
        byte[] bytes;
        using (var input = entry.Open())
        using (var buffer = new MemoryStream()) { await input.CopyToAsync(buffer); bytes = buffer.ToArray(); }
        if (entry.FullName == signaturePath)
        {
            var xml = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            var signedInfo = xml.Descendants().Single(node => node.Name.LocalName == "SignedInfo");
            signedInfo.AddFirst(new XElement(signedInfo.Name.Namespace + "Seal", new XAttribute("BaseLoc", "Seal.ofd")));
            bytes = Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
        }
        using var target = output.CreateEntry(entry.FullName).Open(); await target.WriteAsync(bytes);
    }
    using var sealTarget = output.CreateEntry("Doc_0/Signs/Sign_0/Seal.ofd").Open();
    await sealTarget.WriteAsync(repairedSeal);
}

static string ResolveRepoRoot()
{
    var configuredRoot = Environment.GetEnvironmentVariable("OFDRW_REPO_ROOT");
    if (!string.IsNullOrWhiteSpace(configuredRoot))
    {
        return configuredRoot;
    }

    var current = new DirectoryInfo(AppContext.BaseDirectory);
    while (current is not null)
    {
        if (File.Exists(Path.Combine(current.FullName, "Ofdrw.Net.sln")))
        {
            return current.FullName;
        }

        current = current.Parent;
    }

    throw new InvalidOperationException("Unable to locate repository root. Set OFDRW_REPO_ROOT.");
}

static async Task ValidatePdfSampleAsync(string samplePdf, string outputDir)
{
    var sampleName = Path.GetFileNameWithoutExtension(samplePdf);
    var sampleOutputDir = Path.Combine(outputDir, "upstream-ofdrw", sampleName);
    Directory.CreateDirectory(sampleOutputDir);

    var ofdPath = Path.Combine(sampleOutputDir, $"{sampleName}.ofd");
    var pdfPath = Path.Combine(sampleOutputDir, $"{sampleName}.roundtrip.pdf");

    // This smoke checks raster appearance and page size. Imported private-use
    // CID text has no source font bytes; semantic re-export is tested separately
    // and must refuse unverified glyphs rather than drawing fallback boxes.
    var pdfToOfd = new PdfToOfdConverter(new Ofdrw.Net.Converter.Pdf.PdfToOfdOptions
    {
        TextLayerMode = Ofdrw.Net.Converter.Pdf.PdfTextLayerMode.None
    });
    await using (var pdfInput = File.OpenRead(samplePdf))
    await using (var ofdOutput = File.Create(ofdPath))
    {
        await pdfToOfd.ConvertAsync(pdfInput, ofdOutput);
    }

    var reader = new OfdReader();
    OfdDocumentPackage parsed;
    await using (var ofdInput = File.OpenRead(ofdPath))
    {
        parsed = await reader.ReadAsync(ofdInput);
    }

    if (parsed.Pages.Count == 0)
    {
        throw new InvalidOperationException($"Converted OFD has no pages: {ofdPath}");
    }

    if (parsed.Pages.Any(page => page.Elements.OfType<OfdImageElement>().All(image => image.Data.Length == 0)))
    {
        throw new InvalidOperationException($"Converted OFD has a page without renderable image data: {ofdPath}");
    }

    var ofdToPdf = new OfdToPdfConverter();
    await using (var ofdInput = File.OpenRead(ofdPath))
    await using (var pdfOutput = File.Create(pdfPath))
    {
        await ofdToPdf.ConvertAsync(ofdInput, pdfOutput);
    }

    using var sourcePdf = UglyToad.PdfPig.PdfDocument.Open(samplePdf);
    using var roundtripPdf = UglyToad.PdfPig.PdfDocument.Open(pdfPath);
    var pageCount = sourcePdf.NumberOfPages;
    if (pageCount == 0 || roundtripPdf.NumberOfPages != pageCount || parsed.Pages.Count != pageCount)
        throw new InvalidOperationException(
            $"Raster roundtrip page count mismatch: source={pageCount}, OFD={parsed.Pages.Count}, output={roundtripPdf.NumberOfPages}: {samplePdf}");

    const int dpi = 144;
    // PDFium import and Poppler readback differ in antialiasing and fractional
    // resampling: the reviewed 11-page Test/Test3/ff corpus peaks at 0.0942145
    // (dense Chinese Test page 5). This fixed 0.10 raster-profile limit is not
    // pixel equality; renderer changes require fresh calibration and review.
    const double tolerance = 0.10;
    var resultsPath = Path.Combine(sampleOutputDir, "page-comparisons.csv");
    await File.WriteAllTextAsync(resultsPath, "page,source_pages,ofd_pages,output_pages,dpi,rmse,tolerance,result\n");
    for (var page = 1; page <= pageCount; page++)
    {
        var pageOutputDir = Path.Combine(sampleOutputDir, $"page-{page:D3}");
        Directory.CreateDirectory(pageOutputDir);
        var sourcePng = await RenderPageAsync(samplePdf, page, dpi, pageOutputDir, "source");
        var roundtripPng = await RenderPageAsync(pdfPath, page, dpi, pageOutputDir, "roundtrip");
        await AssertImageHasContentAsync(sourcePng);
        await AssertImageHasContentAsync(roundtripPng);
        var error = await AssertImagesCloseAsync(sourcePng, roundtripPng, tolerance,
            Path.Combine(pageOutputDir, "diff.png"));
        await File.AppendAllTextAsync(resultsPath,
            FormattableString.Invariant($"{page},{pageCount},{parsed.Pages.Count},{roundtripPdf.NumberOfPages},{dpi},{error:R},{tolerance:R},passed\n"));
        Console.WriteLine(FormattableString.Invariant(
            $"[E2E] Raster roundtrip {Path.GetFileName(samplePdf)} page {page}/{pageCount}: RMSE {error:R} <= {tolerance:R} at {dpi} DPI."));
        if (page == 1)
            await ValidateRasterComparisonControlsAsync(sourcePng, pageOutputDir, tolerance);
    }

    Console.WriteLine($"[E2E] PDF raster roundtrip passed all {pageCount} pages: {Path.GetFileName(samplePdf)} -> {ofdPath}");
}

static async Task ValidateRasterComparisonControlsAsync(string sourcePng, string outputDir, double tolerance)
{
    var size = await IdentifySizeAsync(sourcePng);
    var alteredPath = Path.Combine(outputDir, "negative-content.png");
    // Preserve size and nonblank content while replacing the central quarter
    // of the page, including existing sample content, with a visible block.
    var alter = await RunProcessAsync("convert",
        $"\"{sourcePng}\" -fill magenta -draw \"rectangle {size.Width / 4},{size.Height / 4} {size.Width * 3 / 4},{size.Height * 3 / 4}\" \"{alteredPath}\"");
    if (alter.ExitCode != 0)
        throw new InvalidOperationException($"Content negative control generation failed: {alter.Error}");
    await AssertImageHasContentAsync(alteredPath);
    if (await IdentifySizeAsync(alteredPath) != size)
        throw new InvalidOperationException("Content negative control changed image dimensions.");
    var resultsPath = Path.Combine(outputDir, "negative-controls.txt");
    try
    {
        await AssertImagesCloseAsync(sourcePng, alteredPath, tolerance,
            Path.Combine(outputDir, "negative-content-diff.png"));
    }
    catch (RenderedPageDifferenceException rejected)
    {
        await File.WriteAllTextAsync(resultsPath, FormattableString.Invariant(
            $"Content: expected rejection, rejected at RMSE {rejected.Error:R} > {tolerance:R}; dimensions {size.Width}x{size.Height} unchanged; nonblank checked.\n"));
        Console.WriteLine(FormattableString.Invariant(
            $"[E2E] Content negative control rejected by raster comparator: RMSE {rejected.Error:R} > {tolerance:R}."));
        await ValidateDimensionControlAsync(sourcePng, outputDir, size, tolerance, resultsPath);
        return;
    }
    throw new InvalidOperationException("Raster comparator accepted the changed-content negative control.");
}

static async Task ValidateDimensionControlAsync(string sourcePng, string outputDir,
    (int Width, int Height) size, double tolerance, string resultsPath)
{
    var resizedPath = Path.Combine(outputDir, "negative-dimensions.png");
    var resize = await RunProcessAsync("convert",
        $"\"{sourcePng}\" -crop {size.Width - 8}x{size.Height}+0+0 +repage \"{resizedPath}\"");
    if (resize.ExitCode != 0)
        throw new InvalidOperationException($"Dimension negative control generation failed: {resize.Error}");
    await AssertImageHasContentAsync(resizedPath);
    try
    {
        await AssertImagesCloseAsync(sourcePng, resizedPath, tolerance);
    }
    catch (InvalidOperationException rejected) when (rejected.Message.StartsWith("Rendered image size mismatch.", StringComparison.Ordinal))
    {
        await File.AppendAllTextAsync(resultsPath, $"Dimensions: expected rejection, rejected: {rejected.Message}\n");
        Console.WriteLine($"[E2E] Dimension negative control rejected by raster comparator: {rejected.Message}");
        return;
    }
    throw new InvalidOperationException("Raster comparator accepted the dimension-mismatch negative control.");
}

static async Task<string> RenderPageAsync(string pdfPath, int page, int dpi, string outputDir, string name)
{
    var outputPrefix = Path.Combine(outputDir, name);
    var result = await RunProcessAsync("pdftoppm", $"-f {page} -l {page} -r {dpi} -png -singlefile \"{pdfPath}\" \"{outputPrefix}\"");
    var pngPath = outputPrefix + ".png";
    if (result.ExitCode != 0 || !File.Exists(pngPath))
        throw new InvalidOperationException($"pdftoppm failed for page {page} of {pdfPath}: {result.Error}");
    return pngPath;
}

static async Task<string> RenderFirstPageAsync(string pdfPath, string outputDir, string name)
{
    var outputPrefix = Path.Combine(outputDir, name);
    var result = await RunProcessAsync("pdftoppm", $"-f 1 -l 1 -png -singlefile \"{pdfPath}\" \"{outputPrefix}\"");
    if (result.ExitCode != 0)
    {
        throw new InvalidOperationException($"pdftoppm failed for {pdfPath}: {result.Error}");
    }

    var pngPath = outputPrefix + ".png";
    if (!File.Exists(pngPath))
    {
        throw new InvalidOperationException($"pdftoppm did not produce {pngPath}");
    }

    return pngPath;
}

static async Task AssertImageHasContentAsync(string imagePath)
{
    var result = await RunProcessAsync("convert", $"\"{imagePath}\" -colorspace Gray -format \"%[fx:standard_deviation]\" info:");
    if (result.ExitCode != 0)
    {
        throw new InvalidOperationException($"Required image content validation failed: {result.Error}");
    }

    if (!double.TryParse(result.Output, NumberStyles.Float, CultureInfo.InvariantCulture, out var standardDeviation) || standardDeviation <= 0.0001d)
    {
        throw new InvalidOperationException($"Rendered image appears blank: {imagePath}");
    }
}

static async Task AssertImagesSameSizeAsync(string expectedPath, string actualPath)
{
    var expected = await IdentifySizeAsync(expectedPath);
    var actual = await IdentifySizeAsync(actualPath);
    if (Math.Abs(expected.Width - actual.Width) > 1 || Math.Abs(expected.Height - actual.Height) > 1)
    {
        throw new InvalidOperationException($"Rendered image size mismatch. expected={expected.Width} x {expected.Height}, actual={actual.Width} x {actual.Height}");
    }
}

static async Task<double> AssertImagesCloseAsync(string expected, string actual, double tolerance, string? diffPath = null)
{
    await AssertImagesSameSizeAsync(expected, actual);
    var result = await RunProcessAsync("compare", $"-metric RMSE \"{expected}\" \"{actual}\" \"{diffPath ?? "null:"}\"");
    var match = System.Text.RegularExpressions.Regex.Match(result.Error, @"\(([0-9.eE+-]+)\)");
    if (result.ExitCode > 1 || !match.Success ||
        !double.TryParse(match.Groups[1].Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var error) || !double.IsFinite(error))
        throw new InvalidOperationException($"Image comparison failed: {expected} vs {actual}; {result.Error}");
    if (error > tolerance)
        throw new RenderedPageDifferenceException(error,
            $"Rendered pages differ beyond {tolerance}: {expected} vs {actual}; {result.Error}");
    return error;
}

static async Task AssertSignedSealRegionAsync(string imagePath)
{
    var size = await IdentifySizeAsync(imagePath);
    var x = (int)Math.Floor(size.Width * 90d / 210d);
    var y = (int)Math.Floor(size.Height * 8d / 140d);
    var width = Math.Max(1, (int)Math.Ceiling(size.Width * 30d / 210d));
    var height = Math.Max(1, (int)Math.Ceiling(size.Height * 20d / 140d));
    var result = await RunProcessAsync(
        "convert",
        $"\"{imagePath}\" -crop {width}x{height}+{x}+{y} " +
        "-format \"%[fx:mean.r-mean.g]\" info:");
    if (result.ExitCode != 0 ||
        !double.TryParse(
            result.Output,
            NumberStyles.Float,
            CultureInfo.InvariantCulture,
            out var redExcess) ||
        redExcess <= 0.1d)
    {
        throw new InvalidOperationException(
            $"Signed seal is missing or not visibly red in {imagePath}. " +
            $"redExcess={result.Output}; error={result.Error}");
    }
}

static async Task<(int Width, int Height)> IdentifySizeAsync(string imagePath)
{
    var result = await RunProcessAsync("identify", $"-format \"%w x %h\" \"{imagePath}\"");
    if (result.ExitCode != 0)
    {
        throw new InvalidOperationException($"ImageMagick identify failed for {imagePath}: {result.Error}");
    }

    var parts = result.Output.Split('x', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 2 || !int.TryParse(parts[0], out var width) || !int.TryParse(parts[1], out var height))
    {
        throw new InvalidOperationException($"Unable to parse ImageMagick size output for {imagePath}: {result.Output}");
    }

    return (width, height);
}

static async Task<(int ExitCode, string Output, string Error)> RunProcessAsync(string fileName, string arguments)
{
    var variable = "OFDRW_" + fileName.Replace('-', '_').ToUpperInvariant();
    var configured = Environment.GetEnvironmentVariable(variable);
    if (!string.IsNullOrWhiteSpace(configured))
    {
        if (!Path.IsPathFullyQualified(configured) || !File.Exists(configured))
            throw new InvalidOperationException($"{variable} must name an existing absolute executable path.");
        fileName = configured;
    }
    else if (OperatingSystem.IsWindows() && fileName is "convert" or "identify" or "compare")
    {
        // Windows ships an unrelated convert.exe. Never search PATH for that name.
        arguments = fileName + " " + arguments;
        fileName = Environment.GetEnvironmentVariable("OFDRW_MAGICK") ?? "magick";
    }
    using var process = new Process
    {
        StartInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        }
    };

    try
    {
        process.Start();
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException($"{fileName} is required for visual validation.", ex);
    }

    using var timeout = new CancellationTokenSource(TimeSpan.FromMinutes(2));
    var stdout = process.StandardOutput.ReadToEndAsync(timeout.Token);
    var stderr = process.StandardError.ReadToEndAsync(timeout.Token);
    try
    {
        await process.WaitForExitAsync(timeout.Token);
        await Task.WhenAll(stdout, stderr);
        return (process.ExitCode, stdout.Result.Trim(), stderr.Result.Trim());
    }
    catch (OperationCanceledException exception)
    {
        if (!process.HasExited) process.Kill(entireProcessTree: true);
        await process.WaitForExitAsync();
        try { await Task.WhenAll(stdout, stderr); } catch (OperationCanceledException) { }
        throw new TimeoutException($"Visual validation process timed out: {fileName}", exception);
    }
}

sealed class RenderedPageDifferenceException(double error, string message) : InvalidOperationException(message)
{
    public double Error { get; } = error;
}

sealed class HostFailureResolver(PdfSharpCore.Fonts.IFontResolver host, string repoRoot) : PdfSharpCore.Fonts.IFontResolver
{
    private const string Prefix = "Ofdrw-HostFailure-";
    public string DefaultFontName => host.DefaultFontName;
    public PdfSharpCore.Fonts.FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
    {
        if (familyName == Prefix + "Resolve") throw new ArgumentException("Synthetic unavailable host family.");
        return familyName.StartsWith(Prefix, StringComparison.Ordinal)
            ? new PdfSharpCore.Fonts.FontResolverInfo(familyName)
            : host.ResolveTypeface(familyName, isBold, isItalic);
    }
    public byte[] GetFont(string faceName) => faceName switch
    {
        Prefix + "Read" => throw new IOException("Synthetic unavailable font file."),
        Prefix + "Ttc" => File.ReadAllBytes(Path.Combine(repoRoot, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "testdata", "fonts", "style-collection.ttc")),
        Prefix + "Broken" => [0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        _ => host.GetFont(faceName)
    };
}
