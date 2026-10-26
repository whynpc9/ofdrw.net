using System.Globalization;
using System.IO.Compression;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;
using System.Reflection;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Packaging;
using PdfSharpCore.Drawing;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[Collection("Font registry")]
public sealed class Audit025RendererRegressionTests
{
    private readonly Xunit.Abstractions.ITestOutputHelper output;
    public Audit025RendererRegressionTests(Xunit.Abstractions.ITestOutputHelper output) => this.output = output;
    private static readonly Lazy<byte[]> Font = new(() => File.ReadAllBytes(Path.Combine(Root(),
        "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf")));
    private static string Root()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (root is not null && !File.Exists(Path.Combine(root.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return root!.FullName;
    }
    private static OfdTextElement Text(string value, double y, string? xDelta = null, string? yDelta = null)
        => new() { Text = value, FontResourceId = "10", FontSizeMillimeters = 4, XMillimeters = 10,
            YMillimeters = y, WidthMillimeters = 140, HeightMillimeters = 12,
            Runs = { new OfdTextRun { Text = value, XMillimeters = 2, YMillimeters = 5, DeltaX = xDelta, DeltaY = yDelta } } };
    private static async Task<byte[]> Input(OfdTextElement text)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit025", Data = Font.Value });
        package.Pages.Add(new OfdPage { WidthMillimeters = 180, HeightMillimeters = 100, Elements = { text } });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); return output.ToArray();
    }
    private static async Task<(byte[] Pdf, XDocument Svg)> Render(byte[] bytes)
    {
        using var input = new MemoryStream(bytes); using var pdf = new MemoryStream(); using var svg = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(input, pdf); input.Position = 0;
        await new OfdToSvgConverter().ConvertAsync(input, svg); svg.Position = 0;
        return (pdf.ToArray(), XDocument.Load(svg));
    }

    [Theory]
    [InlineData(null, null)] [InlineData("", " ")] [InlineData("g 0 3", null)]
    [InlineData("3 5", "1 2")]
    public async Task Audit025_RendererEquivalence(string? dx, string? dy)
    {
        const string value = "e\u0301AB";
        var bytes = await Input(Text(value, 10, dx, dy)); var rendered = await Render(bytes);
        var node = Assert.Single(rendered.Svg.Descendants().Where(n => n.Name.LocalName == "text"));
        Assert.Equal(value, node.Value);
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(rendered.Pdf);
        Assert.Equal(value.Normalize(NormalizationForm.FormC), pdf.GetPage(1).Text.Normalize(NormalizationForm.FormC));
        if (dx == "3 5")
        {
            var spans = node.Elements().ToArray(); Assert.Equal(new[] { "e\u0301", "A", "B" }, spans.Select(n => n.Value));
            Assert.Equal(new[] { 2d, 5, 10 }, spans.Select(n => double.Parse(n.Attribute("x")!.Value, CultureInfo.InvariantCulture)));
            Assert.Equal(new[] { 5d, 6, 8 }, spans.Select(n => double.Parse(n.Attribute("y")!.Value, CultureInfo.InvariantCulture)));
            var letters = pdf.GetPage(1).Letters; var a = letters.Single(l => l.Value == "A"); var b = letters.Single(l => l.Value == "B");
            Assert.Equal(5 * 72 / 25.4, b.StartBaseLine.X - a.StartBaseLine.X, 3);
            Assert.Equal(-2 * 72 / 25.4, b.StartBaseLine.Y - a.StartBaseLine.Y, 3);
        }
        else Assert.Empty(node.Elements());
    }

    [Fact]
    public void Audit025_SvgNoDeltaAllocation()
    {
        XNamespace ns = "http://www.w3.org/2000/svg";
        var method = typeof(OfdToSvgConverter).GetMethod("AddText", BindingFlags.NonPublic | BindingFlags.Static)!;
        var text = Text(new string('A', 4096), 10); var root = new XElement(ns + "svg");
        object[] args = { root, ns, new OfdPage(), text, Array.Empty<OfdFontResource>(),
            new Dictionary<OfdFontResource, (string Family, OpenTypeCmap? Coverage)>(), new FontParseBudget(), CancellationToken.None, null! };
        method.Invoke(null, args); root.RemoveNodes();
        var before = GC.GetAllocatedBytesForCurrentThread(); method.Invoke(null, args);
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        output.WriteLine($"Measured allocation: {allocated} bytes for 4096 characters.");
        Assert.Equal(text.Text, Assert.Single(root.Elements()).Value);
        Assert.True(allocated < 16_384, $"SVG text branch allocated {allocated} bytes.");
    }

    [Fact]
    public void Audit025_PdfNoDeltaAllocation()
    {
        // Supported missing Hangul fillers leave no visible output. This isolates the
        // real run dispatcher from PDFsharp's unrelated per-character text allocation.
        var coverage = new OpenTypeCmap(new OpenTypeFace(Font.Value)); Assert.Equal(0, coverage.Glyph(0x3164));
        var text = Text(new string('\u3164', 4096), 10);
        using var document = new PdfSharpCore.Pdf.PdfDocument(); using var graphics = XGraphics.FromPdfPage(document.AddPage());
        var method = typeof(OfdToPdfConverter).GetMethod("DrawTextRuns", BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] args = { graphics, text, null, XBrushes.Black, 0d, 0d, null, false, false, coverage, CancellationToken.None };
        method.Invoke(null, args);
        var before = GC.GetAllocatedBytesForCurrentThread(); method.Invoke(null, args);
        var allocated = GC.GetAllocatedBytesForCurrentThread() - before;
        output.WriteLine($"Measured allocation: {allocated} bytes for 4096 characters.");
        Assert.True(allocated < 16_384, $"PDF text branch allocated {allocated} bytes.");
    }

    [Fact]
    public async Task Audit025_VisualBoundaries()
    {
        var package = new OfdDocumentPackage(); package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "audit025", Data = Font.Value });
        package.Pages.Add(new OfdPage { WidthMillimeters = 180, HeightMillimeters = 100,
            Elements = { Text("No delta: AV office e\u0301", 8), Text("e\u0301 A B", 30, "5 3 6 3"), Text("A\u202FB", 52, "5 5", "0 1") } });
        using var stream = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, stream);
        var bytes = stream.ToArray(); var rendered = await Render(bytes);
        Assert.Equal(new[] { "No delta: AV office e\u0301", "e\u0301 A B", "A\u202FB" },
            rendered.Svg.Descendants().Where(n => n.Name.LocalName == "text").Select(n => n.Value));
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(rendered.Pdf); Assert.Equal(1, pdf.NumberOfPages);
        if (Environment.GetEnvironmentVariable("OFDRW_PLAN025_ARTIFACTS") is { Length: > 0 } directory)
        {
            Directory.CreateDirectory(directory); await File.WriteAllBytesAsync(Path.Combine(directory, "coverage-boundaries.ofd"), bytes);
            await File.WriteAllBytesAsync(Path.Combine(directory, "coverage-boundaries.pdf"), rendered.Pdf);
            rendered.Svg.Save(Path.Combine(directory, "coverage-boundaries.svg"));
        }
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit025_NativeAndDefaultDocxControl(bool explicitNative)
    {
        var sample = Path.Combine(Root(), "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx");
        using var archive = ZipFile.OpenRead(sample); using var xml = archive.GetEntry("word/document.xml")!.Open();
        var original = string.Concat(XDocument.Load(xml).Descendants().Where(n => n.Name.LocalName == "t").Select(n => n.Value));
        static string Compact(string value) => string.Concat(value.Where(c => !char.IsWhiteSpace(c)));
        var converter = explicitNative ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }) : new DocxToOfdConverter();
        using var input = File.OpenRead(sample); using var full = new MemoryStream();
        var result = await converter.ConvertWithResultAsync(input, full);
        Assert.Equal(DocxToOfdMode.Native, result.Mode); Assert.Equal(2, result.SourcePageCount);
        Assert.True(result.OriginalTextPreserved); Assert.Equal(new[] { 0, 1 }, result.SourcePages);
        full.Position = 0; var package = await new OfdReader().ReadAsync(full);
        var extracted = new OfdTextExtractor().ExtractPages(package);
        Assert.Equal(Compact(original), Compact(string.Concat(extracted)));
        input.Position = 0; using var selected = new MemoryStream();
        var selection = await converter.ConvertWithResultAsync(input, selected, new[] { 1 });
        Assert.Equal(new[] { 1 }, selection.SourcePages); Assert.True(selection.OriginalTextPreserved);
        selected.Position = 0; var selectedPackage = await new OfdReader().ReadAsync(selected);
        Assert.Equal(extracted[1], Assert.Single(new OfdTextExtractor().ExtractPages(selectedPackage)));
        foreach (var (stream, suffix, count) in new[] { (full, "", 2), (selected, "-selected-page2", 1) })
        {
            stream.Position = 0; using var output = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(stream, output);
            using var pdf = UglyToad.PdfPig.PdfDocument.Open(output.ToArray()); Assert.Equal(count, pdf.NumberOfPages);
            if (Environment.GetEnvironmentVariable("OFDRW_PLAN025_ARTIFACTS") is { Length: > 0 } directory)
            {
                Directory.CreateDirectory(directory); var name = explicitNative ? "docx-native" : "docx-default";
                await File.WriteAllBytesAsync(Path.Combine(directory, name + suffix + ".ofd"), stream.ToArray());
                await File.WriteAllBytesAsync(Path.Combine(directory, name + suffix + ".pdf"), output.ToArray());
            }
        }
    }
}
