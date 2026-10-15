using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;

var root = args.Length > 0 ? Path.GetFullPath(args[0]) : Directory.GetCurrentDirectory();
var output = args.Length > 1 ? Path.GetFullPath(args[1]) : Path.Combine(root, "artifacts/issue05/candidate"); Directory.CreateDirectory(output);
var fonts = Path.Combine(root, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts");
var chinese = File.ReadAllBytes(Path.Combine(fonts, "LXGWWenKai-Regular.ttf"));
var latin = File.ReadAllBytes(Path.Combine(fonts, "NotoSans-Regular.ttf"));
var results = new List<object>();
const string probe = "字体子集与复用 中文比例字体 AV office ffi é e\u0301 12345";
foreach (var full in new[] { true, false })
{
    var package = new OfdDocumentPackage(); package.Options.Metadata.CreationDate = DateTimeOffset.UnixEpoch;
    package.Options.FontEmbedding.Mode = full ? OfdFontEmbeddingMode.Full : OfdFontEmbeddingMode.SubsetWhenSafe;
    package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "LXGW WenKai", Data = chinese });
    package.Fonts.Add(new OfdFontResource { Id = "11", FontName = "LXGW WenKai", Bold = true, Data = chinese.ToArray() });
    package.Fonts.Add(new OfdFontResource { Id = "12", FontName = "LXGW WenKai", Italic = true, Data = chinese.ToArray() });
    package.Fonts.Add(new OfdFontResource { Id = "13", FontName = "Noto Sans", Data = latin });
    for (var pageIndex = 0; pageIndex < 2; pageIndex++)
    {
        var page = new OfdPage { Index = pageIndex, WidthMillimeters = 210, HeightMillimeters = 297 }; package.Pages.Add(page);
        Add(probe, "10", 25, 5);
        Add(pageIndex == 0 ? "第一页" : "第二页增加字形：测试对齐分页", "10", 35, 4);
        Add("局部粗体 Bold text", "11", 45, 5, 700, false, new OfdColor(170, 30, 30));
        Add("局部斜体 Italic baseline", "12", 65, 5, 400, true, new OfdColor(30, 60, 170));
        Add("AV office ffi é e\u0301: composite / ligature probe", "13", 85, 5);
        page.Elements.Add(new OfdPathElement { XMillimeters = 20, YMillimeters = 110, WidthMillimeters = 170, HeightMillimeters = 32,
            AbbreviatedData = "M 0 0 L 170 0 L 170 32 L 0 32 C", Fill = true, FillColor = new OfdColor(225, 235, 250), Stroke = true });
        Add("表格底色 Border / fill 左对齐", "10", 120, 5);
        Add("页码 " + (pageIndex + 1), "10", 270, 4);
        void Add(string text, string id, double y, double size, int weight = 400, bool italic = false, OfdColor? color = null) =>
            page.Elements.Add(new OfdTextElement { FontResourceId = id, FontName = id == "13" ? "Noto Sans" : "LXGW WenKai",
                Text = text, XMillimeters = 20, YMillimeters = y, WidthMillimeters = 175, HeightMillimeters = 12,
                FontSizeMillimeters = size, Weight = weight, Italic = italic, FillColor = color ?? OfdColor.Black });
    }
    var name = full ? "full" : "subset"; var ofdPath = Path.Combine(output, name + ".ofd");
    OfdPackageWriteResult report;
    using (var destination = File.Create(ofdPath)) report = await new OfdPackageWriter().WriteWithResultAsync(package, destination);
    using var input = File.OpenRead(ofdPath); var loaded = await new OfdReader().ReadAsync(input);
    if (!new OfdTextExtractor().Extract(loaded).Contains(probe)) throw new Exception("OFD text lost.");
    using var zip = ZipFile.OpenRead(ofdPath);
    var payloads = zip.Entries.Where(entry => entry.FullName.EndsWith(".ttf") || entry.FullName.EndsWith(".otf")).ToArray();
    if (payloads.Length != 2) throw new Exception("Expected two shared face payloads for four style resources.");
    if (!full && report.FontEmbedding.Any(font => !font.IsSubset || font.PayloadBytes >= font.SourceBytes / 2)) throw new Exception("Real subset reduction failed: " + JsonSerializer.Serialize(report.FontEmbedding));
    results.Add(new { name, ofdBytes = new FileInfo(ofdPath).Length, ofdSha256 = Hash(ofdPath), resourceAliases = loaded.Fonts.Count,
        uniquePayloads = payloads.Length, report.FontEmbedding, report.Diagnostics });
    await Export(ofdPath, name, loaded.Pages.Count);
}
// Use the deterministic DOCX layout fixture, with only its font family changed
// to the exact redistributable source above. Never export direct DOCX->PDF.
var fixture = Path.Combine(root, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx");
var source = Path.Combine(output, "generated-layout-fixed-font.docx");
File.Delete(source);
using (var original = ZipFile.OpenRead(fixture))
using (var converted = ZipFile.Open(source, ZipArchiveMode.Create))
{
    XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    foreach (var entry in original.Entries)
    {
        using var read = entry.Open(); using var write = converted.CreateEntry(entry.FullName).Open();
        if (entry.FullName.StartsWith("word/") && entry.FullName.EndsWith(".xml"))
        {
            var xml = XDocument.Load(read);
            foreach (var node in xml.Descendants(w + "rFonts"))
            {
                foreach (var attribute in node.Attributes().ToArray()) attribute.Remove();
                foreach (var attribute in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) node.SetAttributeValue(w + attribute, "LXGW WenKai");
            }
            xml.Save(write);
        }
        else read.CopyTo(write);
    }
}
foreach (var mode in new[] { "native", "default" })
{
    var options = new DocxConversionOptions { Engine = DocxConversionEngine.BuiltIn, UseInstalledMicrosoftOfficeFonts = false };
    if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
    options.FontDirectories.Add(fonts); options.FontFallbackFamilies.Clear(); options.FontFallbackFamilies.Add("LXGW WenKai");
    var ofd = Path.Combine(output, "docx-" + mode + ".ofd");
    using (var read = File.OpenRead(source)) using (var write = File.Create(ofd)) await new DocxToOfdConverter(options).ConvertAsync(read, write);
    using var input = File.OpenRead(ofd); var package = await new OfdReader().ReadAsync(input);
    var text = new OfdTextExtractor().Extract(package); if (!text.Contains("中文") || !text.Contains("Alpha")) throw new Exception("DOCX text missing.");
    var originalText = new StringBuilder(); using (var original = ZipFile.OpenRead(source))
    using (var document = original.GetEntry("word/document.xml")!.Open())
    {
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (var value in XDocument.Load(document).Descendants(w + "t").Select(node => node.Value))
            if (!System.Text.RegularExpressions.Regex.Replace(text, @"\s+", "").Contains(System.Text.RegularExpressions.Regex.Replace(value, @"\s+", ""))) throw new Exception("DOCX source text fragment missing: " + value);
    }
    results.Add(new { name = "docx-" + mode, ofdBytes = new FileInfo(ofd).Length, ofdSha256 = Hash(ofd), pages = package.Pages.Count,
        fonts = package.Fonts.Select(font => new { font.FontName, bytes = font.Data.Length, identity = Convert.ToHexString(SHA256.HashData(font.Data)) }) });
    File.WriteAllText(Path.Combine(output, "docx-" + mode + ".txt"), text);
    await Export(ofd, "docx-" + mode, package.Pages.Count);
}
File.WriteAllText(Path.Combine(output, "font-comparison.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
async Task Export(string ofd, string name, int pages)
{
    using (var input = File.OpenRead(ofd)) using (var pdf = File.Create(Path.Combine(output, name + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(input, pdf);
    for (var page = 0; page < pages; page++)
        using (var input = File.OpenRead(ofd)) using (var svg = File.Create(Path.Combine(output, name + $"-page-{page + 1}.svg")))
            await new OfdToSvgConverter().ConvertAsync(input, svg, page);
}
