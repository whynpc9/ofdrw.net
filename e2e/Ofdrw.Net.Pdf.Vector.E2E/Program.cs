using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

if (args.Length != 2) throw new ArgumentException("Usage: output-directory licensed-full-static-Noto.ttf");
var output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
var font = File.ReadAllBytes(args[1]); using var fixture = new PdfFixture(font);
var p1 = "0.91 0.945 0.976 rg 28 490 364 70 re f 0.12 0.365 0.65 rg\n" + fixture.Text("矢量转换票面 / VECTOR TICKET", 42, 532, 17) +
    "0.12 0.17 0.23 rg\n" + fixture.Text("原文与绘制顺序 / Unicode and paint order", 42, 507, 10) +
    fixture.Text("No. 2026-021", 34, 465, 11) + fixture.Text("日期 Date: 2026-10-08", 215, 465, 11) +
    "0.88 0.925 0.965 rg 34 403 352 35 re f 0.42 0.51 0.60 RG .7 w 34 268 352 170 re S\n" +
    "34 403 m 386 403 l S 34 358 m 386 358 l S 34 313 m 386 313 l S 230 268 m 230 438 l S 302 268 m 302 438 l S\n" +
    "0.12 0.17 0.23 rg\n" + fixture.Text("项目 Item", 44, 416, 11) + fixture.Text("数量 Qty", 240, 416, 11) + fixture.Text("金额 CNY", 312, 416, 11) +
    fixture.Text("技术文档 Technical notes", 44, 378, 10) + fixture.Text("2", 263, 378, 11) + fixture.Text("128.00", 318, 378, 11) +
    "0.12 0.365 0.65 rg\n" + fixture.Text("矢量示意 Vector diagram", 44, 333, 10, .2) + fixture.Text("1", 263, 333, 11) + fixture.Text("64.00", 318, 333, 11) +
    "0.12 0.17 0.23 rg\n" + fixture.Text("合计 Total", 44, 283, 11) + "0.73 0.23 0.25 rg\n" + fixture.Text("192.00", 318, 283, 11) +
    "0.12 0.17 0.23 rg\n" + fixture.Text("连续空格: A  B / Wi fi 中文", 34, 225, 12) +
    "0.73 0.23 0.25 rg\n" + fixture.Text("局部红色", 34, 202, 11) + "0.12 0.17 0.23 rg\n" + fixture.Text("Regular remains regular", 110, 202, 11) +
    "q .978 -.208 .208 .978 306 142 cm 0.12 0.365 0.65 RG 1.8 w -65 -18 130 40 re S 0.12 0.365 0.65 rg\n" + fixture.Text("已付款 PAID", -54, -3, 14) + "Q\n" +
    "0.12 0.17 0.23 rg\n" + fixture.Text("Native paths + original text - page 1 / 2", 34, 32, 9);
var p2 = "0.12 0.365 0.65 rg\n" + fixture.Text("矩阵与绘制顺序 / AFFINE AND ORDER", 34, 552, 17) +
    "0.12 0.17 0.23 rg\n" + fixture.Text("Native cubic paths, signed rectangles and fill rules", 34, 527, 10) +
    "q 1 .12 .18 1 12 4 cm 0.91 0.945 0.976 rg 35 370 320 110 re f 0.12 0.17 0.23 rg\n" + fixture.Text("A  B中文 / searchable text", 52, 420, 16) +
    "0.73 0.23 0.25 rg 88 350 18 145 re f Q\n" +
    "0.12 0.365 0.65 RG 2 w 45 305 m 75 355 120 255 155 305 c 180 340 215 300 v S\n" +
    "0.12 0.17 0.23 rg\n" + fixture.Text("Cubic + v shorthand", 42, 275, 10) +
    "0.10 0.52 0.44 rg 42 167 120 72 re 64 185 76 36 re f* 230 167 120 72 re 328 185 -76 36 re f\n" +
    "0.12 0.17 0.23 rg\n" + fixture.Text("EvenOdd hole", 42, 147, 11) + fixture.Text("Signed winding hole", 230, 147, 11) +
    fixture.Text("Restored context - page 2 / 2", 34, 32, 9);
var source = fixture.Create(new[] { p1, p2 }); File.WriteAllBytes(Path.Combine(output, "same-source.pdf"), source);
var reports = new List<object>();
foreach (var mode in new[] { "dual", "vector" })
{
    using var input = new MemoryStream(source); using var ofd = new MemoryStream();
    PdfVectorConversionResult? result = null;
    if (mode == "vector") result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, ofd);
    else await new PdfToOfdConverter(new PdfToOfdOptions { PreferExternalPdfToPpm = true }).ConvertAsync(input, ofd);
    await Export(mode, ofd);
    ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
    var text = string.Concat(package.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).Select(element => element.Text));
    if (mode == "vector" && (!text.Contains("A  B") || !text.Contains("中文") || package.Pages.Any(page => page.Elements.OfType<OfdImageElement>().Any()))) throw new Exception("Native vector text/object acceptance failed.");
    if (mode == "vector" && !package.Fonts.All(resource => resource.Data!.SequenceEqual(font))) throw new Exception("Exact font bytes changed.");
    if (mode == "dual" && package.Pages.Any(page => page.Elements.OfType<OfdImageElement>().Count() != 1 || page.Elements.OfType<OfdPathElement>().Any() || page.Elements.OfType<OfdTextElement>().Any(element => element.FillColor.Alpha != 0))) throw new Exception("Default dual-layer changed.");
    reports.Add(new { mode, pages = package.Pages.Select(page => new { page.Index, page.WidthMillimeters, page.HeightMillimeters, paths = page.Elements.OfType<OfdPathElement>().Count(), texts = page.Elements.OfType<OfdTextElement>().Count(), images = page.Elements.OfType<OfdImageElement>().Count() }), result, text, fontBytes = package.Fonts.Sum(resource => resource.Data?.Length ?? 0) });
}
var unsupportedBodies = new[] { "q 20 20 160 160 re W n " + p1 + " Q", "/GS1 gs " + p1, "q 420 0 0 595 0 0 cm /Im1 Do Q", "", fixture.Text("Pure text / 中文", 34, 520, 15) };
var mixed = fixture.Create(unsupportedBodies); File.WriteAllBytes(Path.Combine(output, "fallback-source.pdf"), mixed);
using (var input = new MemoryStream(mixed))
using (var ofd = new MemoryStream())
{
    var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage, Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } }).ConvertWithResultAsync(input, ofd);
    if (result.Pages.Take(4).Any(page => page.IsNative || page.ImageObjects != 1 || page.PathObjects != 0) || !result.Pages[4].IsNative) throw new Exception("Fallback/native page disposition failed.");
    await Export("fallback", ofd); reports.Add(new { mode = "explicit-rasterize-page", result });
}
File.WriteAllText(Path.Combine(output, "sample-report.json"), JsonSerializer.Serialize(new { fontSha256 = Convert.ToHexStringLower(SHA256.HashData(font)), reports }, new JsonSerializerOptions { WriteIndented = true }));
await ImageHintProbe.Run(output, font);
await NonPaintingProbe.Run(output, font);
await PrecisionResourceProbe.Run(output, font);
await IndexedPaletteProbe.Run(output, font);
Console.WriteLine("Actual native/dual same-PDF and explicit fallback fixtures passed; Preview remains separate.");
async Task Export(string name, MemoryStream ofd)
{
    File.WriteAllBytes(Path.Combine(output, name + ".ofd"), ofd.ToArray()); ofd.Position = 0;
    await using (var pdf = File.Create(Path.Combine(output, name + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
    ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
    for (var index = 0; index < package.Pages.Count; index++)
    { ofd.Position = 0; await using var svg = File.Create(Path.Combine(output, name + "-" + (index + 1) + ".svg")); await new OfdToSvgConverter().ConvertAsync(ofd, svg, index); }
}
