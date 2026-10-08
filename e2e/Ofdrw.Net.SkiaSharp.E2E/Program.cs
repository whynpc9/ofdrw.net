using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Converter.Pdf.Converters;
using SkiaSharp;

// A producer that once targeted SKCanvas must explicitly migrate its output
// boundary to this cooperating sink. Render(SKCanvas) is not transparently intercepted.
if (args.Length != 2) throw new ArgumentException("Usage: Skia E2E output-directory licensed-font.ttf");
var output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
var bytes = File.ReadAllBytes(args[1]); using var fontData = SKData.CreateCopy(bytes); using var face = SKTypeface.FromData(fontData);
var adapted = Package(); var direct = Package(); var original = new List<string>();
for (var index = 0; index < 2; index++)
{
    var page = Page(adapted, index); var control = Page(direct, index);
    using var bitmap = new SKBitmap(840, 1191); using var canvas = new SKCanvas(bitmap); canvas.Clear(SKColors.White);
    canvas.Scale(144f / 25.4f, 144f / 25.4f);
    var events = new EventSink();
    Producer.Render(index, new NativeSink(canvas), face);
    Producer.Render(index, events, face);
    Producer.Render(index, new DirectSink(new OfdGraphics(direct, control)), face);
    OfdSkiaAdapter.Append(adapted, page, events.Events);
    original.AddRange(events.Texts);
    using var png = bitmap.Encode(SKEncodedImageFormat.Png, 100); File.WriteAllBytes(Path.Combine(output, $"skia-{index + 1}.png"), png.ToArray());
}
foreach (var (name, package) in new[] { ("adapted", adapted), ("direct04", direct) })
{
    var ofdPath = Path.Combine(output, name + ".ofd");
    await using (var stream = File.Create(ofdPath)) await new OfdPackageWriter().WriteAsync(package, stream);
    OfdDocumentPackage read;
    await using (var stream = File.OpenRead(ofdPath)) read = await new OfdReader().ReadAsync(stream);
    var text = string.Concat(read.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().SelectMany(t => t.Runs).Select(r => r.Text));
    if (text != string.Concat(original)) throw new Exception("Original Unicode mismatch: " + name);
    if (read.Pages.Count != 2 || read.Pages.SelectMany(p => p.Elements).Any(e => e is not (OfdTextElement or OfdPathElement))) throw new Exception("Expected two pages of native PathObject/TextObject only.");
    foreach (var t in read.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>())
        if (t.FontResourceId != read.Fonts.Single().Id) throw new Exception("Unique font binding lost.");
    File.WriteAllText(Path.Combine(output, name + ".txt"), new OfdTextExtractor().Extract(read));
    await using (var input = File.OpenRead(ofdPath))
    await using (var pdf = File.Create(Path.Combine(output, name + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(input, pdf);
}
using (var a = ZipFile.OpenRead(Path.Combine(output, "adapted.ofd")))
using (var b = ZipFile.OpenRead(Path.Combine(output, "direct04.ofd")))
{
    foreach (var entry in a.Entries)
    {
        using var left = entry.Open(); using var right = b.GetEntry(entry.FullName)!.Open(); using var l = new MemoryStream(); using var r = new MemoryStream();
        left.CopyTo(l); right.CopyTo(r); if (!l.ToArray().SequenceEqual(r.ToArray())) throw new Exception("04 control package entry differs: " + entry.FullName);
    }
}
await FontStyleProbe.Run(output, bytes, face);
// Retain the actual failed transparent-capture alternatives, not just an assertion about APIs.
var probe = new Dictionary<string, object>();
probe["skcanvas_virtual_draw_methods"] = typeof(SKCanvas).GetMethods().Count(m => m.Name.StartsWith("Draw") && m.IsVirtual);
using (var stream = new SKDynamicMemoryWStream())
{
    using (var canvas = SKSvgCanvas.Create(new SKRect(0, 0, 148, 90), stream))
    {
        using var paint = new SKPaint { Color = SKColors.Black }; using var font = new SKFont(face, 4);
        canvas.DrawText("Skia  native text", 10, 20, font, paint);
        using var filter = RedFilter(); paint.ColorFilter = filter; canvas.DrawRect(10, 30, 20, 20, paint);
    }
    using var data = stream.DetachAsData(); var xml = Encoding.UTF8.GetString(data.ToArray()); File.WriteAllText(Path.Combine(output, "capture-probe.svg"), xml);
    var doc = XDocument.Parse(xml); var emitted = string.Concat(doc.Descendants().Where(e => e.Name.LocalName == "text").Select(e => e.Value.Trim()));
    probe["original"] = "Skia  native text"; probe["svg_text"] = emitted; probe["original_unicode_preserved"] = emitted == "Skia  native text";
    var rect = doc.Descendants().Single(e => e.Name.LocalName == "rect"); probe["svg_filter_rectangle"] = rect.ToString();
}
using (var bitmap = new SKBitmap(32, 32))
{
    using var canvas = new SKCanvas(bitmap); using var paint = new SKPaint { Color = SKColors.Black }; using var filter = RedFilter();
    paint.ColorFilter = filter; canvas.DrawRect(0, 0, 32, 32, paint); var actual = bitmap.GetPixel(16, 16);
    if (actual != SKColors.Red) throw new Exception("Probe filter did not render red."); probe["actual_filtered_pixel"] = actual.ToString();
    using var png = bitmap.Encode(SKEncodedImageFormat.Png, 100); File.WriteAllBytes(Path.Combine(output, "capture-filter-actual.png"), png.ToArray());
}
File.WriteAllText(Path.Combine(output, "capture-probe.json"), JsonSerializer.Serialize(probe, new JsonSerializerOptions { WriteIndented = true }));
var report = new { status = "passed", mode = "cooperating-producer-events", pages = 2, paths = adapted.Pages.Sum(p => p.Elements.OfType<OfdPathElement>().Count()),
    texts = adapted.Pages.Sum(p => p.Elements.OfType<OfdTextElement>().Count()), images = 0, original_unicode_preserved = true, font_sha256 = Hash(bytes),
    source_assembly = AssemblyInfo(typeof(OfdSkiaAdapter).Assembly), skia_assembly = AssemblyInfo(typeof(SKCanvas).Assembly), environment = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
    files = Directory.GetFiles(output).OrderBy(p => p).Select(p => new { file = Path.GetFileName(p), bytes = new FileInfo(p).Length, sha256 = Hash(File.ReadAllBytes(p)) }).ToArray() };
File.WriteAllText(Path.Combine(output, "sample-report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(JsonSerializer.Serialize(report));
OfdDocumentPackage Package() { var package = new OfdDocumentPackage(); package.Fonts.Add(new OfdFontResource { Id = "sample-face", FontName = face.FamilyName, Data = bytes, FileName = "Noto-Regular.ttf", Bold = face.IsBold, Italic = face.IsItalic }); return package; }
static OfdPage Page(OfdDocumentPackage package, int index) { var page = new OfdPage { Index = index, WidthMillimeters = 148, HeightMillimeters = 210 }; package.Pages.Add(page); return page; }
static string Hash(byte[] payload) => Convert.ToHexString(SHA256.HashData(payload)).ToLowerInvariant();
static object AssemblyInfo(Assembly assembly) => new { name = assembly.FullName, location = assembly.Location, sha256 = Hash(File.ReadAllBytes(assembly.Location)) };
static SKColorFilter RedFilter() => SKColorFilter.CreateColorMatrix(new float[] {0,0,0,0,1, 0,0,0,0,0, 0,0,0,0,0, 0,0,0,1,0});

interface ISkiaSink
{
    void Line(SKPoint a, SKPoint b, SKPaint paint, SKMatrix matrix);
    void Rect(SKRect rect, SKPaint paint, SKMatrix matrix);
    void Path(SKPath path, SKPaint paint, SKMatrix matrix);
    void Text(string text, SKPoint baseline, SKFont font, SKPaint paint, float[] advances, SKMatrix matrix);
}
sealed class NativeSink(SKCanvas canvas) : ISkiaSink
{
    void Draw(SKMatrix matrix, Action action) { canvas.Save(); canvas.Concat(matrix); action(); canvas.Restore(); }
    public void Line(SKPoint a, SKPoint b, SKPaint paint, SKMatrix matrix) => Draw(matrix, () => canvas.DrawLine(a, b, paint));
    public void Rect(SKRect rect, SKPaint paint, SKMatrix matrix) => Draw(matrix, () => canvas.DrawRect(rect, paint));
    public void Path(SKPath path, SKPaint paint, SKMatrix matrix) => Draw(matrix, () => canvas.DrawPath(path, paint));
    public void Text(string text, SKPoint baseline, SKFont font, SKPaint paint, float[] advances, SKMatrix matrix) => Draw(matrix, () => canvas.DrawText(text, baseline, SKTextAlign.Left, font, paint));
}
sealed class EventSink : ISkiaSink
{
    public List<SkiaDrawEvent> Events { get; } = []; public List<string> Texts { get; } = [];
    public void Line(SKPoint a, SKPoint b, SKPaint paint, SKMatrix matrix) => Events.Add(SkiaDrawEvent.Line(a, b, paint, matrix));
    public void Rect(SKRect rect, SKPaint paint, SKMatrix matrix) => Events.Add(SkiaDrawEvent.Rectangle(rect, paint, matrix));
    public void Path(SKPath path, SKPaint paint, SKMatrix matrix) => Events.Add(SkiaDrawEvent.Path(path, paint, matrix));
    public void Text(string text, SKPoint baseline, SKFont font, SKPaint paint, float[] advances, SKMatrix matrix) { Events.Add(SkiaDrawEvent.Text(text, baseline, font, paint, "sample-face", advances, matrix)); Texts.Add(text); }
}
sealed class DirectSink(OfdGraphics graphics) : ISkiaSink
{
    static OfdColor Color(SKPaint p) => new(p.Color.Red, p.Color.Green, p.Color.Blue, p.Color.Alpha);
    void Set(SKMatrix m) => graphics.SetTransform(new OfdMatrix(m.ScaleX, m.SkewY, m.SkewX, m.ScaleY, m.TransX, m.TransY));
    public void Line(SKPoint a, SKPoint b, SKPaint p, SKMatrix m) { Set(m); graphics.DrawLine(new OfdPen(Color(p), p.StrokeWidth), a.X, a.Y, b.X, b.Y); }
    public void Rect(SKRect r, SKPaint p, SKMatrix m) { Set(m); if (p.Style == SKPaintStyle.Fill) graphics.FillRectangle(new OfdBrush(Color(p)), r.Left, r.Top, r.Width, r.Height); else graphics.DrawRectangle(new OfdPen(Color(p), p.StrokeWidth), r.Left, r.Top, r.Width, r.Height); }
    public void Path(SKPath path, SKPaint p, SKMatrix m)
    {
        Set(m); var mapped = new OfdGraphicsPath(path.FillType == SKPathFillType.EvenOdd ? OfdFillRule.EvenOdd : OfdFillRule.NonZero); using var iterator = path.CreateRawIterator(); var points = new SKPoint[4];
        for (var v = iterator.Next(points); v != SKPathVerb.Done; v = iterator.Next(points))
            switch (v) { case SKPathVerb.Move: mapped.MoveTo(points[0].X, points[0].Y); break; case SKPathVerb.Line: mapped.LineTo(points[1].X, points[1].Y); break; case SKPathVerb.Quad: mapped.QuadraticTo(points[1].X, points[1].Y, points[2].X, points[2].Y); break; case SKPathVerb.Cubic: mapped.BezierTo(points[1].X, points[1].Y, points[2].X, points[2].Y, points[3].X, points[3].Y); break; case SKPathVerb.Close: mapped.Close(); break; default: throw new Exception("Fixture verb unsupported."); }
        graphics.DrawPath(mapped, p.Style == SKPaintStyle.Stroke ? new OfdPen(Color(p), p.StrokeWidth) : null, p.Style == SKPaintStyle.Fill ? new OfdBrush(Color(p)) : null);
    }
    public void Text(string text, SKPoint baseline, SKFont font, SKPaint paint, float[] advances, SKMatrix matrix) { Set(matrix); graphics.DrawString(text, new OfdFont("Noto Sans CJK SC", font.Size, resourceId: "sample-face"), new OfdBrush(Color(paint)), baseline.X, baseline.Y, advances.Select(v => (double)v).ToArray()); }
}
static class Producer
{
    public static void Render(int index, ISkiaSink sink, SKTypeface face)
    {
        var matrix = SKMatrix.Identity;
        using var pen = new SKPaint { Color = new SKColor(30, 93, 166), Style = SKPaintStyle.Stroke, StrokeWidth = .5f, StrokeMiter = 10, IsAntialias = true };
        using var fill = new SKPaint { Color = new SKColor(232, 241, 249), IsAntialias = true };
        void Text(string text, float x, float y, float size = 3.6f, SKColor? color = null)
        {
            using var font = new SKFont(face, size) { LinearMetrics = true, Hinting = SKFontHinting.None }; using var paint = new SKPaint { Color = color ?? new SKColor(32, 44, 58), IsAntialias = true };
            // Producer knows its simple unshaped text. It supplies the same native advances,
            // in user units, before SKCanvas turns the original Unicode into a blob.
            var parts = StringInfo.ParseCombiningCharacters(text); var widths = new float[Math.Max(0, parts.Length - 1)];
            for (var i = 0; i < widths.Length; i++) widths[i] = font.MeasureText(text.Substring(parts[i], parts[i + 1] - parts[i]));
            sink.Text(text, new(x, y), font, paint, widths, matrix);
        }
        sink.Rect(new(10, 10, 138, 37), fill, matrix); Text(index == 0 ? "合作适配票面 / SKIA EVENTS" : "变换与路径 / AFFINE PATHS", 14, 23, 5.2f, SKColors.Blue);
        Text("Explicit producer events; native OFD text and paths", 14, 32, 2.7f);
        if (index == 0)
        {
            Text("Skia  native text / 双空格原文", 12, 49, 4);
            sink.Rect(new(12, 58, 136, 128), pen, matrix); sink.Rect(new(12, 58, 136, 71), fill, matrix);
            foreach (var y in new[] {71f, 90, 109}) sink.Line(new(12, y), new(136, y), pen, matrix);
            foreach (var x in new[] {82f, 108}) sink.Line(new(x, 58), new(x, 128), pen, matrix);
            Text("项目 Item", 16, 67); Text("数量 Qty", 85, 67, 3); Text("金额 CNY", 111, 67, 3);
            Text("技术文档 Technical", 16, 84, 3.4f); Text("2", 91, 84); Text("128.00", 111, 84, 3);
            Text("矢量图 Vector", 16, 103, 3.4f, SKColors.Blue); Text("1", 91, 103); Text("64.00", 111, 103, 3);
            Text("合计 Total", 16, 121); Text("192.00", 111, 121, 3, SKColors.Red);
            sink.Line(new(12, 151), new(136, 151), pen, matrix); Text("Wi fi / 中文基线", 12, 151, 4);
            Text("局部红色", 12, 165, 3.6f, SKColors.Red); Text("Regular remains regular", 39, 165);
            matrix = new SKMatrix(1, -.15f, 103, .15f, 1, 183, 0, 0, 1); sink.Rect(new(-24, -8, 24, 8), pen, matrix); Text("已付款 PAID", -21, 1, 4, SKColors.Blue); matrix = SKMatrix.Identity;
        }
        else
        {
            Text("Complete event matrix: scale, shear and translation", 12, 46, 2.8f);
            matrix = new SKMatrix(2, .2f, 15, .1f, 1, 57, 0, 0, 1); sink.Line(new(0, 0), new(35, 0), pen, matrix); sink.Line(new(42, 0), new(42, 16), pen, matrix); matrix = SKMatrix.Identity;
            using var curve = new SKPath(); curve.MoveTo(14, 102); curve.CubicTo(28, 73, 45, 121, 62, 102); curve.QuadTo(75, 84, 86, 102); sink.Path(curve, pen, matrix);
            using var sharp = new SKPath(); sharp.MoveTo(105, 102); sharp.LineTo(110, 75); sharp.LineTo(115, 102); pen.StrokeWidth = 1.4f; sink.Path(sharp, pen, matrix); pen.StrokeWidth = .5f; Text("Miter=10", 101, 111, 2.7f);
            using var hole = new SKPath { FillType = SKPathFillType.EvenOdd }; hole.AddRect(new(16, 126, 58, 151)); hole.AddRect(new(26, 133, 48, 144)); fill.Color = new(26, 133, 112); sink.Path(hole, fill, matrix);
            using var solid = new SKPath(); solid.AddRect(new(78, 126, 120, 151)); solid.AddRect(new(88, 133, 110, 144)); sink.Path(solid, fill, matrix); Text("EvenOdd", 18, 122); Text("Winding", 80, 122);
            matrix = new SKMatrix(-1, 0, 125, 0, 1, 0, 0, 0, 1); Text("Mirror / 镜像", 0, 179, 4, SKColors.Blue); matrix = SKMatrix.Identity;
        }
        Text($"Cooperative producer / page {index + 1} of 2", 12, 199, 2.8f);
    }
}
