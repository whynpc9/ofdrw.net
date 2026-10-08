using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using PdfSharpCore.Fonts;
using SkiaSharp;

// Original MIT symbol fonts isolate file glyph appearance from style requests.
static class FontStyleProbe
{
    public static async Task Run(string output, byte[] captionBytes, SKTypeface captionFace)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new() { Id = "caption", FontName = captionFace.FamilyName, Data = captionBytes, FileName = "Noto-Regular.ttf" });
        var page = new OfdPage { WidthMillimeters = 148, HeightMillimeters = 210 }; package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page); var label = new OfdBrush(new OfdColor(32, 44, 58));
        graphics.DrawString("文件字形与强调 / FILE GLYPH STYLE", new OfdFont(captionFace.FamilyName, 4.8, resourceId: "caption"), label, 12, 17);
        graphics.DrawString("Adapter (left) / direct04 neutral control (right)", new OfdFont(captionFace.FamilyName, 2.9, resourceId: "caption"), label, 12, 25);
        using var bitmap = new SKBitmap(840, 1191); using var canvas = new SKCanvas(bitmap); canvas.Clear(SKColors.White); canvas.Scale(144f / 25.4f, 144f / 25.4f);
        var styles = new[] { ("regular", false, false), ("semibold", false, false), ("oblique", false, false), ("bold", true, false), ("italic", false, true) };
        var results = new List<object>(); var row = 0;
        foreach (var (name, bold, italic) in styles)
        {
            var bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "style-fonts", name + ".ttf"));
            using var data = SKData.CreateCopy(bytes); using var face = SKTypeface.FromData(data); using var font = new SKFont(face, 8) { LinearMetrics = true, Hinting = SKFontHinting.None };
            using var paint = new SKPaint { Color = SKColors.Black, IsAntialias = true };
            package.Fonts.Add(new() { Id = name, FontName = face.FamilyName, Data = bytes, FileName = name + ".ttf", Bold = bold, Italic = italic });
            var y = 42f + row++ * 29;
            graphics.DrawString($"{name}: OS/2 flags B={bold} I={italic}", new OfdFont(captionFace.FamilyName, 3, resourceId: "caption"), label, 12, y);
            graphics.DrawLine(new OfdPen(new OfdColor(170, 185, 200), .12), 12, y + 13, 135, y + 13);
            const string text = "AAAAAA";
            var starts = StringInfo.ParseCombiningCharacters(text); var advances = new float[starts.Length - 1];
            for (var i = 0; i < advances.Length; i++) advances[i] = font.MeasureText(text.Substring(starts[i], starts[i + 1] - starts[i]));
            OfdSkiaAdapter.Append(package, page, new[] { SkiaDrawEvent.Text(text, new(20, y + 13), font, paint, name, advances) });
            graphics.DrawString(text, new OfdFont(face.FamilyName, 8, resourceId: name), new OfdBrush(OfdColor.Black), 80, y + 13, advances.Select(v => (double)v).ToArray());
            canvas.DrawText(text, 20, y + 13, SKTextAlign.Left, font, paint); canvas.DrawText(text, 80, y + 13, SKTextAlign.Left, font, paint);
            var alias = PdfFontRegistry.RegisterFontFace(bytes, bold, italic);
            var resolver = GlobalFontSettings.FontResolver.ResolveTypeface(alias, bold, italic);
            if (resolver.MustSimulateBold || resolver.MustSimulateItalic) throw new Exception("Unexpected synthetic emphasis for " + name);
            results.Add(new { face = name, skia_bold = face.IsBold, skia_italic = face.IsItalic, resource_bold = bold, resource_italic = italic,
                text_weight = 400, text_italic = false, resolver.MustSimulateBold, resolver.MustSimulateItalic, sha256 = Hash(bytes), baseline_mm = y + 13 });
        }
        graphics.DrawString("Original MIT symbol outlines; no font fallback/subset service", new OfdFont(captionFace.FamilyName, 2.7, resourceId: "caption"), label, 12, 199);
        var source = Path.Combine(output, "font-styles.ofd"); await using (var stream = File.Create(source)) await new OfdPackageWriter().WriteAsync(package, stream);
        OfdDocumentPackage read; await using (var stream = File.OpenRead(source)) read = await new OfdReader().ReadAsync(stream);
        var symbolRuns = read.Pages[0].Elements.OfType<OfdTextElement>().Where(t => t.Text == "AAAAAA").ToArray();
        if (symbolRuns.Length != 10 || package.Pages[0].Elements.OfType<OfdTextElement>().Where(t => t.Text == "AAAAAA").Any(t => t.Weight != 400 || t.Italic)) throw new Exception("Font text emphasis was added or lost.");
        // Writer promotes actual resource flags into the XML; Reader reads those
        // effective fields. Check inheritance rather than an extra style request.
        foreach (var (name, bold, italic) in styles)
        {
            var resource = read.Fonts.Single(f => f.Id == name);
            var original = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "style-fonts", name + ".ttf"));
            if (resource.Data is null || Hash(resource.Data) != Hash(original) || resource.Bold != bold || resource.Italic != italic)
                throw new Exception("Roundtrip font payload or file flags changed: " + name);
        }
        foreach (var text in symbolRuns)
        {
            var xml = XElement.Parse(text.SourceXml!);
            var resource = read.Fonts.Single(f => f.Id == text.FontResourceId);
            var actualAlias = PdfFontRegistry.RegisterFontFace(resource.Data!, resource.Bold, resource.Italic);
            var actualResolver = GlobalFontSettings.FontResolver.ResolveTypeface(actualAlias, resource.Bold || text.Weight >= 600, resource.Italic || text.Italic);
            if (actualResolver.MustSimulateBold || actualResolver.MustSimulateItalic) throw new Exception("Roundtrip OFD requested synthetic emphasis.");
            var expectedWeight = resource.Bold ? "700" : "400";
            if ((xml.Attribute("Weight")?.Value ?? "400") != expectedWeight) throw new Exception("Written emphasis exceeds actual file flags.");
            if ((xml.Attribute("Italic")?.Value == "true") != resource.Italic) throw new Exception("Written italic exceeds actual file flags.");
        }
        await using (var input = File.OpenRead(source)) await using (var pdf = File.Create(Path.Combine(output, "font-styles.pdf"))) await new OfdToPdfConverter().ConvertAsync(input, pdf);
        using var png = bitmap.Encode(SKEncodedImageFormat.Png, 100); File.WriteAllBytes(Path.Combine(output, "font-styles-skia.png"), png.ToArray());
        File.WriteAllText(Path.Combine(output, "font-styles-report.json"), JsonSerializer.Serialize(new { status = "passed", pages = 1, symbol_runs = 10, exact_font_payload_binding = true, font_fallback = false, roundtrip_font_payloads_and_effective_style_requests_checked = true, styles = results }, new JsonSerializerOptions { WriteIndented = true }));
    }
    static string Hash(byte[] value) => Convert.ToHexString(SHA256.HashData(value)).ToLowerInvariant();
}
