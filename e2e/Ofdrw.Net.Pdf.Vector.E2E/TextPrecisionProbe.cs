using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class TextPrecisionProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font); var z = fixture.Hex("中");
        string Pair(string matrix, string size, string adjustment) => $"BT /F1 {size} Tf {matrix} Tm [<{z}> {adjustment} <{z}>] TJ ET\n";
        var repeated = new StringBuilder($"BT /F1 16 Tf 1 0 0 1 60 350 Tm [<{z}>");
        for (var i = 0; i < 1024; i++) repeated.Append($" -5250.0001875 <{z}> 7250 <{z}>");
        repeated.Append("] TJ ET\n");
        // Byte-identical source to Astra's real r4 reproduction; it is never corrected to hide the loss.
        var cases = new (string Name, string Content)[] {
            ("matrix-translation", Pair("1 0 0 1 16777217 350", "16", "1048573312.5")),
            ("advance-noncollapsed", Pair("1 0 0 1 100000000 350", "16", "6249997250")),
            ("affine-amplification", Pair("128 0 0 128 16777216 350", ".125", "1048573234.375")),
            ("advance-accumulation", repeated.ToString()),
            ("font-size", fixture.Text("中", 60, 350, 8192.0003)),
            ("ordinary", fixture.Text("Ordinary A B 123", 34.1, 535.2, 15.3)),
            ("sheared", fixture.Text("Sheared A B 中文", 34, 480, 15, .2)),
            ("cjk", fixture.Text("中文字体 原始文本", 34, 430, 16)),
            ("spacing", "BT /F1 16 Tf 1 0 0 1 34 380 Tm 2 Tc 4 Tw [<" + fixture.Hex("A  B中文") + "> 125 <" + fixture.Hex(" C") + ">] TJ ET")
        };
        var source = fixture.Create(cases.Select(c => c.Content).ToArray());
        var sourceHash = Convert.ToHexStringLower(SHA256.HashData(source));
        if (sourceHash != "d78e78154690b11a2e6c8db61d5ef458aef583e2706aa0da8ae065b3460898eb") throw new Exception("Frozen Astra text source changed.");
        File.WriteAllBytes(Path.Combine(output, "text-precision-source.pdf"), source);
        using (var input = new MemoryStream(source)) using (var sentinel = new MemoryStream())
        {
            sentinel.Write(new byte[] { 7, 8 });
            try { await new PdfVectorToOfdConverter().ConvertAsync(input, sentinel); throw new Exception("Text loss wrongly accepted."); }
            catch (NotSupportedException e) { if (!e.Message.Contains("TEXT_FLOAT_PRECISION") || !sentinel.ToArray().SequenceEqual(new byte[] { 7, 8 })) throw new Exception("Text precision Fail contract failed."); }
        }
        using var sourceInput = new MemoryStream(source); using var ofd = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage, Compatibility = new() { PreferExternalPdfToPpm = true } }).ConvertWithResultAsync(sourceInput, ofd);
        var expected = new[] { false, false, false, false, false, true, true, true, true };
        if (!result.Pages.Select(p => p.IsNative).SequenceEqual(expected)) throw new Exception("Text precision disposition changed.");
        ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        for (var i = 0; i < 5; i++)
        {
            if (result.Pages[i].ImageObjects != 1 || result.Pages[i].PathObjects != 0 || !result.Pages[i].Diagnostic.Contains("TEXT_FLOAT_PRECISION") || package.Pages[i].Elements.OfType<OfdTextElement>().Any(t => t.FillColor.Alpha != 0)) throw new Exception("Text fallback visible native leak.");
        }
        for (var i = 5; i < 9; i++) if (result.Pages[i].ImageObjects != 0 || package.Pages[i].Elements.OfType<OfdTextElement>().Count() != 1) throw new Exception("Ordinary text lost native content.");
        if (package.Pages[8].Elements.OfType<OfdTextElement>().Single().Text != "A  B中文 C") throw new Exception("Original Unicode/spacing changed.");
        var embedded = package.Fonts.Where(f => f.Data.Length > 0).ToArray();
        if (embedded.Length != 4 || embedded.Any(f => !f.Data.SequenceEqual(font))) throw new Exception("Original fixed-font payload identity changed.");
        File.WriteAllBytes(Path.Combine(output, "text-precision.ofd"), ofd.ToArray()); ofd.Position = 0;
        using (var pdf = File.Create(Path.Combine(output, "text-precision.pdf"))) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        File.WriteAllText(Path.Combine(output, "text-precision-report.json"), JsonSerializer.Serialize(new { sourceSha256 = sourceHash, result, cases = cases.Select((c, i) => new { page = i + 1, c.Name }), originalEmbeddedFontResources = embedded.Length, preview = "separate actual page inspection required" }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
