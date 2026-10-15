using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class WriterAccuracyProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var f = new PdfFixture(font);
        string Normal(string title) => "0.12 0.365 0.65 rg " + f.Text(title, 34, 535, 17) + "0.91 0.945 0.976 rg 0.12 0.365 0.65 RG 2 w 34 370 350 90 re B";
        var source = f.Create(new[] { Normal("正常原生 / NORMAL NATIVE"),
            "0.12 0.365 0.65 rg " + f.Text("线宽舍入 / STROKE ROUNDING", 34, 535, 17) +
                "q 100000 0 0 100000 0 0 cm 0.73 0.23 0.25 RG .0006 w .0006 .0035 m .0036 .0035 l S Q 0.91 0.945 0.976 rg 34 220 350 60 re f",
            "0.12 0.365 0.65 rg " + f.Text("字号舍入 / FONT ROUNDING", 34, 535, 17) +
                "q 100000 0 0 100000 0 0 cm 0.2 0.55 0.35 rg BT /F1 .0006 Tf 1 0 0 1 .0006 .0035 Tm <" + f.Hex("中文 A") +
                "> Tj ET Q 0.91 0.945 0.976 rg 34 220 350 60 re f", Normal("正常恢复 / NATIVE RECOVERY") });
        File.WriteAllBytes(Path.Combine(output, "writer-accuracy-source.pdf"), source);
        var failures = new List<object>();
        foreach (var box in new[] { "0 0 .001 100", "0 0 100 .001" })
        foreach (var policy in new[] { PdfVectorUnsupportedPagePolicy.Fail, PdfVectorUnsupportedPagePolicy.RasterizePage })
        {
            var tiny = f.Create(new[] { "0 0 .0005 10 re f" }, mediaBox: box);
            var tinyFile = box == "0 0 .001 100" ? "writer-tiny-width-source.pdf" : "writer-tiny-height-source.pdf";
            File.WriteAllBytes(Path.Combine(output, tinyFile), tiny);
            using var input = new MemoryStream(tiny); using var destination = Sentinel();
            try { await new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = policy }).ConvertAsync(input, destination); throw new Exception("Unrepresentable page succeeded."); }
            catch (NotSupportedException e)
            {
                Unchanged(destination); if (!e.Message.Contains("PAGE_WRITER_PRECISION")) throw;
                failures.Add(new { box, policy = policy.ToString(), inputFile = tinyFile, sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(tiny)), exception = e.GetType().FullName, e.Message, sentinel = "040506", position = 0 });
            }
        }
        using (var input = new MemoryStream(source)) using (var destination = Sentinel())
        {
            try { await new PdfVectorToOfdConverter().ConvertAsync(input, destination); throw new Exception("Expected stroke precision rejection."); }
            catch (NotSupportedException e)
            {
                Unchanged(destination); if (!e.Message.Contains("STROKE_WRITER_PRECISION")) throw;
                failures.Add(new { policy = "Fail", exception = e.GetType().FullName, e.Message, sentinel = "040506", position = 0 });
            }
        }
        using var sourceInput = new MemoryStream(source); using var ofd = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            Compatibility = new() { PreferExternalPdfToPpm = true } }).ConvertWithResultAsync(sourceInput, ofd);
        if (!result.Pages.Select(p => p.IsNative).SequenceEqual(new[] { true, false, false, true }) ||
            !result.Pages[1].Diagnostic.Contains("STROKE_WRITER_PRECISION") || !result.Pages[2].Diagnostic.Contains("TEXT_WRITER_PRECISION")) throw new Exception("Written scalar page policy changed.");
        ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        for (var index = 0; index < 4; index++)
        {
            var page = package.Pages[index];
            if (index is 1 or 2)
            {
                if (page.Elements.OfType<OfdImageElement>().Count() != 1 || page.Elements.OfType<OfdPathElement>().Any() ||
                    page.Elements.OfType<OfdTextElement>().Any(t => t.FillColor.Alpha != 0)) throw new Exception("Visible partial native content leaked into fallback.");
            }
            else if (page.Elements.OfType<OfdImageElement>().Any() || page.Elements.OfType<OfdPathElement>().Count() != 2 ||
                page.Elements.OfType<OfdTextElement>().Single().Text != (index == 0 ? "正常原生 / NORMAL NATIVE" : "正常恢复 / NATIVE RECOVERY")) throw new Exception("Normal native neighbor changed.");
        }
        var embedded = package.Fonts.Where(p => p.Data is { Length: > 0 }).ToArray();
        if (embedded.Length != 2 || embedded.Any(p => !p.Data.SequenceEqual(font)) ||
            new[] { 0, 3 }.Any(index => package.Pages[index].Elements.OfType<OfdTextElement>().Any(text => !embedded.Any(p => p.Id == text.FontResourceId))))
            throw new Exception("Native original font identity, reference or discarded-page accounting changed.");
        File.WriteAllBytes(Path.Combine(output, "writer-accuracy.ofd"), ofd.ToArray()); ofd.Position = 0;
        using (var pdf = File.Create(Path.Combine(output, "writer-accuracy.pdf"))) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        File.WriteAllText(Path.Combine(output, "writer-accuracy-report.json"), JsonSerializer.Serialize(new { sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)), failures, result, ofdBytes = ofd.Length,
            unrepresentablePage = "Both policies atomically reject; no output PDF exists for these failure cases." }, new JsonSerializerOptions { WriteIndented = true }));
    }
    private static MemoryStream Sentinel() { var output = new MemoryStream(); output.Write(new byte[] { 4, 5, 6 }); output.Position = 0; return output; }
    private static void Unchanged(MemoryStream output) { if (!output.ToArray().SequenceEqual(new byte[] { 4, 5, 6 }) || output.Position != 0) throw new Exception("Failed conversion changed destination."); }
}
