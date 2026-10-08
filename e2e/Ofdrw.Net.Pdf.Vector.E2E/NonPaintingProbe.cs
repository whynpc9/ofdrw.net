using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class NonPaintingProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font);
        var mixed = "20 20 m S 20 20 m h S 20 20 m 40 40 m f* 0.12 0.365 0.65 rg " +
            fixture.Text("No-op paths / 中文 A  B", 34, 535, 17) + "0.12 0.17 0.23 rg " +
            fixture.Text("Original text and actual paths remain native", 34, 508, 10) +
            "0.91 0.945 0.976 rg 34 330 350 120 re f 0.12 0.365 0.65 RG 2 w 34 330 350 120 re S 20 20 m s";
        var point = "0.12 0.365 0.65 rg " + fixture.Text("Closed point fill / 闭合点填充", 34, 535, 17) +
            "0.12 0.17 0.23 rg " + fixture.Text("Whole-page fallback, no visible native overlay", 34, 508, 10) +
            "0.91 0.945 0.976 rg 34 330 350 120 re f 1 0 0 rg 205 215 m h f";
        var singular = "0.12 0.365 0.65 rg " + fixture.Text("Singular CTM / 奇异矩阵", 34, 535, 17) +
            "0.12 0.17 0.23 rg " + fixture.Text("Text staged before the rejected transform survives raster fallback", 34, 508, 9) +
            "q 0 0 0 0 0 0 cm 20 20 m 40 40 l S Q";
        var rows = new List<object>();
        foreach (var item in new[] { ("no-op", mixed, true, "PDFV_NATIVE"), ("no-content", "20 20 m h S", false, "NO_NATIVE_CONTENT"),
            ("closed-point", point, false, "DEGENERATE_POINT_FILL"), ("singular", singular, false, "SINGULAR_SERIALIZED_MATRIX") })
        {
            var source = fixture.Create(new[] { item.Item2 }); File.WriteAllBytes(Path.Combine(output, item.Item1 + "-source.pdf"), source);
            if (!item.Item3)
            {
                using var failInput = new MemoryStream(fixture.Create(new[] { mixed, item.Item2 })); using var failOutput = new MemoryStream();
                failOutput.Write(new byte[] { 1, 2, 3 });
                try { await new PdfVectorToOfdConverter().ConvertAsync(failInput, failOutput); throw new Exception("Expected unsupported page."); }
                catch (NotSupportedException exception)
                {
                    if (!exception.Message.Contains(item.Item4) || !failOutput.ToArray().SequenceEqual(new byte[] { 1, 2, 3 })) throw;
                }
            }
            using var input = new MemoryStream(source); using var ofd = new MemoryStream();
            var result = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
                Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } }).ConvertWithResultAsync(input, ofd);
            var pageResult = result.Pages.Single();
            if (pageResult.IsNative != item.Item3 || !pageResult.Diagnostic.Contains(item.Item4) || pageResult.ImageObjects != (item.Item3 ? 0 : 1)) throw new Exception("Nonpainting page policy regression.");
            ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
            if (item.Item3 && !string.Concat(package.Pages[0].Elements.OfType<OfdTextElement>().Select(text => text.Text)).Contains("中文 A  B")) throw new Exception("Native original text lost.");
            if (!item.Item3 && (pageResult.PathObjects != 0 || package.Pages[0].Elements.OfType<OfdTextElement>().Any(text => text.FillColor.Alpha != 0))) throw new Exception("Fallback retained visible native paint.");
            File.WriteAllBytes(Path.Combine(output, item.Item1 + ".ofd"), ofd.ToArray()); ofd.Position = 0;
            using (var pdf = File.Create(Path.Combine(output, item.Item1 + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
            rows.Add(new { name = item.Item1, result, sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)), ofdBytes = ofd.Length });
        }
        // A scan plus no-op syntax must still be a real raster fallback, never native blank success.
        var scan = fixture.Create(new[] { "20 20 m S q 420 0 0 595 0 0 cm /Im1 Do Q", mixed });
        File.WriteAllBytes(Path.Combine(output, "no-op-scan-source.pdf"), scan);
        using var scanInput = new MemoryStream(scan); using var scanOutput = new MemoryStream();
        var scanResult = await new PdfVectorToOfdConverter(new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage }).ConvertWithResultAsync(scanInput, scanOutput);
        if (scanResult.Pages[0].IsNative || scanResult.Pages[0].ImageObjects != 1 || scanResult.Pages[0].PathObjects != 0 || !scanResult.Pages[1].IsNative) throw new Exception("No-op scan fallback regression.");
        File.WriteAllBytes(Path.Combine(output, "no-op-scan.ofd"), scanOutput.ToArray());
        rows.Add(new { name = "no-op-scan", result = scanResult, sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(scan)) });
        File.WriteAllText(Path.Combine(output, "nonpainting-report.json"), JsonSerializer.Serialize(rows, new JsonSerializerOptions { WriteIndented = true }));
    }
}
