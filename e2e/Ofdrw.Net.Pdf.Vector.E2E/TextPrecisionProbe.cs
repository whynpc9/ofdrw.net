using System.Security.Cryptography;
using System.IO.Compression;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class TextPrecisionProbe
{
    internal static async Task Run(string output)
    {
        // A frozen input must not be regenerated through platform-dependent font measurement.
        using var compressed = File.OpenRead(Path.Combine(AppContext.BaseDirectory, "testdata", "text-precision-source.pdf.gz"));
        using var gzip = new GZipStream(compressed, CompressionMode.Decompress);
        using var decodedSource = new MemoryStream(); await gzip.CopyToAsync(decodedSource);
        var source = decodedSource.ToArray();
        var sourceHash = Convert.ToHexStringLower(SHA256.HashData(source));
        if (source.Length != 21_698_835 || sourceHash != "d78e78154690b11a2e6c8db61d5ef458aef583e2706aa0da8ae065b3460898eb") throw new Exception("Frozen Astra text source changed.");
        var names = new[] { "matrix-translation", "advance-noncollapsed", "affine-amplification", "advance-accumulation", "font-size", "ordinary", "sheared", "cjk", "spacing" };
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
        if (embedded.Length != 4 || embedded.Any(f => Convert.ToHexStringLower(SHA256.HashData(f.Data)) != "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa")) throw new Exception("Original fixed-font payload identity changed.");
        File.WriteAllBytes(Path.Combine(output, "text-precision.ofd"), ofd.ToArray()); ofd.Position = 0;
        using (var pdf = File.Create(Path.Combine(output, "text-precision.pdf"))) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        File.WriteAllText(Path.Combine(output, "text-precision-report.json"), JsonSerializer.Serialize(new { sourceSha256 = sourceHash, sourceBytes = source.Length, result, cases = names.Select((name, i) => new { page = i + 1, Name = name }), originalEmbeddedFontResources = embedded.Length, preview = "separate actual page inspection required" }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
