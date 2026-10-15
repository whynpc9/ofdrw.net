using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class WriterPrecisionProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font);
        var normal = "0.9 0.94 0.98 rg 25 470 370 80 re f 0 0 0 rg " +
            fixture.Text("写出精度 / WRITER PRECISION", 35, 520, 16) +
            fixture.Text("A  B中文 - original Unicode and page policy", 35, 490, 11) +
            "0 0 1 RG 1 w 25 470 370 80 re S";
        // Local .0004 values are unrepresentable even when CTM amplification
        // makes the original line (2pt) / glyph (16pt) plainly visible.
        var stroke = normal + " 0.8 0 0 RG q 5000 0 0 5000 0 0 cm .0004 w .006 .06 m .06 .06 l S Q";
        var text = normal + $" 0 0.4 0 rg q 40000 0 0 40000 0 0 cm BT /F1 .0004 Tf 1 0 0 1 .001 .0075 Tm <{fixture.Hex("A")}> Tj ET Q";
        var options = new PdfVectorToOfdOptions { UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage,
            Compatibility = new PdfToOfdOptions { PreferExternalPdfToPpm = true } };
        foreach (var item in new[] { (stroke, "STROKE_WRITER_PRECISION"), (text, "TEXT_WRITER_PRECISION") })
        {
            using var input = new MemoryStream(fixture.Create(new[] { normal, item.Item1 }));
            using var sentinel = new MemoryStream(); sentinel.Write(new byte[] { 4, 5, 6 });
            try { await new PdfVectorToOfdConverter().ConvertAsync(input, sentinel); throw new Exception("Fail policy accepted an unrepresentable scalar."); }
            catch (NotSupportedException error) when (error.Message.Contains(item.Item2)) { }
            if (!sentinel.ToArray().SequenceEqual(new byte[] { 4, 5, 6 })) throw new Exception("Fail policy changed destination sentinel.");
        }
        var source = fixture.Create(new[] { normal, stroke, text, normal });
        File.WriteAllBytes(Path.Combine(output, "writer-precision-source.pdf"), source);
        using var mixedInput = new MemoryStream(source); using var destination = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(mixedInput, destination);
        if (!result.Pages.Select(page => page.IsNative).SequenceEqual(new[] { true, false, false, true }) ||
            !result.Pages[1].Diagnostic.Contains("STROKE_WRITER_PRECISION") || !result.Pages[2].Diagnostic.Contains("TEXT_WRITER_PRECISION"))
            throw new Exception("Writer scalar page-policy diagnostics changed.");
        destination.Position = 0; var package = await new OfdReader().ReadAsync(destination);
        foreach (var index in new[] { 1, 2 })
            if (package.Pages[index].Elements.OfType<OfdImageElement>().Count() != 1 ||
                package.Pages[index].Elements.OfType<OfdPathElement>().Any() ||
                package.Pages[index].Elements.OfType<OfdTextElement>().Any(value => value.FillColor.Alpha != 0))
                throw new Exception("Fallback leaked partial visible native content.");
        var controls = new List<object>();
        foreach (var isText in new[] { false, true })
        foreach (var size in new[] { ".000499999999", ".0005", ".00050001", ".001" })
        {
            var scale = size == ".001" ? ".1" : "1";
            var body = isText ? $"BT /F1 {size} Tf 1 0 0 1 30 100 Tm <{fixture.Hex("A")}> Tj ET" : $"{size} w 30 100 m 60 100 l S";
            var bytes = fixture.Create(new[] { $"q {scale} 0 0 {scale} 0 0 cm {body} Q" });
            foreach (var policy in new[] { PdfVectorUnsupportedPagePolicy.Fail, PdfVectorUnsupportedPagePolicy.RasterizePage })
            {
                using var input = new MemoryStream(bytes); using var actual = new MemoryStream();
                if (size != ".001" && policy == PdfVectorUnsupportedPagePolicy.Fail)
                {
                    actual.Write(new byte[] { 4, 5, 6 }); actual.Position = 0;
                    try { await new PdfVectorToOfdConverter().ConvertAsync(input, actual); throw new Exception("Inaccurate positive scalar accepted."); }
                    catch (NotSupportedException error) when (error.Message.Contains(isText ? "TEXT_WRITER_PRECISION" : "STROKE_WRITER_PRECISION"))
                    {
                        if (!actual.ToArray().SequenceEqual(new byte[] { 4, 5, 6 }) || actual.Position != 0) throw;
                        controls.Add(new { isText, size, scale, policy = policy.ToString(), sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(bytes)), exception = error.GetType().FullName, error.Message, sentinel = "040506", position = 0 });
                    }
                }
                else
                {
                    var control = await new PdfVectorToOfdConverter(new() { UnsupportedPagePolicy = policy, Compatibility = new() { PreferExternalPdfToPpm = true } }).ConvertWithResultAsync(input, actual);
                    var native = size == ".001";
                    if (control.Pages.Single().IsNative != native || control.Pages[0].ImageObjects != (native ? 0 : 1) || (!native && control.Pages[0].PathObjects != 0))
                        throw new Exception("Written scalar accuracy control changed policy.");
                    controls.Add(new { isText, size, scale, policy = policy.ToString(), sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(bytes)), native, result = control });
                }
            }
        }
        File.WriteAllBytes(Path.Combine(output, "writer-precision.ofd"), destination.ToArray());
        destination.Position = 0;
        await using (var pdf = File.Create(Path.Combine(output, "writer-precision.pdf")))
            await new OfdToPdfConverter().ConvertAsync(destination, pdf);
        File.WriteAllText(Path.Combine(output, "writer-precision-report.json"), JsonSerializer.Serialize(new
        {
            result, controls, failSentinelsUnchanged = true, fallbackHasNoVisibleNativeLeakage = true,
            sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)),
            ofdSha256 = Convert.ToHexStringLower(SHA256.HashData(destination.ToArray())),
            scope = "actual PackageReference consumer; Preview acceptance recorded separately"
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
