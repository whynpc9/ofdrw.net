using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class BudgetOrderingProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font);
        var first = "0.12 0.365 0.65 rg " + fixture.Text("事件预算 / EVENT BUDGET", 34, 535, 17) +
            "0.91 0.945 0.976 rg 34 400 350 60 re f " +
            "0.73 0.23 0.25 RG 12 w q 1 0 0 1 -16777156 0 cm 16777216 350 m 16777217 350 l S Q";
        var second = "0.12 0.365 0.65 rg " + fixture.Text("正常原生 / NORMAL NATIVE", 34, 535, 17) +
            "0.91 0.945 0.976 rg 0.12 0.365 0.65 RG 2 w 34 370 350 90 re B " +
            "20 20 m h S BT /F1 12 Tf () Tj ET";
        var source = fixture.Create(new[] { first, second });
        File.WriteAllBytes(Path.Combine(output, "budget-ordering-source.pdf"), source);
        var reports = new List<object>();
        foreach (var policy in new[] { PdfVectorUnsupportedPagePolicy.Fail, PdfVectorUnsupportedPagePolicy.RasterizePage })
        {
            using var input = new MemoryStream(source); using var destination = new MemoryStream();
            destination.Write(new byte[] { 4, 5, 6 }); destination.Position = 0;
            try
            {
                await new PdfVectorToOfdConverter(new() { MaxEventsPerPage = 2, UnsupportedPagePolicy = policy }).ConvertAsync(input, destination);
                throw new Exception("Exhausted event capacity unexpectedly chose page fallback.");
            }
            catch (InvalidDataException error)
            {
                if (!error.Message.Contains("draw event limit") || !destination.ToArray().SequenceEqual(new byte[] { 4, 5, 6 }) || destination.Position != 0)
                    throw new Exception("Event budget hard failure modified the output or changed failure category.", error);
                reports.Add(new { policy = policy.ToString(), maxEvents = 2, exception = error.GetType().FullName, error.Message,
                    sentinel = Convert.ToHexString(destination.ToArray()), position = destination.Position });
            }
        }
        using (var input = new MemoryStream(source))
        using (var destination = new MemoryStream())
        {
            destination.Write(new byte[] { 4, 5, 6 }); destination.Position = 0;
            try { await new PdfVectorToOfdConverter(new() { MaxEventsPerPage = 3 }).ConvertAsync(input, destination); throw new Exception("Expected precision-profile rejection."); }
            catch (NotSupportedException error)
            {
                if (!error.Message.Contains("PATH_FLOAT_PRECISION") || !destination.ToArray().SequenceEqual(new byte[] { 4, 5, 6 }) || destination.Position != 0)
                    throw new Exception("Enough-capacity Fail changed its precision diagnostic or destination.", error);
                reports.Add(new { policy = "Fail", maxEvents = 3, exception = error.GetType().FullName, error.Message,
                    sentinel = Convert.ToHexString(destination.ToArray()), position = destination.Position });
            }
        }
        using var rasterInput = new MemoryStream(source); using var ofd = new MemoryStream();
        var result = await new PdfVectorToOfdConverter(new() { MaxEventsPerPage = 3,
            UnsupportedPagePolicy = PdfVectorUnsupportedPagePolicy.RasterizePage, Compatibility = new() { PreferExternalPdfToPpm = true } })
            .ConvertWithResultAsync(rasterInput, ofd);
        if (!result.Pages.Select(page => page.IsNative).SequenceEqual(new[] { false, true }) ||
            !result.Pages[0].Diagnostic.Contains("PATH_FLOAT_PRECISION") || result.Pages[0].ImageObjects != 1 || result.Pages[0].PathObjects != 0 ||
            result.Pages[1].PathObjects != 2 || result.Pages[1].TextObjects != 1)
            throw new Exception("Enough-capacity fallback/native event boundary changed.");
        ofd.Position = 0; var package = await new OfdReader().ReadAsync(ofd);
        if (package.Pages[0].Elements.OfType<OfdPathElement>().Any() ||
            package.Pages[0].Elements.OfType<OfdImageElement>().Count() != 1 ||
            package.Pages[0].Elements.OfType<OfdTextElement>().Any(text => text.FillColor.Alpha != 0))
            throw new Exception("Partial visible native content leaked into whole-page fallback.");
        var native = package.Pages[1]; var paths = native.Elements.OfType<OfdPathElement>().ToArray();
        if (native.Elements.OfType<OfdTextElement>().Single().Text != "正常原生 / NORMAL NATIVE" ||
            paths.Length != 2 || !paths[0].Fill || paths[0].Stroke || paths[1].Fill || !paths[1].Stroke ||
            native.Elements.OfType<OfdImageElement>().Any())
            throw new Exception("Native control text, fill/stroke order or zero-charge trailing no-op changed.");
        File.WriteAllBytes(Path.Combine(output, "budget-ordering.ofd"), ofd.ToArray()); ofd.Position = 0;
        using (var pdf = File.Create(Path.Combine(output, "budget-ordering.pdf"))) await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        reports.Add(new { policy = "RasterizePage", maxEvents = 3, result, ofdBytes = ofd.Length });
        File.WriteAllText(Path.Combine(output, "budget-ordering-report.json"), JsonSerializer.Serialize(new
        { sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)), reports }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
