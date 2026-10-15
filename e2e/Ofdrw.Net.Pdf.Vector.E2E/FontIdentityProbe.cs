using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

// Real PackageReference-only execution. Times are observations, not assertions;
// hash/read counters require separately labelled source instrumentation.
internal static class FontIdentityProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        var reports = new List<object>();
        foreach (var count in new[] { 1, 10, 50, 100 })
        foreach (var useArray in new[] { false, true })
        {
            using var fixture = new PdfFixture(font);
            var content = string.Concat(Enumerable.Range(0, count).Select(i =>
            {
                // Reset the text matrix, preserving independent painted objects
                // and an intervening path. No coalescing may hide repeated work.
                var text = fixture.Text(i % 2 == 0 ? "A" : "B", 20, 550 - i * 4);
                if (useArray) text = text.Replace($"<{fixture.Hex(i % 2 == 0 ? "A" : "B")}> Tj", $"[<{fixture.Hex(i % 2 == 0 ? "A" : "B")}>] TJ");
                return text + "0 0 0 RG 1 w 10 10 m 11 11 l S\n";
            }));
            var source = fixture.Create(new[] { content });
            using var input = new MemoryStream(source); using var destination = new MemoryStream();
            var allocatedBefore = GC.GetTotalAllocatedBytes(true); var timer = Stopwatch.StartNew();
            var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, destination);
            timer.Stop(); var allocated = GC.GetTotalAllocatedBytes(true) - allocatedBefore;
            destination.Position = 0; var package = await new OfdReader().ReadAsync(destination);
            var page = package.Pages.Single(); var resource = package.Fonts.Single();
            if (!result.Pages.Single().IsNative || page.Elements.Count != 2 * count || !font.SequenceEqual(resource.Data!))
                throw new Exception("Repeated-font package identity or element count changed.");
            for (var i = 0; i < count; i++)
            {
                if (page.Elements[2 * i] is not OfdTextElement text || text.Text != (i % 2 == 0 ? "A" : "B") ||
                    text.FontResourceId != resource.Id || page.Elements[2 * i + 1] is not OfdPathElement)
                    throw new Exception("Repeated-font original Unicode/resource/draw order changed.");
            }
            var name = "font-identity-" + (useArray ? "tj-array-" : "tj-") + count;
            File.WriteAllBytes(Path.Combine(output, name + "-source.pdf"), source);
            File.WriteAllBytes(Path.Combine(output, name + ".ofd"), destination.ToArray());
            reports.Add(new { name, count, elapsedMs = timer.Elapsed.TotalMilliseconds, allocatedBytes = allocated,
                sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)),
                ofdSha256 = Convert.ToHexStringLower(SHA256.HashData(destination.ToArray())),
                fontBytes = resource.Data!.Length, fontSha256 = Convert.ToHexStringLower(SHA256.HashData(resource.Data)),
                textObjects = count, pathObjects = count, originalUnicodeAndOrder = true,
                performanceCounters = "not instrumented; see separate source diagnostic and loaded package IL" });
        }
        File.WriteAllText(Path.Combine(output, "font-identity-report.json"), JsonSerializer.Serialize(reports, new JsonSerializerOptions { WriteIndented = true }));
    }
}
