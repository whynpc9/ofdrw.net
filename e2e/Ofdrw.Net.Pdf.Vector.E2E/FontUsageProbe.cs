using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class FontUsageProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font);
        const string aliases = "/F1 3 0 R /F2 3 0 R";
        const string shapes = "0.9 0.94 0.98 rg 25 470 370 80 re f 0 0 1 RG 1 w 25 470 370 80 re S " +
            "0.8 0 0 RG 2 w 30 300 m 300 300 l S 0 0.4 0 rg 50 180 70 40 re f ";
        var empty = shapes + "BT /F1 12 Tf <> Tj /F2 24 Tf [] TJ ET";
        var numeric = shapes + "BT /F1 12 Tf 1 0 0 1 60 90 Tm [100 <> -100 () 25] TJ /F2 24 Tf ET";
        var normal = "0.9 0.94 0.98 rg 25 470 370 80 re f 0 0 1 RG 1 w 25 470 370 80 re S 0 0 0 rg " +
            "BT /F1 24 Tf <> Tj ET " + fixture.Text("按需字体 / FONT USAGE", 35, 520, 16).Replace("/F1", "/F2") +
            fixture.Text("A  B中文 - exact original font", 35, 490, 11).Replace("/F1", "/F2") +
            "0 0.4 0 rg 30 380 360 35 re f 0 0 0 rg " + fixture.Text("表格 Table / 128.00", 40, 390, 12).Replace("/F1", "/F2");
        var a = fixture.Hex("A");
        var restored = normal + $" 0.8 0 0 rg BT /F1 12 Tf 1 0 0 1 60 300 Tm q /F2 24 Tf [100 <> -50 () 25] TJ Q <{a}> Tj ET";
        // Only p3 F2 and p4 F2/F1 produce text. p1/p2 contribute no fonts.
        var source = fixture.Create(new[] { empty, numeric, normal, restored }, fontResources: aliases);
        File.WriteAllBytes(Path.Combine(output, "font-usage-source.pdf"), source);
        using var input = new MemoryStream(source); using var destination = new MemoryStream();
        var options = new PdfVectorToOfdOptions { MaxTotalFontBytes = font.Length * 3L,
            MaxFontReadBytesPerPage = font.Length * 2L };
        var result = await new PdfVectorToOfdConverter(options).ConvertWithResultAsync(input, destination);
        if (result.Pages.Any(page => !page.IsNative || page.ImageObjects != 0)) throw new Exception("Font usage fixture left native profile.");
        destination.Position = 0; var package = await new OfdReader().ReadAsync(destination);
        if (package.Fonts.Count != 3 || package.Fonts.Sum(f => (long)f.Data.Length) != font.Length * 3L ||
            package.Fonts.Any(f => !f.Data.SequenceEqual(font)) ||
            package.Pages.Take(2).Any(page => page.Elements.OfType<OfdTextElement>().Any()) ||
            package.Pages.Skip(2).Any(page => !string.Concat(page.Elements.OfType<OfdTextElement>().Select(text => text.Text)).Contains("A  B中文")))
            throw new Exception("Unused font output or original text/font identity changed.");
        // Duplicated nonpainting selections still have no output font charge.
        using var selectedInput = new MemoryStream(source); using var selectedOutput = new MemoryStream();
        await new PdfVectorToOfdConverter(new() { MaxTotalFontBytes = font.Length }).ConvertAsync(selectedInput, selectedOutput, new[] { 0, 0, 1, 2 });
        using var usedInput = new MemoryStream(source); using var sentinel = new MemoryStream(); sentinel.Write(new byte[] { 7, 8, 9 });
        try { await new PdfVectorToOfdConverter(new() { MaxTotalFontBytes = font.Length }).ConvertAsync(usedInput, sentinel, new[] { 2, 2 }); throw new Exception("Used output font budget was ignored."); }
        catch (InvalidDataException) { }
        if (!sentinel.ToArray().SequenceEqual(new byte[] { 7, 8, 9 })) throw new Exception("Font output budget changed destination sentinel.");
        File.WriteAllBytes(Path.Combine(output, "font-usage.ofd"), destination.ToArray()); destination.Position = 0;
        await using (var pdf = File.Create(Path.Combine(output, "font-usage.pdf"))) await new OfdToPdfConverter().ConvertAsync(destination, pdf);
        File.WriteAllText(Path.Combine(output, "font-usage-report.json"), JsonSerializer.Serialize(new {
            result, fonts = package.Fonts.Select(f => new { f.Id, bytes = f.Data.Length, sha256 = Convert.ToHexStringLower(SHA256.HashData(f.Data)) }),
            sourceSha256 = Convert.ToHexStringLower(SHA256.HashData(source)), unusedSelectionsAccepted = true, usedBudgetSentinelUnchanged = true,
            scope = "actual PackageReference consumer; Preview remains separate"
        }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
