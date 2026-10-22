using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

internal static class R06ProfileProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font);
        var frame = "0.92 0.95 0.99 rg 50 70 320 450 re f 0.10 0.31 0.59 rg " +
            fixture.Text("旋转裁切 / ROTATED CROP", 65, 485, 18) +
            "0.75 0.20 0.25 rg 65 85 85 55 re f 0.10 0.31 0.59 rg " +
            fixture.Text("Top-left after 90 degree rotation", 70, 430, 11);
        await Convert("r06-frame", fixture.Create(new[] { frame }, pageEntries: "/CropBox [40 60 380 540] /Rotate 90"),
            output, expectedPaths: 2, expectedImages: 0, expectedText: "旋转裁切 / ROTATED CROP");

        var mixed = "0.94 0.96 0.99 rg 20 30 380 530 re f 0.10 0.20 0.30 rg " +
            fixture.Text("裁切 + 图像 + 画笔 / CLIP IMAGE STROKE", 35, 515, 15) +
            "q 40 250 220 180 re W n 0.16 0.51 0.79 rg 10 240 300 200 re f " +
            "q 80 0 0 70 230 325 cm /Im1 Do Q " +
            "0.78 0.19 0.22 RG 3 w 1 J 1 j [8 4] 1 d 45 270 m 275 410 l S 1 1 1 rg " +
            fixture.Text("A  B中文 / native image", 48, 285, 13) + "Q 0.10 0.20 0.30 rg " +
            fixture.Text("Outside clip remains visible", 35, 100, 12);
        await Convert("r06-mixed", fixture.Create(new[] { mixed }), output,
            expectedPaths: 3, expectedImages: 1, expectedText: "A  B中文 / native image");

        var cid = "0.11 0.36 0.62 rg " + fixture.Text("CID mapping stream: A  B中文", 35, 500, 15) +
            "0.18 0.56 0.41 rg 40 365 210 80 re f";
        var map = Enumerable.Range(0, 65536).SelectMany(value => new[] { (byte)(value >> 8), (byte)value }).ToArray();
        await Convert("r06-cid-map", fixture.Create(new[] { cid }, cidMapStream: map), output,
            expectedPaths: 1, expectedImages: 0, expectedText: "CID mapping stream: A  B中文");

        using var winAnsi = new PdfFixture(font);
        var latin = "0.12 0.36 0.62 rg " + winAnsi.TextWinAnsi("WinAnsi native: A  Bé", 35, 500, 16) +
            "0.73 0.23 0.25 rg 40 365 210 80 re f";
        await Convert("r06-winansi", winAnsi.Create(new[] { latin }, simpleWinAnsi: true), output,
            expectedPaths: 1, expectedImages: 0, expectedText: "WinAnsi native: A  Bé");
    }

    private static async Task Convert(string name, byte[] source, string output, int expectedPaths, int expectedImages, string expectedText)
    {
        var sourcePath = Path.Combine(output, name + "-source.pdf");
        File.WriteAllBytes(sourcePath, source);
        using var input = new MemoryStream(source);
        using var ofd = new MemoryStream();
        var result = await new PdfVectorToOfdConverter().ConvertWithResultAsync(input, ofd);
        if (!result.Pages.Single().IsNative || result.Pages[0].PathObjects != expectedPaths || result.Pages[0].ImageObjects != expectedImages)
            throw new InvalidDataException(name + " did not remain in the expected native profile.");
        var ofdBytes = ofd.ToArray();
        File.WriteAllBytes(Path.Combine(output, name + ".ofd"), ofdBytes);
        ofd.Position = 0;
        var page = (await new OfdReader().ReadAsync(ofd)).Pages.Single();
        if (!string.Concat(page.Elements.OfType<OfdTextElement>().Select(text => text.Text)).Contains(expectedText, StringComparison.Ordinal))
            throw new InvalidDataException(name + " lost original Unicode text.");
        ofd.Position = 0;
        await using (var pdf = File.Create(Path.Combine(output, name + ".pdf")))
            await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        ofd.Position = 0;
        await using (var svg = File.Create(Path.Combine(output, name + ".svg")))
            await new OfdToSvgConverter().ConvertAsync(ofd, svg, 0);
        var report = new { name, native = true, paths = result.Pages[0].PathObjects, texts = result.Pages[0].TextObjects,
            images = result.Pages[0].ImageObjects, sourceSha256 = System.Convert.ToHexStringLower(SHA256.HashData(source)),
            ofdSha256 = System.Convert.ToHexStringLower(SHA256.HashData(ofdBytes)), text = expectedText };
        File.WriteAllText(Path.Combine(output, name + "-report.json"), JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
    }
}
