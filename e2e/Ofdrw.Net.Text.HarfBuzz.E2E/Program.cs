using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

if (args.Length != 1) throw new ArgumentException("Usage: <output-directory>");
var root = Directory.GetCurrentDirectory();
var output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var fixture = Path.Combine(root, "e2e", "Ofdrw.Net.FontSubset.E2E", "testdata", "fonts");
var noto = new TextFontFace("Noto Sans", File.ReadAllBytes(Path.Combine(fixture, "NotoSans-Regular.ttf")));
var wenkai = new TextFontFace("LXGW WenKai", File.ReadAllBytes(Path.Combine(fixture, "LXGWWenKai-Regular.ttf")));
var engine = HarfBuzzTextPrelayout.Create();
const string ascii = "JAV To\nWide WWW\nNarrow iii f";
const string hanText = "中文中文";
var layout = engine.Layout(new TextPrelayoutRequest(ascii, new[] { noto }, 5, 55));
var han = engine.Layout(new TextPrelayoutRequest(hanText, new[] { noto, wenkai }, 5, 11,
    script: "Hani", language: "zh"));
if (han.Overflows || han.Lines.Count != 2 || !ReferenceEquals(han.Face, wenkai))
    throw new InvalidDataException("The fixed Han sample did not select the expected fallback and line geometry.");

// This E2E probe is limited to one-glyph-per-BMP-Han output. The public
// typed writer intentionally continues to reject Han until the GID-aware stage.
OfdDocumentPackage CreateHanVisualProbe()
{
    var package = new OfdDocumentPackage();
    package.Options.DefaultPageWidthMillimeters = 100;
    package.Options.DefaultPageHeightMillimeters = 75;
    package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
    package.Fonts.Add(new OfdFontResource { Id = "F1", FontName = han.Face.Name,
        FamilyName = han.Face.Name, Charset = "unicode", Data = han.Face.GetFaceBytes() });
    var page = new OfdPage { Id = "1", Index = 0, WidthMillimeters = 100, HeightMillimeters = 75 };
    var scale = han.Request.FontSizeMillimeters / han.Face.UnitsPerEm;
    foreach (var line in han.Lines)
    {
        var text = han.Request.Text.Substring(line.SourceStart, line.SourceLength);
        if (line.ShapedRun.Glyphs.Count != text.Length || line.ShapedRun.Request.Text != text ||
            !ReferenceEquals(line.ShapedRun.Request.Face, han.Face))
            throw new InvalidDataException("Han visual probe requires one shaped glyph per source character.");
        for (var i = 0; i < text.Length; i++)
        {
            var glyph = line.ShapedRun.Glyphs[i];
            if (char.IsSurrogate(text[i]) || glyph.ClusterStart != line.SourceStart + i ||
                glyph.ClusterLength != 1 || glyph.GlyphId == 0 || !han.Face.Covers(text[i]) ||
                glyph.XOffset != 0 || glyph.YOffset != 0 || glyph.YAdvance != 0)
                throw new InvalidDataException("Han visual probe cannot represent the shaped glyph losslessly.");
        }
        var ink = line.InkBoundsMillimeters;
        var inkTop = ink?.Y - line.TopMillimeters ?? 0;
        var boundaryLeft = Math.Min(0, ink?.X ?? 0);
        var boundaryTop = Math.Min(0, inkTop);
        var boundaryRight = Math.Max(line.AdvanceMillimeters, ink is null ? 0 : ink.X + ink.Width);
        var boundaryBottom = Math.Max(line.HeightMillimeters, ink is null ? 0 : inkTop + ink.Height);
        if (10 + boundaryLeft < 0 || 10 + line.TopMillimeters + boundaryTop < 0 ||
            10 + boundaryRight > 100 || 10 + line.TopMillimeters + boundaryBottom > 75)
            throw new InvalidDataException("Han visual probe exceeds its destination page.");
        var element = new OfdTextElement
        {
            LayerType = "Body", XMillimeters = 10 + boundaryLeft,
            YMillimeters = 10 + line.TopMillimeters + boundaryTop,
            WidthMillimeters = Math.Max(0.1, boundaryRight - boundaryLeft),
            HeightMillimeters = Math.Max(0.1, boundaryBottom - boundaryTop),
            FontName = han.Face.Name, FontResourceId = "F1",
            FontSizeMillimeters = han.Request.FontSizeMillimeters, Text = text,
        };
        element.Runs.Add(new OfdTextRun
        {
            Text = text, XMillimeters = -boundaryLeft,
            YMillimeters = line.BaselineMillimeters - line.TopMillimeters - boundaryTop,
            DeltaX = text.Length > 1 ? string.Join(" ", line.ShapedRun.Glyphs.Take(text.Length - 1)
                .Select(glyph => (glyph.XAdvance * scale).ToString("0.###############", CultureInfo.InvariantCulture))) : null,
        });
        page.Elements.Add(element);
    }
    package.Pages.Add(page);
    return package;
}

async Task<(string Text, string[] Files)> RenderAndRead(string name, TextPrelayoutResult result, OfdDocumentPackage package)
{
    var ofd = Path.Combine(output, "r21-" + name + ".ofd");
    var pdf = Path.Combine(output, "r21-" + name + ".pdf");
    var png = Path.Combine(output, "r21-" + name + ".png");
    await using (var stream = File.Create(ofd)) await new OfdPackageWriter().WriteAsync(package, stream);
    await using (var source = File.OpenRead(ofd)) await using (var target = File.Create(pdf))
        await new OfdToPdfConverter().ConvertAsync(source, target);
    await using (var source = File.OpenRead(ofd)) await using (var target = File.Create(png))
        await new OfdToImageConverter().ConvertAsync(source, target);
    await using var readStream = File.OpenRead(ofd);
    var read = await new OfdReader().ReadAsync(readStream);
    var actualLines = read.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
        .Select(element => element.Text).ToArray();
    var expectedLines = result.Lines.Where(line => line.SourceLength > 0)
        .Select(line => result.Request.Text.Substring(line.SourceStart, line.SourceLength)).ToArray();
    if (!actualLines.SequenceEqual(expectedLines))
        throw new InvalidDataException(name + " OFD line text changed during save/readback.");
    var actual = string.Join("\n", actualLines);
    return (actual, new[] { ofd, pdf, png });
}
var asciiOutput = await RenderAndRead("ascii", layout, TextPrelayoutOfdWriter.CreateSinglePage(layout, 100, 75, 10, 10));
var hanOutput = await RenderAndRead("han", han, CreateHanVisualProbe());
if (asciiOutput.Text != ascii) throw new InvalidDataException("ASCII OFD hard line text changed during save/readback.");
var receipt = new
{
    sourceSha = Environment.GetEnvironmentVariable("R21_SOURCE_SHA"),
    engine = layout.Lines[0].ShapedRun.EngineVersion,
    font = new { noto.Name, noto.SourceSha256, noto.FaceSha256, noto.UnitsPerEm },
    ascii, ofdText = asciiOutput.Text,
    lines = layout.Lines.Select(line => new
    {
        line.SourceStart, line.SourceLength, line.TerminatorLength, line.AdvanceMillimeters,
        line.TopMillimeters, line.BaselineMillimeters, line.HeightMillimeters,
        ink = line.InkBoundsMillimeters is null ? null : new
        {
            line.InkBoundsMillimeters.X, line.InkBoundsMillimeters.Y,
            line.InkBoundsMillimeters.Width, line.InkBoundsMillimeters.Height,
        },
        gids = line.ShapedRun.Glyphs.Select(glyph => glyph.GlyphId).ToArray(),
    }).ToArray(),
    han = new { source = hanText, face = han.Face.FaceSha256, fallback = han.Diagnostics.Select(d => d.Code).ToArray(),
        advanceMillimeters = han.WidthMillimeters, ofdText = hanOutput.Text,
        lines = han.Lines.Select(line => new { line.SourceStart, line.SourceLength,
            line.AdvanceMillimeters, line.InkBoundsMillimeters }).ToArray() },
    outputs = asciiOutput.Files.Concat(hanOutput.Files).Select(path => new
    {
        name = Path.GetFileName(path), bytes = new FileInfo(path).Length,
        sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant(),
    }).ToArray(),
};
await File.WriteAllTextAsync(Path.Combine(output, "r21-receipt.json"),
    JsonSerializer.Serialize(receipt, new JsonSerializerOptions { WriteIndented = true }) + "\n");
Console.WriteLine("R21 ASCII and Han fixed-font OFD/PDF/PNG and query receipt: " + output);
