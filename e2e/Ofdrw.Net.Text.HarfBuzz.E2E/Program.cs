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
const string ascii = "AV To\nWide WWW\nNarrow iii";
var layout = engine.Layout(new TextPrelayoutRequest(ascii, new[] { noto }, 5, 55));
var han = engine.Layout(new TextPrelayoutRequest("中文", new[] { noto, wenkai }, 5, 55,
    script: "Hani", language: "zh"));
var package = TextPrelayoutOfdWriter.CreateSinglePage(layout, 100, 75, 10, 10);
var ofd = Path.Combine(output, "r21-ascii.ofd");
var pdf = Path.Combine(output, "r21-ascii.pdf");
var png = Path.Combine(output, "r21-ascii.png");
await using (var stream = File.Create(ofd)) await new OfdPackageWriter().WriteAsync(package, stream);
await using (var source = File.OpenRead(ofd)) await using (var target = File.Create(pdf))
    await new OfdToPdfConverter().ConvertAsync(source, target);
await using (var source = File.OpenRead(ofd)) await using (var target = File.Create(png))
    await new OfdToImageConverter().ConvertAsync(source, target);
await using var readStream = File.OpenRead(ofd);
var read = await new OfdReader().ReadAsync(readStream);
var actual = string.Join("\n", read.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
    .Select(element => element.Text));
if (actual != ascii) throw new InvalidDataException("Native OFD text changed during save/readback.");
var receipt = new
{
    sourceSha = Environment.GetEnvironmentVariable("R21_SOURCE_SHA"),
    engine = layout.Lines[0].ShapedRun.EngineVersion,
    font = new { noto.Name, noto.SourceSha256, noto.FaceSha256, noto.UnitsPerEm },
    ascii, ofdText = actual,
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
    han = new { face = han.Face.FaceSha256, fallback = han.Diagnostics.Select(d => d.Code).ToArray(),
        advanceMillimeters = han.WidthMillimeters },
    outputs = new[] { ofd, pdf, png }.Select(path => new
    {
        name = Path.GetFileName(path), bytes = new FileInfo(path).Length,
        sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant(),
    }).ToArray(),
};
await File.WriteAllTextAsync(Path.Combine(output, "r21-receipt.json"),
    JsonSerializer.Serialize(receipt, new JsonSerializerOptions { WriteIndented = true }) + "\n");
Console.WriteLine("R21 fixed-font OFD/PDF/PNG and query receipt: " + output);
