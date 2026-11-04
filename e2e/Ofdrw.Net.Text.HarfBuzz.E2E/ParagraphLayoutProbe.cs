using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Layout.Text.Paragraphs;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

// Issue #109 paragraph layout through public APIs only: layout → OFD → read back → PDF/SVG.
internal static class ParagraphLayoutProbe
{
    internal const string Sample =
        "Paragraph layout 段落排版：中文与 English 混排，Unicode 17.0.0 双向算法、断行与字素簇。" +
        "The quick brown fox jumps over the lazy dog; 敏捷的棕色狐狸跳过了懒狗。\n" +
        "Hebrew in an LTR line: שלום עולם (ver. 2.0) and back to Latin.\n" +
        "مرحبا بالعالم، هذا نص عربي يلتف عبر عدة أسطر مع أرقام 123 وكلمة English داخل السطر.\n" +
        "தமிழ் எழுத்துகள் கூட்டெழுத்துகளுடன் சரியாக வடிவமைக்கப்படுகின்றன க்ஷ ஸ்ரீ.\n" +
        "Line one\u2028line two after a line separator.";

    internal static async Task Run(string output, string fixture, TextFontFace noto, TextFontFace wenkai,
        TextFontFace dejavu)
    {
        var tamil = new TextFontFace("Noto Sans Tamil", File.ReadAllBytes(Path.Combine(fixture, "NotoSansTamil-Regular.ttf")));
        TextStyle Style(double size, TextPaint? paint = null) => new(new[]
        {
            new TextFontFamily("Noto Sans", noto), new TextFontFamily("LXGW WenKai", wenkai),
            new TextFontFamily("DejaVu Sans", dejavu), new TextFontFamily("Noto Sans Tamil", tamil),
        }, size, paint);
        var cases = new (string Name, ParagraphLayoutOptions Options)[]
        {
            ("r22-paragraph-justify", new ParagraphLayoutOptions(90, 60, alignment: TextAlignment.Justify,
                lineHeight: TextLineHeight.Multiple(1.2), paragraphSpacingMillimeters: 2, firstLineIndentMillimeters: 6)),
            ("r22-paragraph-rtl-center", new ParagraphLayoutOptions(90, null, TextParagraphDirection.RightToLeft,
                TextAlignment.Center)),
        };
        var receipts = new List<object>();
        foreach (var (name, options) in cases)
        {
            var english = Sample.IndexOf("English", StringComparison.Ordinal);
            var spans = new[]
            {
                new TextStyleSpan(0, english, Style(4.2)),
                new TextStyleSpan(english, 7, Style(5, new TextPaint(new TextColor(180, 30, 30)))),
                new TextStyleSpan(english + 7, Sample.Length - english - 7, Style(4.2)),
            };
            var layout = HarfBuzzTextPrelayout.CreateParagraphLayout().Layout(new ParagraphLayoutRequest(Sample, spans, options));
            var package = ParagraphLayoutOfdWriter.CreateDocument(layout, 110, options.PageHeightMillimeters is null ? 160 : 80, 10, 10);
            var ofd = Path.Combine(output, name + ".ofd");
            await using (var stream = File.Create(ofd)) await new OfdPackageWriter().WriteAsync(package, stream);

            // Read back: every run is one TextObject with its logical text, in order.
            OfdDocumentPackage read;
            await using (var stream = File.OpenRead(ofd)) read = await new OfdReader().ReadAsync(stream);
            var texts = read.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>()).Select(text => text.Text).ToArray();
            var runs = layout.Lines.SelectMany(line => line.Runs).ToArray();
            if (!texts.SequenceEqual(runs.Select(run => run.Text)) || read.Pages.Count != layout.Pages.Count)
                throw new InvalidDataException(name + ": OFD read-back lost logical text or pages.");
            var restored = string.Concat(layout.Lines.Select(line => string.Concat(line.Runs.OrderBy(run => run.SourceStart).Select(run => run.Text)) +
                Sample.Substring(line.SourceStart + line.SourceLength - line.TerminatorLength, line.TerminatorLength)));
            if (restored != Sample)
                throw new InvalidDataException(name + ": line source intervals do not restore the original text.");

            var pdf = Path.Combine(output, name + ".pdf");
            await using (var input = File.OpenRead(ofd))
            await using (var stream = File.Create(pdf)) await new OfdToPdfConverter().ConvertAsync(input, stream);
            var svg = Path.Combine(output, name + ".svg");
            await using (var input = File.OpenRead(ofd))
            await using (var stream = File.Create(svg)) await new OfdToSvgConverter().ConvertAsync(input, stream);
            receipts.Add(new
            {
                name, layout.UnicodeVersion, layout.EngineVersion, layout.ShapeCalls, pages = layout.Pages.Count,
                lines = layout.Lines.Select(line => new { line.PageIndex, line.SourceStart, line.SourceLength,
                    line.TerminatorLength, line.BaseLevel, line.TopMillimeters, line.BaselineMillimeters,
                    line.ContentLeftMillimeters, line.ContentWidthMillimeters,
                    runs = line.Runs.Select(run => new { run.SourceStart, run.Text, face = run.Face.FaceSha256,
                        run.Script, direction = run.Direction.ToString(), run.BidiLevel, run.OriginXMillimeters,
                        gids = run.Glyphs.Select(glyph => glyph.GlyphId) }) }),
                outputs = new[] { ofd, pdf, svg }.Select(path => new { name = Path.GetFileName(path),
                    bytes = new FileInfo(path).Length,
                    sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant() }),
            });
        }
        await File.WriteAllTextAsync(Path.Combine(output, "r22-paragraph-receipt.json"), JsonSerializer.Serialize(new
        {
            sourceSha = Environment.GetEnvironmentVariable("R21_SOURCE_SHA"),
            source = Sample,
            cases = receipts,
        }, new JsonSerializerOptions { WriteIndented = true }) + "\n");
        Console.WriteLine("R22 paragraph layout OFD/PDF/SVG candidates: " + output);
    }
}
