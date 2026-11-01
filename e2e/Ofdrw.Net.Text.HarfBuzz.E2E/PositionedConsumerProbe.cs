using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Text.HarfBuzz;

internal static class PositionedConsumerProbe
{
    internal static async Task Run(string output, TextFontFace dejavu, TextFontFace noto)
    {
        var rows = new List<object>();
        var files = new List<string>();
        foreach (var item in new[]
        {
            (Name: "dejavu-lines", Text: "AV\nffi\nfi\na\u0323\njAV", Face: dejavu, Size: 4d),
            (Name: "dejavu-size6", Text: "fi", Face: dejavu, Size: 6d),
            (Name: "noto-size6", Text: "AV", Face: noto, Size: 6d),
        })
        {
            var result = HarfBuzzTextPrelayout.Create().Layout(new(item.Text, new[] { item.Face }, item.Size, 100));
            var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(result, 100, 75, 10, 10);
            var elements = package.Pages.Single().Elements.OfType<OfdTextElement>().ToArray();
            var lines = result.Lines.Where(line => line.SourceLength > 0).ToArray();
            if (elements.Length != lines.Length) throw new InvalidDataException("Typed consumer lost a line.");
            var captured = new List<object>();
            for (var index = 0; index < lines.Length; index++)
            {
                var line = lines[index];
                var ink = line.InkBoundsMillimeters;
                var left = Math.Min(0, ink?.X ?? 0);
                var top = Math.Min(0, ink?.Y - line.TopMillimeters ?? 0);
                var run = elements[index].Runs.Single();
                var typed = run.PositionedGlyphs!;
                var strict = TextPositionedRun.Create(line.ShapedRun, item.Size, -left,
                    line.BaselineMillimeters - line.TopMillimeters - top);
                if (typed.SourceStart != line.SourceStart || typed.FaceSha256 != item.Face.FaceSha256 ||
                    typed.LogicalText != item.Text.Substring(line.SourceStart, line.SourceLength) ||
                    typed.Glyphs.Count != line.ShapedRun.Glyphs.Count)
                    throw new InvalidDataException("Typed consumer lost source or face identity.");
                double penX = 0, penY = 0;
                var scale = item.Size / item.Face.UnitsPerEm;
                var origins = new List<double>();
                for (var i = 0; i < typed.Glyphs.Count; i++)
                {
                    var glyph = line.ShapedRun.Glyphs[i]; var actual = typed.Glyphs[i];
                    if ((actual.GlyphId, actual.ClusterStart, actual.ClusterLength) !=
                        (glyph.GlyphId, glyph.ClusterStart, glyph.ClusterLength))
                        throw new InvalidDataException("Typed consumer lost original GID/cluster.");
                    var x = -left + (penX + glyph.XOffset) * scale;
                    var y = line.BaselineMillimeters - line.TopMillimeters - top - (penY + glyph.YOffset) * scale;
                    SameBits(x, actual.OriginXMillimeters); SameBits(y, actual.OriginYMillimeters);
                    SameBits(glyph.XAdvance * scale, actual.XAdvanceMillimeters);
                    SameBits(glyph.YAdvance * scale, actual.YAdvanceMillimeters);
                    SameBits(strict.Glyphs[i].OriginXMillimeters, actual.OriginXMillimeters);
                    SameBits(strict.Glyphs[i].OriginYMillimeters, actual.OriginYMillimeters);
                    origins.Add(x); penX += glyph.XAdvance; penY += glyph.YAdvance;
                }
                var delta = origins.Count > 1 ? string.Join(" ", origins.Skip(1).Select((x, i) =>
                    (x - origins[i]).ToString("0.###############", CultureInfo.InvariantCulture))) : null;
                if (run.DeltaX != delta || run.DeltaY is not null)
                    throw new InvalidDataException("Typed consumer changed adjacent-origin deltas.");
                captured.Add(new { line.SourceStart, line.SourceLength, line.TopMillimeters, line.BaselineMillimeters,
                    line.AdvanceMillimeters, line.InkBoundsMillimeters, boundaryLeft = left, boundaryTop = top,
                    designGlyphs = line.ShapedRun.Glyphs, glyphs = typed.Glyphs,
                    verticalAdvanceBits = typed.Glyphs.Select(g => BitConverter.DoubleToInt64Bits(g.YAdvanceMillimeters)),
                    run.DeltaX, run.DeltaY });
            }
            var prefix = Path.Combine(output, "r22-shared-consumer-" + item.Name);
            var ofd = prefix + ".ofd"; var pdf = prefix + ".pdf";
            var svg = prefix + ".svg"; var png = prefix + ".png";
            await using (var target = New(ofd))
            {
                var write = await new OfdPackageWriter().WriteWithResultAsync(package, target);
                if (!write.FontEmbedding.Single().IsSubset) throw new InvalidDataException("Typed consumer lost subsetting.");
            }
            await using (var input = File.OpenRead(ofd)) await using (var target = New(pdf))
                await new OfdToPdfConverter().ConvertAsync(input, target);
            await using (var input = File.OpenRead(ofd)) await using (var target = New(svg))
                await new OfdToSvgConverter().ConvertAsync(input, target);
            await using (var input = File.OpenRead(ofd)) await using (var target = New(png))
                await new OfdToImageConverter().ConvertAsync(input, target);
            await using (var input = File.OpenRead(ofd))
            {
                var read = await new OfdReader().ReadAsync(input);
                if (!read.Pages.Single().Elements.OfType<OfdTextElement>().Select(element => element.Text)
                    .SequenceEqual(elements.Select(element => element.Text)))
                    throw new InvalidDataException("Typed consumer OFD changed original text.");
            }
            var svgGids = XDocument.Load(svg).Descendants().Where(node => node.Name.LocalName == "path")
                .Select(node => (uint?)node.Attribute("data-original-gid")).Where(gid => gid.HasValue)
                .Select(gid => gid!.Value);
            if (!svgGids.SequenceEqual(lines.SelectMany(line => line.ShapedRun.Glyphs).Select(glyph => glyph.GlyphId)))
                throw new InvalidDataException("Typed consumer SVG changed original GIDs.");
            rows.Add(new { item.Name, item.Text, fontSizeMillimeters = item.Size, item.Face.FaceSha256,
                item.Face.UnitsPerEm, lines = captured });
            files.AddRange(new[] { ofd, pdf, svg, png });
        }
        await File.WriteAllTextAsync(Path.Combine(output, "r22-shared-position-consumer.json"),
            JsonSerializer.Serialize(new
            {
                sourceSha = Environment.GetEnvironmentVariable("R21_SOURCE_SHA"), cases = rows,
                outputs = files.Select(path => new { name = Path.GetFileName(path), bytes = new FileInfo(path).Length,
                    sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant() }),
                scope = "Three real typed OFD/PDF/SVG/PNG pages, seven native lines. Shared arithmetic only; existing admission, strict query guards and legacy writer behavior remain separate. No Flow, new scripts/styles, precision-format repair or overall visual acceptance."
            }, new JsonSerializerOptions { WriteIndented = true }) + "\n");
        Console.WriteLine("R22 real typed consumer shared arithmetic, source/GID and bitwise geometry passed.");
    }

    private static FileStream New(string path) => new(path, FileMode.CreateNew, FileAccess.Write);
    private static void SameBits(double expected, double actual)
    {
        if (BitConverter.DoubleToInt64Bits(expected) != BitConverter.DoubleToInt64Bits(actual))
            throw new InvalidDataException("Typed consumer changed legacy geometry bits.");
    }
}
