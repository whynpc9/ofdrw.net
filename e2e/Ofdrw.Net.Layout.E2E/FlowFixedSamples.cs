using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

internal static class FlowFixedSamples
{
    internal static async Task Run(string directory)
    {
        Directory.CreateDirectory(directory);
        byte[] encoded;
        using (var image = new Image<Rgba32>(40, 20))
        {
            for (var y = 0; y < 20; y++) for (var x = 0; x < 40; x++)
                image[x, y] = x < 20 ? new Rgba32(30, 100, 190) : new Rgba32(235, 190, 35);
            await using var stream = new MemoryStream();
            image.SaveAsPng(stream);
            encoded = stream.ToArray();
        }
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 170;
        flow.Options.PageHeightMillimeters = 120;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 15;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 15;
        var pages = new List<FlowBoundary>();
        var placements = new List<FlowElementPlacement>();
        flow.OnPage = (graphics, page) =>
        {
            pages.Add(page);
            graphics.FillRectangle(new OfdBrush(new OfdColor(245, 249, 253)),
                0, 0, page.WidthMillimeters, 10);
        };
        flow.OnElement = (_, placement) => placements.Add(placement);

        flow.Blocks.Add(new Paragraph("流式图片与 Canvas / Flow image and canvas")
        { Alignment = ParagraphAlignment.Center, SpaceAfterMillimeters = 4 });
        var styled = new Paragraph { SpaceAfterMillimeters = 4 };
        styled.Spans.Add(new Span("中文 English "));
        styled.Spans.Add(new Span("AV spacing")
        { FontWeight = 500, LetterSpacingMillimeters = 0.6, Underline = true, Color = new OfdColor(170, 30, 30) });
        styled.Spans.Add(new Span(" italic") { Italic = true });
        flow.Blocks.Add(styled);
        var table = new Table { SpaceAfterMillimeters = 4 };
        table.ColumnWidthsMillimeters.Add(70);
        table.ColumnWidthsMillimeters.Add(70);
        var row = new Row { MinimumHeightMillimeters = 18 };
        row.Cells.Add(new Cell("蓝黄图片 / image") { BackgroundColor = new OfdColor(220, 235, 250) });
        row.Cells.Add(new Cell("画布下一页 / canvas next") { BackgroundColor = new OfdColor(255, 245, 210) });
        table.Rows.Add(row);
        flow.Blocks.Add(table);
        var imageBlock = new ImageBlock(encoded, 60, 40) { Alignment = ParagraphAlignment.Center, SpaceAfterMillimeters = 4 };
        flow.Blocks.Add(imageBlock);
        var canvasBlock = new CanvasBlock(100, 28, (graphics, boundary) =>
        {
            graphics.FillRectangle(new OfdBrush(new OfdColor(226, 240, 224)), 0, 0,
                boundary.WidthMillimeters, boundary.HeightMillimeters);
            graphics.DrawLine(new OfdPen(new OfdColor(30, 90, 170), 0.8), 2, 2,
                boundary.WidthMillimeters - 2, boundary.HeightMillimeters - 2);
            graphics.DrawString("Canvas 画布", new OfdFont("SimSun", 5),
                new OfdBrush(OfdColor.Black), 5, 14);
        }) { Alignment = ParagraphAlignment.Center };
        flow.Blocks.Add(canvasBlock);
        flow.Blocks.Add(new AreaHolderBlock("reserved", 10));
        flow.Blocks.Add(new Paragraph("画布后文字 / after canvas"));

        var package = flow.Render();
        if (package.Pages.Count != 2 || pages.Count != 2 ||
            placements.Count(placement => ReferenceEquals(placement.Block, imageBlock)) != 1 ||
            placements.Count(placement => ReferenceEquals(placement.Block, canvasBlock)) != 1 ||
            placements.Single(placement => ReferenceEquals(placement.Block, imageBlock)).Boundary.PageIndex != 0 ||
            placements.Single(placement => ReferenceEquals(placement.Block, canvasBlock)).Boundary.PageIndex != 1)
            throw new InvalidOperationException("Flow image/canvas pagination or callbacks changed.");
        var ofdPath = Path.Combine(directory, "flow-fixed.ofd");
        var pdfPath = Path.Combine(directory, "flow-fixed.pdf");
        await using (var output = File.Create(ofdPath)) await new OfdPackageWriter().WriteAsync(package, output);
        await using (var input = File.OpenRead(ofdPath))
        await using (var output = File.Create(pdfPath))
            await new OfdToPdfConverter().ConvertAsync(input, output);
        await using var read = File.OpenRead(ofdPath);
        var reopened = await new OfdReader().ReadAsync(read);
        var areas = new OfdNamedAreaEditor().GetAreas(reopened);
        if (reopened.Pages.Count != 2 ||
            reopened.Pages.SelectMany(page => page.Elements).OfType<OfdImageElement>().Count() != 1 ||
            areas.Count != 1 || areas[0].Name != "reserved" || areas[0].PageIndex != 1 ||
            !string.Concat(reopened.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
                .Select(element => element.Text)).Contains("画布后文字"))
            throw new InvalidOperationException("Flow image/canvas output did not round trip.");
        var evidence = new
        {
            sourceSha = Environment.GetEnvironmentVariable("R04_SOURCE_SHA"),
            inputPngSha256 = Convert.ToHexString(SHA256.HashData(encoded)),
            chain = "FlowDocument -> native OFD -> PDF",
            pages = reopened.Pages.Count,
            imageBoundary = placements.Single(placement => ReferenceEquals(placement.Block, imageBlock)).Boundary,
            canvasBoundary = placements.Single(placement => ReferenceEquals(placement.Block, canvasBlock)).Boundary,
            ofdSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ofdPath))),
            pdfSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(pdfPath)))
        };
        await File.WriteAllTextAsync(Path.Combine(directory, "flow-fixed.json"),
            JsonSerializer.Serialize(evidence, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"flow-fixed: pages={evidence.pages}, ofd={evidence.ofdSha256}, pdf={evidence.pdfSha256}");
        await RunLetterSpacingWrap(directory);
    }

    private static async Task RunLetterSpacingWrap(string directory)
    {
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 19;
        flow.Options.PageHeightMillimeters = 50;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 5;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 5;
        var paragraph = new Paragraph();
        paragraph.Spans.Add(new Span("X A\u00A0B")
        { FontSizeMillimeters = 4, LetterSpacingMillimeters = 1, Underline = true });
        flow.Blocks.Add(paragraph);
        var package = flow.Render();
        var text = package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().ToArray();
        if (package.Pages.Count != 1 || text.Length != 2 || text[0].Text != "X " ||
            text[1].Text != "A\u00A0B" || text[1].YMillimeters <= text[0].YMillimeters)
            throw new InvalidOperationException("Letter-spaced nonbreaking unit did not wrap onto a fresh line.");
        var ofd = Path.Combine(directory, "flow-letterspacing-wrap.ofd");
        var pdf = Path.Combine(directory, "flow-letterspacing-wrap.pdf");
        await using (var output = File.Create(ofd)) await new OfdPackageWriter().WriteAsync(package, output);
        await using (var input = File.OpenRead(ofd))
        await using (var output = File.Create(pdf))
            await new OfdToPdfConverter().ConvertAsync(input, output);
        await File.WriteAllTextAsync(Path.Combine(directory, "flow-letterspacing-wrap.json"),
            JsonSerializer.Serialize(new
            {
                sourceSha = Environment.GetEnvironmentVariable("R04_SOURCE_SHA"),
                chain = "FlowDocument -> native OFD -> PDF",
                pages = 1,
                ofdSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(ofd))),
                pdfSha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(pdf)))
            }, new JsonSerializerOptions { WriteIndented = true }));
    }
}
