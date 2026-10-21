using System.Security.Cryptography;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

internal static class NamedAreaSamples
{
    internal static async Task Run(string directory)
    {
        Directory.CreateDirectory(directory);
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 170;
        flow.Options.PageHeightMillimeters = 120;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 15;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 15;
        flow.Blocks.Add(new Paragraph("命名区域 / Named Areas") { Alignment = ParagraphAlignment.Center, SpaceAfterMillimeters = 3 });
        var styled = new Paragraph();
        styled.Spans.Add(new Span("中文和 proportional English. "));
        styled.Spans.Add(new Span("粗体 Bold ") { Bold = true });
        styled.Spans.Add(new Span("斜体 Italic ") { Italic = true });
        styled.Spans.Add(new Span("红色 Red") { Color = new OfdColor(180, 20, 20) });
        flow.Blocks.Add(styled);
        var table = new Table { SpaceBeforeMillimeters = 2, SpaceAfterMillimeters = 2 };
        table.ColumnWidthsMillimeters.Add(70); table.ColumnWidthsMillimeters.Add(70);
        var row = new Row { MinimumHeightMillimeters = 14 };
        row.Cells.Add(new Cell("姓名 / Name") { BackgroundColor = new OfdColor(220, 235, 250) });
        row.Cells.Add(new Cell("填写区域 ↓") { BackgroundColor = new OfdColor(255, 245, 210) });
        table.Rows.Add(row); flow.Blocks.Add(table);
        flow.Blocks.Add(new AreaHolderBlock("customer", 20));
        flow.Blocks.Add(new Paragraph("区域后文字不应移动 / after region") { SpaceBeforeMillimeters = 2 });
        flow.Blocks.Add(new Paragraph("第二页 / Page Two") { PageBreakBefore = true, Alignment = ParagraphAlignment.Center });
        flow.Blocks.Add(new AreaHolderBlock("logo", 24));
        flow.Blocks.Add(new Paragraph("图像后文字 / after image"));

        var blankPath = Path.Combine(directory, "named-areas-blank.ofd");
        var filledPath = Path.Combine(directory, "named-areas-filled.ofd");
        await using (var output = File.Create(blankPath)) await new OfdPackageWriter().WriteAsync(flow.Render(), output);
        var blank = await Read(blankPath);
        var outside = blank.Pages.SelectMany((page, index) => page.Elements.OfType<OfdTextElement>()
            .Select(text => (index, text.Text, text.XMillimeters, text.YMillimeters))).ToArray();
        var editor = new OfdNamedAreaEditor();
        if (editor.GetAreas(blank).Count != 2) throw new InvalidOperationException("Named area index did not survive reopen.");
        editor.FillText(blank, "customer", "王小明 / Alice Zhang", new NamedAreaTextOptions { FontSizeMillimeters = 5 });
        using (var image = new Image<Rgba32>(40, 20))
        {
            for (var y = 0; y < 20; y++) for (var x = 0; x < 40; x++)
                image[x, y] = x < 20 ? new Rgba32(30, 100, 190) : new Rgba32(235, 190, 35);
            await using var encoded = new MemoryStream();
            image.SaveAsPng(encoded);
            editor.FillImage(blank, "logo", encoded.ToArray(), "image/png", 40, 20);
        }
        await using (var output = File.Create(filledPath)) await new OfdPackageWriter().WriteAsync(blank, output);
        var filled = await Read(filledPath);
        if (editor.GetAreas(filled).Any(area => !area.IsFilled)) throw new InvalidOperationException("Fill state did not survive reopen.");
        foreach (var before in outside)
        {
            var after = filled.Pages[before.index].Elements.OfType<OfdTextElement>().Single(text => text.Text == before.Text);
            if (Math.Abs(before.XMillimeters - after.XMillimeters) > 0.00001 ||
                Math.Abs(before.YMillimeters - after.YMillimeters) > 0.00001)
                throw new InvalidOperationException("Content outside the named area moved.");
        }
        if (!filled.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>()
            .Any(text => text.Text == "王小明 / Alice Zhang"))
            throw new InvalidOperationException("Filled logical text was not retained.");
        foreach (var stem in new[] { "named-areas-blank", "named-areas-filled" })
        {
            var ofd = Path.Combine(directory, stem + ".ofd");
            var pdf = Path.Combine(directory, stem + ".pdf");
            await using (var input = File.OpenRead(ofd))
            await using (var output = File.Create(pdf))
                await new OfdToPdfConverter().ConvertAsync(input, output);
            Console.WriteLine($"{stem}: ofdSha256={Hash(ofd)} pdfSha256={Hash(pdf)} pages={filled.Pages.Count}");
        }
    }

    private static async Task<OfdDocumentPackage> Read(string path)
    {
        await using var input = File.OpenRead(path);
        return await new OfdReader().ReadAsync(input);
    }
    private static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();
}
