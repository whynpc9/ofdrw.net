using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using UglyToad.PdfPig;

internal static class ComplexTableSamples
{
    private static readonly byte[] Pixel = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");

    internal static async Task Run(string directory)
    {
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 150;
        flow.Options.PageHeightMillimeters = 100;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 10;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 10;
        flow.Options.DefaultFontSizeMillimeters = 4;
        var table = new Table { RepeatHeaderRowCount = 1, SplitRowsAcrossPages = true,
            SpaceAfterMillimeters = 3 };
        foreach (var width in new[] { 45d, 45d, 40d }) table.ColumnWidthsMillimeters.Add(width);
        var header = new Row { MinimumHeightMillimeters = 12 };
        var heading = new Cell("续页表头 / Header") { ColumnSpan = 2,
            BackgroundColor = new OfdColor(212, 230, 248), VerticalAlignment = CellVerticalAlignment.Center };
        heading.Paragraphs[0].Alignment = ParagraphAlignment.Center;
        header.Cells.Add(heading);
        header.Cells.Add(new Cell("金额 / Sum") { BackgroundColor = new OfdColor(212, 230, 248) });
        table.Rows.Add(header);

        var mergedTop = new Row { MinimumHeightMillimeters = 18 };
        mergedTop.Cells.Add(new Cell("纵向合并 / RowSpan") { RowSpan = 2,
            BackgroundColor = new OfdColor(241, 245, 253), VerticalAlignment = CellVerticalAlignment.Center });
        mergedTop.Cells.Add(new Cell("第一行 / First") { BackgroundColor = new OfdColor(248, 248, 248) });
        mergedTop.Cells.Add(new Cell("12.50")); table.Rows.Add(mergedTop);
        var mergedBottom = new Row { MinimumHeightMillimeters = 18 };
        mergedBottom.Cells.Add(new Cell("第二行 / Second") { BackgroundColor = new OfdColor(248, 248, 248) });
        mergedBottom.Cells.Add(new Cell("25.00")); table.Rows.Add(mergedBottom);

        var longRow = new Row();
        var longCell = new Cell { ColumnSpan = 2, BackgroundColor = new OfdColor(250, 250, 240) };
        var paragraph = new Paragraph();
        for (var i = 1; i <= 28; i++)
        {
            paragraph.Spans.Add(new Span($"段{i:00} 中文 / WiMi ")
            { Color = i % 4 == 0 ? new OfdColor(170, 35, 35) : null, Bold = i % 5 == 0,
                Italic = i % 7 == 0 });
            if (i < 28) paragraph.Spans.Add(new Span("\n"));
        }
        longCell.Paragraphs.Add(paragraph); longRow.Cells.Add(longCell);
        longRow.Cells.Add(new Cell { Image = new ImageBlock(Pixel, 25, 16)
        { Alignment = ParagraphAlignment.Center }, BackgroundColor = new OfdColor(240, 248, 240) });
        table.Rows.Add(longRow);
        var tail = new Row { MinimumHeightMillimeters = 15 };
        tail.Cells.Add(new Cell("终行 / Final") { ColumnSpan = 3,
            BackgroundColor = new OfdColor(235, 242, 250) });
        table.Rows.Add(tail);
        flow.Blocks.Add(table);
        flow.Blocks.Add(new Paragraph("表后正文 / After the table"));

        var placements = new List<FlowBoundary>();
        flow.OnElement = (_, placement) =>
        {
            if (ReferenceEquals(placement.Block, table)) placements.Add(placement.Boundary);
        };
        var package = flow.Render();
        var ofd = Path.Combine(directory, "complex-table.ofd");
        var pdf = Path.Combine(directory, "complex-table.pdf");
        await using (var output = File.Create(ofd)) await new OfdPackageWriter().WriteAsync(package, output);
        await using (var input = File.OpenRead(ofd))
        await using (var output = File.Create(pdf)) await new OfdToPdfConverter().ConvertAsync(input, output);
        await using var read = File.OpenRead(ofd);
        var reopened = await new OfdReader().ReadAsync(read);
        using var pdfDocument = PdfDocument.Open(pdf);
        var ofdText = Compact(string.Concat(reopened.Pages.SelectMany(page => page.Elements)
            .OfType<OfdTextElement>().Select(element => element.Text)));
        var pdfText = Compact(string.Concat(pdfDocument.GetPages().Select(page => page.Text)));
        for (var i = 1; i <= 28; i++)
        {
            var token = $"段{i:00}";
            if (Count(ofdText, token) != 1 || Count(pdfText, token) != 1)
                throw new InvalidOperationException($"Complex table lost or duplicated {token}.");
        }
        var tablePages = placements.Select(fragment => fragment.PageIndex).Distinct().ToArray();
        if (tablePages.Length < 3 || placements.Count < tablePages.Length ||
            reopened.Pages.Sum(page => page.Elements.OfType<OfdImageElement>().Count()) != 1)
            throw new InvalidOperationException("Complex table page, fragment, or image count is incorrect.");
        foreach (var expected in new[] { "纵向合并/RowSpan", "第一行/First", "第二行/Second",
                     "终行/Final", "表后正文/Afterthetable" })
            if (Count(ofdText, expected) != 1 || Count(pdfText, expected) != 1)
                throw new InvalidOperationException($"Complex table lost or duplicated {expected}.");
        foreach (var text in new[] { ofdText, pdfText })
            if (Count(text, "续页表头/Header") != tablePages.Length ||
                Count(text, "金额/Sum") != tablePages.Length)
                throw new InvalidOperationException("Complex table header copies do not match table pages.");
        foreach (var pageIndex in tablePages)
        {
            var pageText = Compact(string.Concat(reopened.Pages[pageIndex].Elements.OfType<OfdTextElement>().Select(element => element.Text)));
            if (!pageText.StartsWith("续页表头/Header", StringComparison.Ordinal))
                throw new InvalidOperationException("A complex table continuation lost its header.");
        }
        var outputs = new[] { ofd, pdf }.Select(path => new
        {
            name = Path.GetFileName(path), bytes = new FileInfo(path).Length,
            sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant()
        }).ToArray();
        var receipt = new
        {
            sourceSha = Environment.GetEnvironmentVariable("R19_SOURCE_SHA") ?? "unset",
            pages = reopened.Pages.Count,
            tableFragments = placements.Count,
            imageElements = 1,
            repeatedHeaderCopies = tablePages.Length,
            outputs
        };
        await File.WriteAllTextAsync(Path.Combine(directory, "complex-table-receipt.json"),
            JsonSerializer.Serialize(receipt, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine($"complex-table: pages={reopened.Pages.Count}, fragments={placements.Count}, ofdBytes={outputs[0].bytes}, pdfBytes={outputs[1].bytes}");
    }

    private static int Count(string text, string token)
    {
        var count = 0;
        for (var index = 0; (index = text.IndexOf(token, index, StringComparison.Ordinal)) >= 0; index += token.Length) count++;
        return count;
    }
    private static string Compact(string text) => new(text.Where(value => !char.IsWhiteSpace(value)).ToArray());
}
