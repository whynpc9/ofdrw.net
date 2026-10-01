using System.IO.Compression;
using System.Security;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using UglyToad.PdfPig;

internal static class PublicTableSamples
{
    internal static async Task Run(string directory, string baselineDocx)
    {
        var flow = new FlowDocument();
        flow.Options.PageWidthMillimeters = 170; flow.Options.PageHeightMillimeters = 150;
        flow.Options.MarginLeftMillimeters = flow.Options.MarginRightMillimeters = 15;
        flow.Options.MarginTopMillimeters = flow.Options.MarginBottomMillimeters = 15;
        flow.Blocks.Add(new Paragraph("公开表格 / Public Tables") { SpaceAfterMillimeters = 4 });
        var table = new Table { SpaceAfterMillimeters = 4 };
        table.ColumnWidthsMillimeters.Add(50); table.ColumnWidthsMillimeters.Add(50); table.ColumnWidthsMillimeters.Add(40);
        var header = new Row { MinimumHeightMillimeters = 15 };
        var heading = new Cell("合并标题 / Merged heading") { ColumnSpan = 2, BackgroundColor = new OfdColor(210, 230, 250), VerticalAlignment = CellVerticalAlignment.Center };
        heading.Paragraphs[0].Alignment = ParagraphAlignment.Center; header.Cells.Add(heading);
        var final = new Cell("右 / Right") { BackgroundColor = new OfdColor(210, 230, 250), VerticalAlignment = CellVerticalAlignment.Bottom };
        final.Paragraphs[0].Alignment = ParagraphAlignment.Right; header.Cells.Add(final); table.Rows.Add(header);
        for (var i = 1; i <= 11; i++)
        {
            var row = new Row { MinimumHeightMillimeters = 19 };
            row.Cells.Add(new Cell($"行 {i:00} 中文 / WiMi {i:00}") { BackgroundColor = new OfdColor(246, 248, 250) });
            var styled = new Cell { VerticalAlignment = CellVerticalAlignment.Center };
            var p = new Paragraph { Alignment = ParagraphAlignment.Center };
            p.Spans.Add(new Span($"Bold {i:00}") { Bold = true, Color = new OfdColor(190, 30, 30) });
            p.Spans.Add(new Span(" / ")); p.Spans.Add(new Span("Italic") { Italic = true }); styled.Paragraphs.Add(p); row.Cells.Add(styled);
            var right = new Cell($"金额 {i * 12}.50") { VerticalAlignment = CellVerticalAlignment.Bottom, BackgroundColor = new OfdColor(255, 246, 218) };
            right.Paragraphs[0].Alignment = ParagraphAlignment.Right; row.Cells.Add(right); table.Rows.Add(row);
        }
        flow.Blocks.Add(table); flow.Blocks.Add(new Paragraph("表后正文 / After the table"));
        var expected = Compact(string.Concat(flow.Blocks.SelectMany(block => block is Paragraph p ? new[] { p } :
            ((Table)block).Rows.SelectMany(r => r.Cells).SelectMany(c => c.Paragraphs)).SelectMany(p => p.Spans).Select(s => s.Text)));
        await using (var output = File.Create(Path.Combine(directory, "tables-public.ofd")))
            await new OfdPackageWriter().WriteAsync(flow.Render(), output);

        var source = Path.Combine(directory, "tables-source.docx");
        File.Delete(source);
        using (var archive = ZipFile.Open(source, ZipArchiveMode.Create))
        {
            void Entry(string name, string xml) { using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(xml); }
            Entry("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Default Extension=\"xml\" ContentType=\"application/xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Entry("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            string ParagraphXml(Paragraph p) => "<w:p><w:pPr><w:jc w:val=\"" + p.Alignment.ToString().ToLowerInvariant() + "\"/></w:pPr>" +
                string.Concat(p.Spans.Select(s => "<w:r><w:rPr>" + (s.Bold ? "<w:b/>" : "") + (s.Italic ? "<w:i/>" : "") +
                (s.Color is null ? "" : $"<w:color w:val=\"{s.Color.Red:X2}{s.Color.Green:X2}{s.Color.Blue:X2}\"/>") + "</w:rPr><w:t xml:space=\"preserve\">" + SecurityElement.Escape(s.Text) + "</w:t></w:r>")) + "</w:p>";
            var xml = new StringBuilder("<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>");
            xml.Append(ParagraphXml((Paragraph)flow.Blocks[0]));
            xml.Append("<w:tbl><w:tblPr><w:tblBorders>");
            foreach (var side in new[] { "top", "bottom", "left", "right", "insideH", "insideV" }) xml.Append($"<w:{side} w:val=\"single\" w:sz=\"5\" w:color=\"000000\"/>");
            xml.Append("</w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"2835\"/><w:gridCol w:w=\"2835\"/><w:gridCol w:w=\"2268\"/></w:tblGrid>");
            foreach (var row in table.Rows)
            {
                xml.Append("<w:tr>");
                foreach (var cell in row.Cells)
                {
                    xml.Append($"<w:tc><w:tcPr><w:gridSpan w:val=\"{cell.ColumnSpan}\"/><w:vAlign w:val=\"{cell.VerticalAlignment.ToString().ToLowerInvariant()}\"/>");
                    if (cell.BackgroundColor is { } fill) xml.Append($"<w:shd w:fill=\"{fill.Red:X2}{fill.Green:X2}{fill.Blue:X2}\"/>");
                    xml.Append("</w:tcPr>"); foreach (var paragraph in cell.Paragraphs) xml.Append(ParagraphXml(paragraph)); xml.Append("</w:tc>");
                }
                xml.Append("</w:tr>");
            }
            xml.Append("</w:tbl>").Append(ParagraphXml((Paragraph)flow.Blocks[2]));
            xml.Append("<w:sectPr><w:pgSz w:w=\"9638\" w:h=\"6000\"/><w:pgMar w:top=\"850\" w:bottom=\"850\" w:left=\"850\" w:right=\"850\"/></w:sectPr></w:body></w:document>");
            Entry("word/document.xml", xml.ToString());
        }
        foreach (var (name, docx) in new[] { ("tables", source), ("baseline", baselineDocx) })
            foreach (var mode in new[] { "native", "default" })
            {
                await using var input = File.OpenRead(docx); await using var output = File.Create(Path.Combine(directory, name + "-" + mode + ".ofd"));
                var converter = mode == "native" ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }) : new DocxToOfdConverter();
                await converter.ConvertAsync(input, output);
            }
        var ns = XNamespace.Get("http://schemas.openxmlformats.org/wordprocessingml/2006/main");
        using var baselineArchive = ZipFile.OpenRead(baselineDocx);
        using var baselineXml = baselineArchive.GetEntry("word/document.xml")!.Open();
        var baselineExpected = Compact(string.Concat(XDocument.Load(baselineXml).Descendants(ns + "t").Select(e => e.Value)));
        foreach (var name in new[] { "tables-public", "tables-native", "tables-default", "baseline-native", "baseline-default" })
        {
            var ofd = Path.Combine(directory, name + ".ofd"); var pdf = Path.Combine(directory, name + ".pdf");
            await using (var input = File.OpenRead(ofd)) await using (var output = File.Create(pdf)) await new OfdToPdfConverter().ConvertAsync(input, output);
            await using var read = File.OpenRead(ofd); var package = await new OfdReader().ReadAsync(read);
            var text = Compact(string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => e.Text)));
            if (text != (name.StartsWith("tables") ? expected : baselineExpected)) throw new InvalidOperationException(name + ": OFD lost or repeated raw text.");
            using var pdfDocument = PdfDocument.Open(pdf); var pdfText = Compact(string.Concat(pdfDocument.GetPages().Select(p => p.Text)));
            if (pdfText != text) throw new InvalidOperationException(name + ": PDF lost or repeated text.");
            Console.WriteLine($"{name}: pages={package.Pages.Count}, text={text.Length}, ofdBytes={new FileInfo(ofd).Length}, pdfBytes={new FileInfo(pdf).Length}");
        }
    }
    private static string Compact(string text) => new(text.Where(c => !char.IsWhiteSpace(c)).ToArray());
}
