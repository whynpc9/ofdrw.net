using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit019RegressionTests
{
    private static BuiltInParagraphModel Paragraph(string prefix, params BuiltInPageFieldKind[] fields)
    {
        var paragraph = new BuiltInParagraphModel();
        paragraph.Inlines.Add(new BuiltInTextModel { Text = prefix, Format = { FontFamily = "Arial", FontSizePoints = 12 } });
        foreach (var kind in fields)
        {
            paragraph.Inlines.Add(new BuiltInPageNumberModel { Kind = kind, Format = { FontFamily = "Arial", FontSizePoints = 12 } });
            paragraph.Inlines.Add(new BuiltInTextModel { Text = "/", Format = { FontFamily = "Arial", FontSizePoints = 12 } });
        }
        return paragraph;
    }

    private static OfdDocumentPackage Render(BuiltInDocumentModel model) =>
        new BuiltInOfdRenderer(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default).Render(model);
    private static string Text(OfdPage page) => string.Concat(page.Elements.OfType<OfdTextElement>().Select(text => text.Text));
    private static BuiltInSectionModel Section(int pages, int start = 1)
    {
        var section = new BuiltInSectionModel { PageNumberStart = start };
        for (var index = 0; index < pages; index++)
        {
            var p = Paragraph("Body ", BuiltInPageFieldKind.Page, BuiltInPageFieldKind.TotalPages, BuiltInPageFieldKind.SectionPages);
            p.Format.PageBreakBefore = index > 0;
            section.Blocks.Add(p);
        }
        return section;
    }

    [Fact]
    public void Audit019_DecorationControl()
    {
        var model = new BuiltInDocumentModel(); var section = Section(3);
        section.DifferentFirstPage = true; section.DifferentOddAndEvenPages = true;
        section.FirstHeaders.Add(Paragraph("First ", BuiltInPageFieldKind.Page, BuiltInPageFieldKind.TotalPages));
        section.EvenHeaders.Add(Paragraph("Even ", BuiltInPageFieldKind.Page, BuiltInPageFieldKind.TotalPages));
        section.Headers.Add(Paragraph("Default ", BuiltInPageFieldKind.Page, BuiltInPageFieldKind.TotalPages));
        model.Sections.Add(section); var result = Render(model);
        Assert.Equal(3, result.Pages.Count);
        Assert.Contains("First 1/3/", Text(result.Pages[0])); Assert.Contains("Even 2/3/", Text(result.Pages[1])); Assert.Contains("Default 3/3/", Text(result.Pages[2]));
    }

    [Fact]
    public void Audit019_BodyTotals()
    {
        var model = new BuiltInDocumentModel(); model.Sections.Add(Section(3)); var result = Render(model);
        Assert.Equal(3, result.Pages.Count);
        for (var i = 0; i < 3; i++) Assert.Equal($"Body {i + 1}/3/3/", Text(result.Pages[i]));
    }

    [Fact]
    public void Audit019_SectionRestart()
    {
        var model = new BuiltInDocumentModel(); model.Sections.Add(Section(2)); model.Sections.Add(Section(3, 10)); var result = Render(model);
        Assert.Equal(5, result.Pages.Count);
        for (var i = 0; i < 5; i++) Assert.Equal($"Body {(i < 2 ? i + 1 : i + 8)}/5/{(i < 2 ? 2 : 3)}/", Text(result.Pages[i]));
    }

    [Fact]
    public void Audit019_TableActualPage()
    {
        var model = new BuiltInDocumentModel(); var section = Section(1, 9); model.Sections.Add(section);
        section.PageHeightPoints = 150; section.MarginTopPoints = section.MarginBottomPoints = 20;
        var table = new BuiltInTableModel(); section.Blocks.Add(table);
        for (var i = 0; i < 12; i++)
        {
            var row = new BuiltInTableRowModel(); var cell = new BuiltInTableCellModel(); row.Cells.Add(cell); table.Rows.Add(row);
            cell.Paragraphs.Add(Paragraph($"Row{i:D2} ", BuiltInPageFieldKind.Page, BuiltInPageFieldKind.TotalPages, BuiltInPageFieldKind.SectionPages));
        }
        var result = Render(model); Assert.True(result.Pages.Count > 1);
        for (var i = 0; i < result.Pages.Count; i++)
        {
            foreach (var field in result.Pages[i].Elements.OfType<OfdTextElement>().Where(t => t.Text == (i + 9).ToString())) Assert.True(field.WidthMillimeters > 0);
            var text = Text(result.Pages[i]);
            foreach (var row in System.Text.RegularExpressions.Regex.Matches(text, @"Row\d\d (\d+)/(\d+)/(\d+)/").Cast<System.Text.RegularExpressions.Match>())
            { Assert.Equal(i + 9, int.Parse(row.Groups[1].Value)); Assert.Equal(result.Pages.Count, int.Parse(row.Groups[2].Value)); Assert.Equal(result.Pages.Count, int.Parse(row.Groups[3].Value)); }
        }
        Assert.Equal(12, result.Pages.Sum(p => System.Text.RegularExpressions.Regex.Matches(Text(p), "Row").Count));
    }

    [Fact]
    public void Audit019_DigitTransition()
    {
        var model = new BuiltInDocumentModel(); var section = Section(10); model.Sections.Add(section);
        section.PageWidthPoints = 175; section.MarginLeftPoints = section.MarginRightPoints = 72;
        section.Blocks.Clear();
        for (var i = 0; i < 10; i++)
        {
            var p = Paragraph("XX", BuiltInPageFieldKind.TotalPages); p.Format.PageBreakBefore = i > 0; section.Blocks.Add(p);
        }
        var result = Render(model); Assert.Equal(10, result.Pages.Count);
        foreach (var page in result.Pages)
        {
            Assert.Equal("XX10/", Text(page));
            var texts = page.Elements.OfType<OfdTextElement>().ToArray();
            Assert.True(texts.Max(t => t.YMillimeters) > texts.Min(t => t.YMillimeters), "Two-digit final total must be remeasured and wrap.");
            Assert.All(texts, t => Assert.True(t.XMillimeters + t.WidthMillimeters <= (175 - 72) * 25.4 / 72 + .001));
        }
    }

    [Fact]
    public void Audit019_TableDigitTransition_RemeasuresOnPlacementPage()
    {
        var model = new BuiltInDocumentModel();
        var section = new BuiltInSectionModel { PageNumberStart = 9, PageWidthPoints = 179,
            PageHeightPoints = 100, MarginTopPoints = 20, MarginBottomPoints = 20 };
        model.Sections.Add(section);
        var filler = Paragraph("Filler"); filler.Format.SpaceAfterPoints = 25; section.Blocks.Add(filler);
        var table = new BuiltInTableModel(); var row = new BuiltInTableRowModel(); var cell = new BuiltInTableCellModel();
        table.Rows.Add(row); row.Cells.Add(cell); cell.Paragraphs.Add(Paragraph("XX", BuiltInPageFieldKind.Page)); section.Blocks.Add(table);
        var result = Render(model); Assert.Equal(2, result.Pages.Count); Assert.Equal("XX10/", Text(result.Pages[1]));
        var texts = result.Pages[1].Elements.OfType<OfdTextElement>().ToArray();
        Assert.True(texts.Max(t => t.YMillimeters) > texts.Min(t => t.YMillimeters));
        Assert.All(texts, t => Assert.True(t.YMillimeters + t.HeightMillimeters <= 80 * 25.4 / 72 + .001));
    }

    [Fact]
    public void Audit019_BoundedConvergence()
    {
        var calls = 0;
        var error = Assert.Throws<InvalidDataException>(() => BuiltInOfdRenderer.FinalizePagination(() => { calls++; return false; }));
        Assert.Equal(8, calls); Assert.Contains("did not converge", error.Message);
        calls = 0;
        BuiltInOfdRenderer.FinalizePagination(() => ++calls == 8);
        Assert.Equal(8, calls);
    }

    [Fact]
    public async Task Audit019_BoundedFailure_LeavesCallerOutputUnchanged()
    {
        using var source = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new W.Document(new W.Body(new W.Paragraph(new W.Run(new W.Text("one"))),
                new W.Paragraph(new W.ParagraphProperties(new W.PageBreakBefore()), new W.SimpleField(new W.Run(new W.Text("stale"))) { Instruction = "NUMPAGES" })));
        }
        source.Position = 0; using var output = new MemoryStream(); output.WriteByte(42);
        await Assert.ThrowsAsync<InvalidDataException>(() => new DocxToOfdConverter(new DocxConversionOptions { MaxPageCount = 1 }).ConvertAsync(source, output));
        Assert.Equal(new byte[] { 42 }, output.ToArray()); Assert.Equal(1, output.Position);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit019_PublicBodyAndTable(bool explicitNative)
    {
        using var source = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(source, WordprocessingDocumentType.Document, true))
        {
            var main = doc.AddMainDocumentPart();
            string Field(string name) => $"<w:fldSimple w:instr=\"{name}\"><w:r><w:t>stale</w:t></w:r></w:fldSimple>";
            main.Document = new W.Document($"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:t>First 中文</w:t></w:r></w:p><w:p><w:pPr><w:pageBreakBefore/></w:pPr><w:r><w:t>Body </w:t></w:r>{Field("PAGE")}{Field("NUMPAGES")}{Field("SECTIONPAGES")}</w:p><w:tbl><w:tblGrid><w:gridCol w:w=\"7200\"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>Cell </w:t></w:r>{Field("PAGE")}{Field("NUMPAGES")}{Field("SECTIONPAGES")}</w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>");
        }
        var options = new DocxConversionOptions(); if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        var converter = new DocxToOfdConverter(options); source.Position = 0; using var output = new MemoryStream();
        var result = await converter.ConvertWithResultAsync(source, output); Assert.True(result.OriginalTextPreserved); Assert.Equal(2, result.SourcePageCount);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("First 中文", Text(package.Pages[0])); Assert.Equal("Body 222Cell 222", Text(package.Pages[1]));
        source.Position = 0; using var selected = new MemoryStream();
        var selection = await converter.ConvertWithResultAsync(source, selected, [1, 0, 1]); Assert.Equal(new[] { 1, 0, 1 }, selection.SourcePages);
        selected.Position = 0; var selectedPackage = await new OfdReader().ReadAsync(selected);
        Assert.Equal(Text(package.Pages[1]), Text(selectedPackage.Pages[0])); Assert.Equal(Text(package.Pages[1]), Text(selectedPackage.Pages[2]));
    }
}
