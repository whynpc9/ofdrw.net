using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed partial class DocxConversionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Native_CellInheritedPageBreak_IsDiagnosedOrRejected(bool strict)
    {
        await using var input = CreateMinimalDocx("<w:tbl><w:tr><w:tc><w:p><w:pPr><w:pStyle w:val=\"CellBreak\"/></w:pPr><w:r><w:t>A</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        using (var document = DocumentFormat.OpenXml.Packaging.WordprocessingDocument.Open(input, true))
        {
            var part = document.MainDocumentPart!.AddNewPart<DocumentFormat.OpenXml.Packaging.StyleDefinitionsPart>();
            using var xml = new MemoryStream(System.Text.Encoding.UTF8.GetBytes("<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:style w:type=\"paragraph\" w:styleId=\"CellBreak\"><w:pPr><w:pageBreakBefore/></w:pPr></w:style></w:styles>"));
            part.FeedData(xml);
        }
        input.Position = 0; await using var output = new MemoryStream();
        var converter = new DocxToOfdConverter(new DocxConversionOptions { UnsupportedFeatureBehavior = strict ?
            UnsupportedDocxFeatureBehavior.Throw : UnsupportedDocxFeatureBehavior.Placeholder });
        if (strict)
        {
            await Assert.ThrowsAsync<NotSupportedException>(() => converter.ConvertAsync(input, output)); Assert.Equal(0, output.Length);
        }
        else
        {
            var result = await converter.ConvertWithResultAsync(input, output);
            Assert.Contains(result.Diagnostics, d => d.Code == "DOCX_CELL_PAGE_BREAK_DEGRADED");
        }
    }
    [Theory]
    [InlineData(true, "restart")]
    [InlineData(true, "")]
    [InlineData(false, "restart")]
    [InlineData(false, "")]
    public async Task Native_VerticalMerge_IsDiagnosedAndStrictModeRejects(bool explicitNative, string value)
    {
        var attribute = value.Length == 0 ? "" : $" w:val=\"{value}\"";
        var body = $"<w:tbl><w:tblGrid><w:gridCol w:w=\"4000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge{attribute}/></w:tcPr><w:p><w:r><w:t>merge text</w:t></w:r></w:p></w:tc></w:tr></w:tbl>";
        await using var input = CreateMinimalDocx(body); await using var output = new MemoryStream();
        var converter = explicitNative ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }) : new DocxToOfdConverter();
        var result = await converter.ConvertWithResultAsync(input, output);
        var warning = Assert.Single(result.Diagnostics, d => d.Code == "DOCX_VERTICAL_MERGE_DEGRADED");
        Assert.Contains("row 1, cell 1", warning.Message); Assert.Contains("independent cells", warning.Message);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("merge text", string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => e.Text)));
        await using var strictInput = CreateMinimalDocx(body); await using var strictOutput = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() => new DocxToOfdConverter(new DocxConversionOptions
        { OfdMode = DocxToOfdMode.Native, UnsupportedFeatureBehavior = UnsupportedDocxFeatureBehavior.Throw }).ConvertAsync(strictInput, strictOutput));
        Assert.Equal(0, strictOutput.Length);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public async Task Native_MergedTableRows_MoveTogetherWithoutDroppingText(bool explicitNative)
    {
        var rows = string.Concat(Enumerable.Range(0, 6).Select(i => $"""
            <w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:fill="DDEEFF"/></w:tcPr><w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:color w:val="C00000"/></w:rPr><w:t>合并 Row{i}</w:t></w:r></w:p></w:tc>
            <w:tc><w:p><w:r><w:t>Cell{i}</w:t></w:r></w:p></w:tc></w:tr>
            """));
        var body = $"<w:tbl><w:tblPr><w:tblBorders><w:top w:val=\"single\" w:sz=\"8\"/><w:bottom w:val=\"single\" w:sz=\"8\"/><w:insideH w:val=\"single\" w:sz=\"8\"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"1000\"/></w:tblGrid>{rows}</w:tbl><w:sectPr><w:pgSz w:w=\"6000\" w:h=\"2200\"/><w:pgMar w:top=\"283\" w:bottom=\"283\" w:left=\"283\" w:right=\"283\"/></w:sectPr>";
        await using var input = CreateMinimalDocx(body); await using var output = new MemoryStream();
        var converter = explicitNative ? new DocxToOfdConverter(new DocxConversionOptions { OfdMode = DocxToOfdMode.Native }) : new DocxToOfdConverter();
        await converter.ConvertAsync(input, output); output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.True(package.Pages.Count > 1);
        var all = package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().ToArray();
        Assert.Equal(string.Concat(Enumerable.Range(0, 6).Select(i => $"合并 Row{i}Cell{i}")), string.Concat(all.Select(e => e.Text)));
        foreach (var page in package.Pages)
        {
            var text = page.Elements.OfType<OfdTextElement>().ToArray();
            Assert.Equal(0, text.Length % 2);
            for (var i = 0; i < text.Length; i += 2) Assert.Equal(text[i].YMillimeters, text[i + 1].YMillimeters);
            Assert.All(page.Elements, e => Assert.True(e.YMillimeters + e.HeightMillimeters <= page.HeightMillimeters - 283 * 25.4 / 1440 + 0.001));
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(2)]
    public async Task Native_InvalidGridSpan_FailsAtomically(int span)
    {
        await using var input = CreateMinimalDocx($"<w:tbl><w:tblGrid><w:gridCol w:w=\"3000\"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val=\"{span}\"/></w:tcPr><w:p><w:r><w:t>Must not disappear</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        await using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new DocxToOfdConverter().ConvertAsync(input, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task Native_CellPageBreak_DegradationIsVisible()
    {
        await using var input = CreateMinimalDocx("<w:tbl><w:tr><w:tc><w:p><w:r><w:t>A</w:t><w:br w:type=\"page\"/><w:t>B</w:t></w:r></w:p></w:tc></w:tr></w:tbl>");
        await using var output = new MemoryStream();
        var result = await new DocxToOfdConverter().ConvertWithResultAsync(input, output);
        Assert.Contains(result.Diagnostics, d => d.Code == "DOCX_CELL_PAGE_BREAK_DEGRADED");
        output.Position = 0; Assert.Equal("AB", string.Concat((await new OfdReader().ReadAsync(output)).Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => e.Text)));
    }
}
