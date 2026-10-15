using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit018RegressionTests
{
    private const string BaseProperties = """
        <w:spacing w:before="120" w:after="240" w:line="360"/>
        <w:ind w:left="400" w:right="600" w:firstLine="100"/>
        <w:pageBreakBefore/><w:keepNext/>
        """;

    private sealed class Source : IDisposable
    {
        internal string Path { get; } = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "audit018-" + Guid.NewGuid() + ".docx");

        internal Source(string properties, string styles = "", string defaults = "", bool cell = false, string? body = null)
        {
            using var document = WordprocessingDocument.Create(Path, WordprocessingDocumentType.Document);
            var main = document.AddMainDocumentPart();
            var paragraph = $"<w:p><w:pPr>{properties}</w:pPr><w:r><w:t>Inheritance AV 中文</w:t></w:r></w:p>";
            if (cell) paragraph = $"<w:tbl><w:tblGrid><w:gridCol w:w=\"7200\"/></w:tblGrid><w:tr><w:tc>{paragraph}</w:tc></w:tr></w:tbl>";
            main.Document = new W.Document($"<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body>{body ?? paragraph}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:right=\"1440\"/></w:sectPr></w:body></w:document>");
            main.AddNewPart<StyleDefinitionsPart>().Styles = new W.Styles($"<w:styles xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docDefaults><w:pPrDefault><w:pPr>{defaults}</w:pPr></w:pPrDefault></w:docDefaults>{styles}</w:styles>");
        }

        internal BuiltInParagraphFormat Read()
        {
            var model = new DocxModelReader(new DocxConversionOptions(), new List<DocxConversionDiagnostic>(), default).Read(Path);
            return Assert.IsType<BuiltInParagraphModel>(Assert.Single(Assert.Single(model.Sections).Blocks)).Format;
        }

        public void Dispose() => File.Delete(Path);
    }

    private static string Style(string id, string properties, string? parent = null) =>
        $"<w:style w:type=\"paragraph\" w:styleId=\"{id}\">{(parent is null ? "" : $"<w:basedOn w:val=\"{parent}\"/>")}<w:pPr>{properties}</w:pPr></w:style>";

    [Fact]
    public void Audit018_Control_InheritsAllAbsentProperties()
    {
        using var source = new Source("<w:pStyle w:val=\"Base\"/>", Style("Base", BaseProperties));
        var format = source.Read();
        Assert.Equal(6, format.SpaceBeforePoints); Assert.Equal(12, format.SpaceAfterPoints);
        Assert.Equal(18, format.LineSpacingPoints); Assert.Equal(20, format.LeftIndentPoints);
        Assert.Equal(30, format.RightIndentPoints); Assert.Equal(5, format.FirstLineIndentPoints);
        Assert.True(format.PageBreakBefore); Assert.True(format.KeepWithNext);
    }

    [Fact]
    public void Audit018_PartialSpacing()
    {
        using var source = new Source("<w:pStyle w:val=\"Base\"/><w:spacing w:before=\"60\"/>", Style("Base", BaseProperties));
        var format = source.Read();
        Assert.Equal(3, format.SpaceBeforePoints); Assert.Equal(12, format.SpaceAfterPoints);
        Assert.Equal(18, format.LineSpacingPoints);
    }

    [Theory]
    [InlineData("left", "right")]
    [InlineData("start", "end")]
    public void Audit018_PartialIndent(string left, string right)
    {
        using var source = new Source("<w:pStyle w:val=\"Base\"/><w:ind w:firstLine=\"200\"/>",
            Style("Base", $"<w:ind w:{left}=\"400\" w:{right}=\"600\"/>"));
        var format = source.Read();
        Assert.Equal(20, format.LeftIndentPoints); Assert.Equal(30, format.RightIndentPoints);
        Assert.Equal(10, format.FirstLineIndentPoints);
    }

    [Theory]
    [InlineData("0")]
    [InlineData("false")]
    [InlineData("off")]
    public void Audit018_ExplicitZeroAndFalse(string value)
    {
        using var source = new Source($"<w:pStyle w:val=\"Base\"/><w:spacing w:before=\"0\" w:after=\"0\" w:line=\"0\"/><w:ind w:left=\"0\" w:right=\"0\" w:firstLine=\"0\"/><w:pageBreakBefore w:val=\"{value}\"/><w:keepNext w:val=\"{value}\"/>", Style("Base", BaseProperties));
        var format = source.Read();
        Assert.Equal(0, format.SpaceBeforePoints); Assert.Equal(0, format.SpaceAfterPoints);
        Assert.Equal(0, format.LineSpacingPoints); Assert.Equal(0, format.LeftIndentPoints);
        Assert.Equal(0, format.RightIndentPoints); Assert.Equal(0, format.FirstLineIndentPoints);
        Assert.False(format.PageBreakBefore); Assert.False(format.KeepWithNext);
    }

    [Fact]
    public void Audit018_StyleChain()
    {
        using var source = new Source("<w:pStyle w:val=\"Leaf\"/><w:spacing w:before=\"0\"/><w:ind w:firstLine=\"0\"/>",
            Style("Base", "<w:spacing w:before=\"200\"/><w:ind w:left=\"800\"/>") +
            Style("Middle", "<w:spacing w:after=\"400\"/><w:ind w:right=\"1000\"/><w:pageBreakBefore w:val=\"false\"/>", "Base") +
            Style("Leaf", "<w:spacing w:line=\"600\"/><w:ind w:hanging=\"300\"/><w:keepNext w:val=\"0\"/>", "Middle"), BaseProperties);
        var format = source.Read();
        Assert.Equal(0, format.SpaceBeforePoints); Assert.Equal(20, format.SpaceAfterPoints);
        Assert.Equal(30, format.LineSpacingPoints); Assert.Equal(40, format.LeftIndentPoints);
        Assert.Equal(50, format.RightIndentPoints); Assert.Equal(0, format.FirstLineIndentPoints);
        Assert.False(format.PageBreakBefore); Assert.False(format.KeepWithNext);
    }

    [Theory]
    [InlineData("<w:ind/>", 5)]
    [InlineData("<w:ind w:hanging=\"200\"/>", -10)]
    [InlineData("<w:ind w:hanging=\"0\"/>", 0)]
    public void Audit018_FirstLinePresenceAndHanging(string direct, double expected)
    {
        using var source = new Source("<w:pStyle w:val=\"Base\"/><w:spacing/>" + direct, Style("Base", BaseProperties));
        var format = source.Read();
        Assert.Equal(expected, format.FirstLineIndentPoints);
        Assert.Equal(6, format.SpaceBeforePoints); Assert.Equal(12, format.SpaceAfterPoints); Assert.Equal(18, format.LineSpacingPoints);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit018_CellFalseBreak(bool explicitNative)
    {
        using var source = new Source("<w:pStyle w:val=\"Base\"/><w:pageBreakBefore w:val=\"false\"/>", Style("Base", "<w:pageBreakBefore/>"), cell: true);
        var options = new DocxConversionOptions { UnsupportedFeatureBehavior = UnsupportedDocxFeatureBehavior.Throw };
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        using var input = File.OpenRead(source.Path); using var output = new MemoryStream();
        var result = await new DocxToOfdConverter(options).ConvertWithResultAsync(input, output);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == "DOCX_CELL_PAGE_BREAK_DEGRADED");
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Single(package.Pages);
        Assert.Equal("Inheritance AV 中文", string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.Text)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit018_PublicBody_PrecedencePreservesGeometryTextAndPageSelection(bool explicitNative)
    {
        const string body = """
            <w:p><w:pPr><w:pStyle w:val="Base"/><w:spacing w:before="0"/><w:pageBreakBefore w:val="false"/><w:keepNext w:val="false"/></w:pPr>
              <w:r><w:t xml:space="preserve">继承 Inheritance AV </w:t></w:r><w:r><w:rPr><w:b/><w:color w:val="C02020"/></w:rPr><w:t xml:space="preserve">粗体 Bold </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>斜体 Italic</w:t></w:r></w:p>
            <w:p><w:r><w:t>继承段后距 After spacing</w:t></w:r></w:p>
            <w:tbl><w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8"/><w:left w:val="single" w:sz="8"/><w:bottom w:val="single" w:sz="8"/><w:right w:val="single" w:sz="8"/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="7200"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:shd w:fill="DCEEFF"/></w:tcPr><w:p><w:pPr><w:pStyle w:val="Base"/><w:pageBreakBefore w:val="0"/><w:keepNext w:val="0"/><w:jc w:val="center"/></w:pPr><w:r><w:t>单元格 Cell false break</w:t></w:r></w:p></w:tc></w:tr></w:tbl>
            <w:p><w:pPr><w:pageBreakBefore/><w:jc w:val="right"/></w:pPr><w:r><w:t>第二页 Page two</w:t></w:r></w:p>
            """;
        using var source = new Source("", Style("Base", BaseProperties), body: body);
        var options = new DocxConversionOptions { UnsupportedFeatureBehavior = UnsupportedDocxFeatureBehavior.Throw };
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        var converter = new DocxToOfdConverter(options);
        using var input = File.OpenRead(source.Path); using var output = new MemoryStream();
        await converter.ConvertAsync(input, output); output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, package.Pages.Count);
        var first = package.Pages[0].Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal("继承 Inheritance AV 粗体 Bold 斜体 Italic继承段后距 After spacing单元格 Cell false break", string.Concat(first.Select(text => text.Text)));
        Assert.Equal((72 + 20 + 5) * 25.4 / 72, first[0].XMillimeters, 3);
        Assert.Equal(72 * 25.4 / 72, first[0].YMillimeters, 3);
        Assert.Equal("第二页 Page two", string.Concat(package.Pages[1].Elements.OfType<OfdTextElement>().Select(text => text.Text)));
        input.Position = 0; using var selected = new MemoryStream();
        // Converter page positions are zero-based; the CLI translates human page numbers.
        await converter.ConvertAsync(input, selected, new[] { 1 }); selected.Position = 0;
        var selectedPackage = await new OfdReader().ReadAsync(selected);
        Assert.Equal("第二页 Page two", string.Concat(Assert.Single(selectedPackage.Pages).Elements.OfType<OfdTextElement>().Select(text => text.Text)));
    }
}
