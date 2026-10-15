using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit005RegressionTests
{
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit005_NormalMergedTable(bool explicitNative)
    {
        using var input = CreateTable(true, 2, 1);
        using var output = new MemoryStream();
        var options = new DocxConversionOptions { MaxTableColumnCount = 3 };
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        await new DocxToOfdConverter(options).ConvertAsync(input, output);
        output.Position = 0;
        var package = await new OfdReader().ReadAsync(output);
        Assert.Equal("Cell1Cell2", string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(t => t.Text)));
        var text = package.Pages.Single().Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(text[0].YMillimeters, text[1].YMillimeters);
        Assert.True(text[1].XMillimeters > text[0].XMillimeters);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit005_AbsentGridOverLimit(bool builtInPdf)
    {
        using var input = CreateTable(false, 3);
        using var output = CreateExistingOutput();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => ConvertAsync(input, output, 2, builtInPdf));
        Assert.Contains("row 1, cell 1", error.Message);
        Assert.Contains("column count limit of 2", error.Message);
        Assert.DoesNotContain("Cell1", error.Message);
        AssertExistingOutput(output);
    }

    [Theory]
    [InlineData(false, false, 3)]
    [InlineData(false, false, 4)]
    [InlineData(true, false, 3)]
    [InlineData(true, false, 4)]
    [InlineData(false, true, 3)]
    [InlineData(false, true, 4)]
    [InlineData(true, true, 3)]
    [InlineData(true, true, 4)]
    public async Task Audit005_BoundaryColumns(bool explicitGrid, bool builtInPdf, int columns)
    {
        using var input = CreateTable(explicitGrid, columns);
        using var output = new MemoryStream();
        if (columns == 3)
        {
            await ConvertAsync(input, output, 3, builtInPdf);
            Assert.True(output.Length > 0);
            if (builtInPdf)
            {
                using var pdf = UglyToad.PdfPig.PdfDocument.Open(output.ToArray());
                Assert.Equal("Cell1", pdf.GetPage(1).Text);
            }
            else
            {
                output.Position = 0;
                var package = await new OfdReader().ReadAsync(output);
                Assert.Equal("Cell1", string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(t => t.Text)));
            }
        }
        else
        {
            var error = await Assert.ThrowsAsync<InvalidDataException>(() => ConvertAsync(input, output, 3, builtInPdf));
            Assert.Contains(explicitGrid ? "grid column 4" : "row 1, cell 1", error.Message);
            Assert.Contains("column count limit of 3", error.Message);
            Assert.Equal(0, output.Length);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit005_SpanSumOverflow(bool builtInPdf)
    {
        // No grid is materialized: the second cell must reject an Int32-overflowing sum first.
        using var input = CreateTable(false, int.MaxValue, 1);
        using var output = CreateExistingOutput();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => ConvertAsync(input, output, int.MaxValue, builtInPdf));
        Assert.Contains("row 1, cell 2", error.Message);
        Assert.Contains("column count limit", error.Message);
        AssertExistingOutput(output);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit005_ExplicitGridAndCancellation(bool builtInPdf)
    {
        using var input = CreateTable(true, 4);
        using var output = CreateExistingOutput();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => ConvertAsync(input, output, 3, builtInPdf));
        Assert.Contains("grid column 4", error.Message);
        AssertExistingOutput(output);
        input.Position = 0;
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => ConvertAsync(input, output, 3, builtInPdf, cancellation.Token));
        AssertExistingOutput(output);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void Audit005_NonPositiveLimitRejectedAtConstruction(int limit)
    {
        var options = new DocxConversionOptions { MaxTableColumnCount = limit };
        Assert.Throws<ArgumentOutOfRangeException>(() => new DocxToPdfConverter(options));
        Assert.Throws<ArgumentOutOfRangeException>(() => new DocxToOfdConverter(options));
    }

    private static Task ConvertAsync(Stream input, Stream output, int limit, bool builtInPdf,
        CancellationToken cancellationToken = default)
    {
        var options = new DocxConversionOptions { MaxTableColumnCount = limit, Engine = DocxConversionEngine.BuiltIn };
        return builtInPdf
            ? new DocxToPdfConverter(options).ConvertAsync(input, output, cancellationToken)
            : new DocxToOfdConverter(options).ConvertAsync(input, output, cancellationToken: cancellationToken);
    }

    private static MemoryStream CreateExistingOutput()
    {
        var stream = new MemoryStream();
        stream.Write(new byte[] { 17, 23, 41, 59 });
        stream.Position = 2;
        return stream;
    }

    private static void AssertExistingOutput(MemoryStream output)
    {
        Assert.Equal(new byte[] { 17, 23, 41, 59 }, output.ToArray());
        Assert.Equal(2, output.Position);
    }

    private static MemoryStream CreateTable(bool explicitGrid, params int[] spans)
    {
        var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document, true))
        {
            var part = document.AddMainDocumentPart();
            var table = new Table();
            if (explicitGrid)
                table.Append(new TableGrid(Enumerable.Range(0, spans.Sum()).Select(_ => new GridColumn { Width = "1000" })));
            table.Append(new TableRow(spans.Select((span, index) => new TableCell(
                new TableCellProperties(new GridSpan { Val = span }),
                new Paragraph(new Run(new Text("Cell" + (index + 1))))))));
            part.Document = new Document(new Body(table));
            part.Document.Save();
        }
        stream.Position = 0;
        return stream;
    }
}
