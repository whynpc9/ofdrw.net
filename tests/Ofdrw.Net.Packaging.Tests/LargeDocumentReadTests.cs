using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class LargeDocumentReadTests
{
    [Theory]
    [InlineData(1200)]
    [InlineData(4000)]
    // 20,000 paragraphs exceed the writer's preflight output estimate (MaxMaterializedBytes), not the reader.
    [InlineData(10000)]
    public async Task Reader_ShouldReadLargeFlowDocument(int paragraphs)
    {
        var (stream, rendered) = await WriteFlow(paragraphs);
        using (stream)
        {
            var read = await new OfdReader().ReadAsync(stream);

            Assert.Equal(rendered.Pages.Count, read.Pages.Count);
            Assert.Equal(CountText(rendered), CountText(read));
        }
    }

    [Fact]
    public async Task TextExtractor_ShouldExtractLargeDocument()
    {
        var (stream, _) = await WriteFlow(4000);
        using (stream)
        {
            var read = await new OfdReader().ReadAsync(stream);
            var pages = new OfdTextExtractor().ExtractPages(read, includeTemplates: true);

            Assert.Equal(read.Pages.Count, pages.Count);
            Assert.Contains("Paragraph 3999:", pages[pages.Count - 1]);
        }
    }

    // Guard: listing reads only OFD.xml, so it passed before the budget fix too.
    [Fact]
    public async Task ListDocuments_ShouldListLargeDocument()
    {
        var (stream, _) = await WriteFlow(4000);
        using (stream)
        {
            Assert.Single(await new OfdReader().ListDocumentsAsync(stream));
        }
    }

    internal static async Task<(MemoryStream Stream, OfdDocumentPackage Rendered)> WriteFlow(int paragraphs)
    {
        var document = new FlowDocument();
        for (var i = 0; i < paragraphs; i++)
        {
            document.Blocks.Add(new Paragraph($"Paragraph {i}: The quick brown fox jumps over the lazy dog, 中文段落测试内容第{i}段。"));
        }

        var rendered = document.Render();
        var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(rendered, stream);
        stream.Position = 0;
        return (stream, rendered);
    }

    private static int CountText(OfdDocumentPackage package) =>
        package.Pages.Sum(page => page.Elements.OfType<OfdTextElement>().Count());
}
