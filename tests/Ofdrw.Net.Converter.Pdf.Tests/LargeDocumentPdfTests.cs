using System.IO.Compression;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class LargeDocumentPdfTests
{
    [Fact]
    public async Task OfdToPdf_ShouldConvertLargeFlowDocument()
    {
        var document = new FlowDocument();
        for (var i = 0; i < 2000; i++)
        {
            document.Blocks.Add(new Paragraph($"Paragraph {i}: The quick brown fox jumps over the lazy dog, 中文段落测试内容第{i}段。"));
        }

        var rendered = document.Render();
        // Deterministic CJK coverage: embed a subset of the pinned CI font instead of relying on host fonts.
        var characters = rendered.Pages.SelectMany(page => page.Elements.OfType<OfdTextElement>())
            .SelectMany(text => text.Text).Distinct().Select(c => (int)c).ToArray();
        var font = CjkFont(characters);
        foreach (var resource in rendered.Fonts) resource.Data = font;
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(rendered, ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream();

        await new OfdToPdfConverter().ConvertAsync(ofd, pdf);

        using var parsed = UglyToad.PdfPig.PdfDocument.Open(pdf.ToArray());
        Assert.Equal(rendered.Pages.Count, parsed.NumberOfPages);

        // Optional visual-review output; never written in normal runs.
        var visualDirectory = Environment.GetEnvironmentVariable("OFDRW_LARGE_DOCUMENT_VISUAL_DIR");
        if (!string.IsNullOrEmpty(visualDirectory))
        {
            Directory.CreateDirectory(visualDirectory);
            await File.WriteAllBytesAsync(Path.Combine(visualDirectory, "large.ofd"), ofd.ToArray());
            await File.WriteAllBytesAsync(Path.Combine(visualDirectory, "large.pdf"), pdf.ToArray());
        }
    }

    private static byte[] CjkFont(IEnumerable<int> characters)
    {
        using var archive = ZipFile.OpenRead(Path.Combine(AppContext.BaseDirectory, "fonts", "trusted-fonts.zip"));
        using var entry = archive.GetEntry("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf")!.Open();
        using var bytes = new MemoryStream();
        entry.CopyTo(bytes);
        return TrueTypeSubset.Create(new OpenTypeFace(bytes.ToArray()), characters,
            CancellationToken.None, 1_000_000, out _, out _);
    }
}
