using System.IO.Compression;
using System.Text;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class Audit011DocxRegressionTests
{
    private static DocxConversionOptions Options(long work, bool explicitNative)
    {
        var options = new DocxConversionOptions { UseInstalledMicrosoftOfficeFonts = false };
        if (explicitNative) options.OfdMode = DocxToOfdMode.Native;
        options.FontEmbedding.MaximumCmapParseWork = work;
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        options.FontDirectories.Add(Path.Combine(root.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts"));
        options.FontFallbackFamilies.Clear(); options.FontFallbackFamilies.Add("Noto Sans");
        return options;
    }

    private static MemoryStream Source()
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
        {
            void Add(string name, string xml)
            { using var writer = new StreamWriter(zip.CreateEntry(name).Open(), new UTF8Encoding(false)); writer.Write(xml); }
            Add("[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/><Override PartName=\"/word/document.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml\"/></Types>");
            Add("_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"r1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"word/document.xml\"/></Relationships>");
            Add("word/document.xml", "<w:document xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii=\"Noto Sans\" w:hAnsi=\"Noto Sans\"/></w:rPr><w:t>AV office é</w:t></w:r></w:p></w:body></w:document>");
        }
        stream.Position = 0; return stream;
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit011_PublicNativeAndDefault_RespectTinyCmapBudgetBeforeOutput(bool explicitNative)
    {
        var options = Options(1, explicitNative);
        using var source = Source(); using var output = new MemoryStream();
        var failure = await Assert.ThrowsAsync<InvalidDataException>(() => new DocxToOfdConverter(options).ConvertAsync(source, output));
        Assert.Contains("MaximumCmapParseWork", failure.Message); Assert.Equal(0, output.Length);
        Assert.Equal(1, options.FontEmbedding.MaximumCmapParseWork);
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Audit011_PublicNativeAndDefault_NondefaultCmapBudgetPreservesEmbeddedText(bool explicitNative)
    {
        var options = Options(100_003, explicitNative);
        using var source = Source(); using var output = new MemoryStream();
        await new DocxToOfdConverter(options).ConvertAsync(source, output);
        output.Position = 0; var package = await new OfdReader().ReadAsync(output);
        Assert.Contains(package.Fonts, font => font.Data.Length > 0);
        Assert.Equal("AV office é", string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().Select(text => text.Text)));
        Assert.Equal(100_003, options.FontEmbedding.MaximumCmapParseWork);
    }

    [Fact]
    public void Audit011_RendererCopiesNondefaultCmapLimitIntoIndependentPackageOptions()
    {
        var options = Options(100_003, true);
        var renderer = new BuiltInOfdRenderer(options, new List<DocxConversionDiagnostic>(), default);
        var package = renderer.Render(new BuiltInDocumentModel());
        Assert.NotSame(options.FontEmbedding, package.Options.FontEmbedding);
        Assert.Equal(100_003, package.Options.FontEmbedding.MaximumCmapParseWork);
        package.Options.FontEmbedding.MaximumCmapParseWork = 17;
        Assert.Equal(100_003, options.FontEmbedding.MaximumCmapParseWork);
    }
}
