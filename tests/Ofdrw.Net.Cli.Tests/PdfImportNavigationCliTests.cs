using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class PdfImportNavigationCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-pdf-import-" + Guid.NewGuid().ToString("N"));

    public PdfImportNavigationCliTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task DefaultImportCopiesNothing()
    {
        var package = await Import("pdf-to-ofd", "-i", await WriteFixture(), "-o", PathFor("out.ofd"));

        Assert.Null(package.Outlines);
        Assert.Empty(package.Attachments);
    }

    [Fact]
    public async Task CopyBookmarksCopiesTheOutline()
    {
        var package = await Import("pdf-to-ofd", "-i", await WriteFixture(), "-o", PathFor("out.ofd"), "--copy-bookmarks");

        AssertOutline(package);
        Assert.Empty(package.Attachments);
    }

    [Fact]
    public async Task CopyAttachmentsCopiesEmbeddedFiles()
    {
        var package = await Import("pdf-to-ofd", "-i", await WriteFixture(), "-o", PathFor("out.ofd"), "--copy-attachments");

        Assert.Null(package.Outlines);
        AssertAttachment(package);
    }

    [Fact]
    public async Task BothFlagsWorkThroughConvert()
    {
        var package = await Import("convert", await WriteFixture(), PathFor("out.ofd"), "--copy-bookmarks", "--copy-attachments");

        AssertOutline(package);
        AssertAttachment(package);
    }

    [Theory]
    [InlineData("--copy-bookmarks")]
    [InlineData("--copy-attachments")]
    public async Task FlagsAreRejectedForOtherCommands(string flag)
    {
        var input = PathFor("input.ofd");
        var output = PathFor("out.pdf");
        var text = PathFor("input.txt");
        await File.WriteAllTextAsync(text, "Chapter A");
        var (creationCode, creationErrors) = await Run("text-to-ofd", "-i", text, "-o", input);
        Assert.True(creationCode == 0, creationErrors);

        var (code, errors) = await Run("ofd-to-pdf", "-i", input, "-o", output, flag);

        Assert.Equal(1, code);
        Assert.Contains($"{flag} applies only to PDF to OFD conversion.", errors);
        Assert.False(File.Exists(output));
    }

    private static void AssertOutline(OfdDocumentPackage package)
    {
        Assert.NotNull(package.Outlines);
        var outline = Assert.Single(package.Outlines.Items);
        Assert.Equal("Chapter A", outline.Title);
        Assert.NotNull(outline.Actions);
        var action = Assert.IsType<OfdGoToAction>(Assert.Single(outline.Actions.Items));
        Assert.NotNull(action.Destination);
        Assert.Equal(OfdDestinationType.Fit, action.Destination.Type);
        Assert.Same(Assert.Single(package.Pages), action.Destination.Page);
    }

    private static void AssertAttachment(OfdDocumentPackage package)
    {
        var attachment = Assert.Single(package.Attachments);
        Assert.Equal("note.txt", attachment.Name);
        Assert.Equal(Encoding.ASCII.GetBytes("HELLO"), attachment.Data);
    }

    private async Task<OfdDocumentPackage> Import(params string[] args)
    {
        var (code, errors) = await Run(args);
        Assert.True(code == 0, errors);
        await using var stream = File.OpenRead(PathFor("out.ofd"));
        return await new OfdReader().ReadAsync(stream);
    }

    private async Task<string> WriteFixture()
    {
        var input = PathFor("input.pdf");
        await File.WriteAllBytesAsync(input, PdfWithOutlineAndAttachment());
        return input;
    }

    private static byte[] PdfWithOutlineAndAttachment()
    {
        var objects = new[]
        {
            "<< /Type /Catalog /Pages 2 0 R /Outlines 5 0 R /Names << /EmbeddedFiles 7 0 R >> >>",
            "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Resources << >> /Contents 4 0 R >>",
            "<< /Length 0 >>\nstream\nendstream",
            "<< /Type /Outlines /First 6 0 R /Last 6 0 R /Count 1 >>",
            "<< /Title (Chapter A) /Parent 5 0 R /Dest [3 0 R /Fit] >>",
            "<< /Names [(note.txt) 8 0 R] >>",
            "<< /Type /Filespec /F (note.txt) /EF << /F 9 0 R >> >>",
            "<< /Type /EmbeddedFile /Length 5 >>\nstream\nHELLO\nendstream",
        };
        var pdf = new StringBuilder("%PDF-1.4\n");
        var offsets = new List<int>();
        for (var index = 0; index < objects.Length; index++)
        {
            offsets.Add(pdf.Length); // The entire fixture is ASCII, including the xref.
            pdf.Append(FormattableString.Invariant($"{index + 1} 0 obj\n{objects[index]}\nendobj\n"));
        }
        var xref = pdf.Length;
        pdf.Append(FormattableString.Invariant($"xref\n0 {objects.Length + 1}\n0000000000 65535 f \n"));
        foreach (var offset in offsets)
            pdf.Append(FormattableString.Invariant($"{offset:0000000000} 00000 n \n"));
        pdf.Append(FormattableString.Invariant($"trailer\n<< /Size {objects.Length + 1} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n"));
        return Encoding.ASCII.GetBytes(pdf.ToString());
    }

    private static async Task<(int Code, string Errors)> Run(params string[] args)
    {
        var savedError = Console.Error;
        var savedOutput = Console.Out;
        using var errors = new StringWriter();
        using var output = new StringWriter();
        try
        {
            Console.SetError(errors);
            Console.SetOut(output);
            return (await global::Cli.RunAsync(args), errors.ToString());
        }
        finally
        {
            Console.SetError(savedError);
            Console.SetOut(savedOutput);
        }
    }

    private string PathFor(string name) => Path.Combine(_directory, name);

    public void Dispose() => Directory.Delete(_directory, recursive: true);
}
