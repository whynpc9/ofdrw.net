using System.IO.Compression;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class NativeEmbeddedBoldArgumentsTests : IDisposable
{
    private const string Option = "--native-embedded-bold-behavior";
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-bold-" + Guid.NewGuid().ToString("N"));

    public NativeEmbeddedBoldArgumentsTests() => Directory.CreateDirectory(directory);

    [Theory]
    [InlineData("docx-to-ofd", false, "throw")]
    [InlineData("docx-to-ofd", true, " THROW ")]
    [InlineData("convert", false, "throw")]
    [InlineData("convert", true, "throw")]
    public async Task StrictPolicyReachesNativeConverterAndPreservesOutput(string command, bool explicitNative, string policy)
    {
        var (input, fonts) = CreateRegularOnlyInput();
        var output = Path.Combine(directory, "output.ofd");
        await File.WriteAllTextAsync(output, "existing output");
        var args = Arguments(command, input, output, fonts, explicitNative);
        args.AddRange([Option, policy]);
        var (code, errors) = await RunAsync(args.ToArray());
        Assert.Equal(1, code);
        Assert.Contains("actual bold face", errors);
        Assert.DoesNotContain("Unknown option", errors);
        Assert.Equal("existing output", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData(false, null)]
    [InlineData(true, null)]
    [InlineData(false, "text-stroke")]
    [InlineData(true, " TEXT-STROKE ")]
    public async Task DefaultAndExplicitTextStrokeKeepCompatibilityBehavior(bool explicitNative, string? policy)
    {
        var (input, fonts) = CreateRegularOnlyInput();
        var output = Path.Combine(directory, "output.ofd");
        var args = Arguments("docx-to-ofd", input, output, fonts, explicitNative);
        if (policy is not null) args.AddRange([Option, policy]);
        var (code, errors) = await RunAsync(args.ToArray());
        Assert.Equal(0, code);
        Assert.Contains("DOCX_EMBEDDED_BOLD_SIMULATED", errors);
        using var stream = File.OpenRead(output);
        var package = await new OfdReader().ReadAsync(stream);
        var text = Assert.Single(Assert.Single(package.Pages).Elements.OfType<OfdTextElement>());
        Assert.Equal("A", text.Text);
        Assert.Equal(700, text.Weight);
        Assert.Contains("Stroke=\"true\"", text.SourceXml);
        Assert.False(Assert.Single(package.Fonts).Bold);
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("TextStroke")]
    [InlineData("0")]
    [InlineData("")]
    [InlineData("--help")]
    public async Task InvalidPolicyFailsBeforeInputOrOutput(string value)
    {
        var (code, errors) = await RunAsync(["docx-to-ofd", "missing.docx", "unused.ofd", Option, value]);
        Assert.Equal(1, code);
        Assert.Contains("Invalid " + Option + " value. Use text-stroke or throw.", errors);
    }

    [Fact]
    public async Task MissingPolicyValueIsReported()
    {
        var (code, errors) = await RunAsync(["docx-to-ofd", Option]);
        Assert.Equal(1, code);
        Assert.Contains("Missing value for " + Option, errors);
    }

    [Fact]
    public async Task HelpDocumentsPolicyAndScope()
    {
        var saved = Console.Out;
        using var output = new StringWriter();
        try
        {
            Console.SetOut(output);
            Assert.Equal(0, await global::Cli.RunAsync(["--help"]));
        }
        finally { Console.SetOut(saved); }
        Assert.Contains(Option, output.ToString());
        Assert.Contains("text-stroke (default) or throw", output.ToString());
        Assert.Contains("Native embedded bold", output.ToString());
    }

    private (string Input, string Fonts) CreateRegularOnlyInput()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        var fonts = Path.Combine(directory, "fonts");
        Directory.CreateDirectory(fonts);
        // Original deterministic metric fixture: enough to prove CLI routing, not visual fidelity.
        File.Copy(Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"), Path.Combine(fonts, "regular.ttf"));
        var input = Path.Combine(directory, "input.docx");
        using var archive = ZipFile.Open(input, ZipArchiveMode.Create);
        Write("[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
            """);
        Write("_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
            """);
        Write("word/document.xml", """
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="Ofdrw Test Face" w:hAnsi="Ofdrw Test Face"/><w:b/></w:rPr><w:t>A</w:t></w:r></w:p></w:body></w:document>
            """);
        return (input, fonts);
        void Write(string name, string contents) { using var writer = new StreamWriter(archive.CreateEntry(name).Open()); writer.Write(contents); }
    }

    private static List<string> Arguments(string command, string input, string output, string fonts, bool explicitNative)
    {
        var args = new List<string> { command, input, output, "--font-directory", fonts };
        if (explicitNative) args.AddRange(["--docx-ofd-mode", "native"]);
        return args;
    }

    private static async Task<(int Code, string Errors)> RunAsync(string[] args)
    {
        var saved = Console.Error;
        using var errors = new StringWriter();
        try { Console.SetError(errors); return (await global::Cli.RunAsync(args), errors.ToString()); }
        finally { Console.SetError(saved); }
    }

    public void Dispose() => Directory.Delete(directory, true);
}
