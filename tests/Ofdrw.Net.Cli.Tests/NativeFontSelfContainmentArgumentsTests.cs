using System.IO.Compression;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class NativeFontSelfContainmentArgumentsTests : IDisposable
{
    private const string Option = "--native-font-self-containment";
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-fonts-" + Guid.NewGuid().ToString("N"));

    public NativeFontSelfContainmentArgumentsTests() => Directory.CreateDirectory(directory);

    [Theory]
    [InlineData("docx-to-ofd", false, "require-embedded")]
    [InlineData("docx-to-ofd", true, " REQUIRE-EMBEDDED ")]
    [InlineData("convert", false, "require-embedded")]
    public async Task RequireEmbeddedReachesNativeConverterAndPreservesOutputOnFailure(string command, bool explicitNative, string policy)
    {
        var (input, fonts) = CreateInput("SimSun");
        var output = Path.Combine(directory, "output.ofd");
        await File.WriteAllTextAsync(output, "existing output");
        var args = Arguments(command, input, output, fonts, explicitNative);
        args.AddRange([Option, policy]);
        var (code, errors) = await RunAsync(args.ToArray());
        Assert.Equal(1, code);
        Assert.Contains("DOCX_FONT_NOT_SELF_CONTAINED", errors);
        Assert.Contains("'SimSun'", errors);
        Assert.Equal("existing output", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task RequireEmbeddedEmbedsAnAvailableFamily()
    {
        var (input, fonts) = CreateInput("Ofdrw Test Face");
        var output = Path.Combine(directory, "output.ofd");
        var args = Arguments("docx-to-ofd", input, output, fonts, explicitNative: true);
        args.AddRange([Option, "require-embedded"]);
        var (code, errors) = await RunAsync(args.ToArray());
        Assert.Equal(0, code);
        Assert.Contains("DOCX_FONT_LICENSE_NOT_VERIFIED", errors);
        using var stream = File.OpenRead(output);
        var package = await new OfdReader().ReadAsync(stream);
        Assert.NotEmpty(Assert.Single(package.Fonts).Data);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("compatibility")]
    [InlineData(" Compatibility ")]
    public async Task DefaultAndExplicitCompatibilityKeepNameOnlySystemFonts(string? policy)
    {
        var (input, fonts) = CreateInput("SimSun");
        var output = Path.Combine(directory, "output.ofd");
        var args = Arguments("docx-to-ofd", input, output, fonts, explicitNative: false);
        if (policy is not null) args.AddRange([Option, policy]);
        var (code, errors) = await RunAsync(args.ToArray());
        Assert.Equal(0, code);
        Assert.DoesNotContain("DOCX_FONT_LICENSE_NOT_VERIFIED", errors);
        using var stream = File.OpenRead(output);
        var package = await new OfdReader().ReadAsync(stream);
        Assert.Empty(Assert.Single(package.Fonts).Data);
    }

    [Fact]
    public async Task DualLayerRejectsRequireEmbedded()
    {
        var (input, fonts) = CreateInput("Ofdrw Test Face");
        var output = Path.Combine(directory, "output.ofd");
        var (code, errors) = await RunAsync(["docx-to-ofd", input, output, "--font-directory", fonts,
            "--docx-ofd-mode", "dual-layer", Option, "require-embedded"]);
        Assert.Equal(1, code);
        Assert.Contains("applies only to Native DOCX-to-OFD output", errors);
        Assert.False(File.Exists(output));
    }

    [Theory]
    [InlineData("invalid")]
    [InlineData("RequireEmbedded")]
    [InlineData("1")]
    [InlineData("")]
    public async Task InvalidPolicyFailsBeforeInputOrOutput(string value)
    {
        var (code, errors) = await RunAsync(["docx-to-ofd", "missing.docx", "unused.ofd", Option, value]);
        Assert.Equal(1, code);
        Assert.Contains("Invalid " + Option + " value. Use compatibility or require-embedded.", errors);
    }

    [Fact]
    public async Task MissingPolicyValueIsReported()
    {
        var (code, errors) = await RunAsync(["docx-to-ofd", Option]);
        Assert.Equal(1, code);
        Assert.Contains("Missing value for " + Option, errors);
    }

    [Fact]
    public async Task HelpDocumentsPolicyAndLicensingCaveat()
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
        Assert.Contains("compatibility (default) or require-embedded", output.ToString());
        Assert.Contains("not a licence grant", output.ToString());
    }

    private (string Input, string Fonts) CreateInput(string family)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        var fonts = Path.Combine(directory, "fonts");
        Directory.CreateDirectory(fonts);
        // Deterministic open fixture whose own family is "Ofdrw Test Face".
        File.Copy(Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"), Path.Combine(fonts, "regular.ttf"), true);
        var input = Path.Combine(directory, "input.docx");
        if (File.Exists(input)) File.Delete(input);
        using var archive = ZipFile.Open(input, ZipArchiveMode.Create);
        Write("[Content_Types].xml", """
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>
            """);
        Write("_rels/.rels", """
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>
            """);
        Write("word/document.xml", $"""
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:rFonts w:ascii="{family}" w:hAnsi="{family}" w:eastAsia="{family}"/></w:rPr><w:t>A</w:t></w:r></w:p></w:body></w:document>
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
