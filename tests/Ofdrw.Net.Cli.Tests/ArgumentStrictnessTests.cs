namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class ArgumentStrictnessTests : IAsyncLifetime, IDisposable
{
    private readonly string _directory = Directory.CreateTempSubdirectory("ofdrw-cli-arguments-").FullName;

    private string PathFor(string name) => Path.Combine(_directory, name);
    private string Input => PathFor("a.ofd");
    private string Output => PathFor("z.pdf");

    public async Task InitializeAsync()
    {
        foreach (var name in new[] { "a", "b" })
        {
            var text = PathFor(name + ".txt");
            await File.WriteAllTextAsync(text, "Source " + name);
            var (code, errors) = await Run("text-to-ofd", text, PathFor(name + ".ofd"));
            Assert.True(code == 0, errors);
        }
    }

    [Fact]
    public async Task MergeRejectsUnknownOption()
    {
        var output = PathFor("out.ofd");
        await Reject("Unknown option: -o", ["merge", "-o", output, Input, PathFor("b.ofd")],
            output, PathFor("-o"), Path.GetFullPath("-o"));
    }

    [Fact]
    public Task ExtraPositionalIsRejectedWithInputOption() =>
        Reject("Too many positional arguments.", ["ofd-to-pdf", "-i", Input, PathFor("x.pdf"), PathFor("y.pdf")],
            PathFor("x.pdf"), PathFor("y.pdf"));

    [Fact]
    public Task ExtraPositionalIsRejectedWithBothOptions() =>
        Reject("Too many positional arguments.", ["ofd-to-pdf", "-i", Input, "-o", Output, PathFor("stray.pdf")],
            Output, PathFor("stray.pdf"));

    [Fact]
    public Task DuplicateInputIsRejected() =>
        Reject("Duplicate --input.", ["extract-text", "-i", Input, "--input", PathFor("b.ofd")]);

    [Fact]
    public Task DuplicateOutputIsRejected() =>
        Reject("Duplicate --output.", ["ofd-to-pdf", "-i", Input, "-o", Output, "-o", PathFor("y.pdf")],
            Output, PathFor("y.pdf"));

    [Fact]
    public Task DuplicatePagesIsRejected() =>
        Reject("Duplicate --pages.", ["ofd-to-pdf", "-i", Input, "-o", Output, "-p", "1", "--pages", "1"], Output);

    [Theory]
    [InlineData("extract-text")]
    [InlineData("verify-signatures")]
    public Task PagesDoesNotApplyToReadOnlyCommands(string command) =>
        Reject($"--pages does not apply to {command}.", [command, Input, "--pages", "1"]);

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public Task VerifySignaturesTakesNoOutput(bool useOption)
    {
        var output = PathFor("out.txt");
        return Reject("verify-signatures does not take an output path.",
            useOption ? ["verify-signatures", Input, "-o", output] : ["verify-signatures", Input, output], output);
    }

    [Fact]
    public Task IncludeTemplatesAppliesOnlyToExtractText() =>
        Reject("--include-templates applies only to extract-text.",
            ["ofd-to-pdf", "-i", Input, "-o", Output, "--include-templates"], Output);

    [Theory]
    [InlineData("--docx-engine", "word")]
    [InlineData("--docx-ofd-mode", "native")]
    [InlineData("--native-embedded-bold-behavior", "throw")]
    [InlineData("--native-font-self-containment", "require-embedded")]
    [InlineData("--libreoffice", "/x")]
    [InlineData("--font-directory", "/x")]
    public Task DocxOptionsApplyOnlyToDocxInput(string option, string value) =>
        Reject($"{option} applies only to DOCX conversion.",
            ["ofd-to-pdf", "-i", Input, "-o", Output, option, value], Output);

    [Fact]
    public Task InferredOfdConversionRejectsDocxOptions() =>
        Reject("--docx-engine applies only to DOCX conversion.",
            ["convert", Input, Output, "--docx-engine", "auto"], Output);

    [Fact]
    public Task SvgConversionRejectsDocxOptions() =>
        Reject("--font-directory applies only to DOCX conversion.",
            ["ofd-to-svg", Input, PathFor("z.svg"), "--font-directory", "/x"], PathFor("z.svg"));

    [Fact]
    public Task SignedPageNumberIsRejected() =>
        Reject("Invalid page number: +1", ["ofd-to-pdf", "-i", Input, "-o", Output, "--pages", "+1"], Output);

    [Fact]
    public async Task OutputOptionWithPositionalInputStillWorks()
    {
        var (code, errors) = await Run("ofd-to-pdf", "-o", Output, Input);
        Assert.True(code == 0, errors);
        Assert.True(File.Exists(Output));
    }

    [Fact]
    public async Task InputOptionWithPositionalOutputWorks()
    {
        var (code, errors) = await Run("ofd-to-pdf", "-i", Input, Output);
        Assert.True(code == 0, errors);
        Assert.True(File.Exists(Output));
    }

    private static async Task Reject(string error, string[] args, params string[] outputs)
    {
        var (code, errors) = await Run(args);
        Assert.Equal(1, code);
        Assert.Contains(error, errors);
        Assert.All(outputs, path => Assert.False(File.Exists(path), $"Unexpected output: {path}"));
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
            var code = await global::Cli.RunAsync(args);
            return (code, errors.ToString());
        }
        finally
        {
            Console.SetError(savedError);
            Console.SetOut(savedOutput);
        }
    }

    public Task DisposeAsync() => Task.CompletedTask;
    public void Dispose() => Directory.Delete(_directory, true);
}
