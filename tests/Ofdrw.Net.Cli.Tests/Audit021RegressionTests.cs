using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class Audit021RegressionTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofdrw-audit021-" + Guid.NewGuid().ToString("N"));
    public Audit021RegressionTests() => Directory.CreateDirectory(directory);
    private string PathFor(string name) => Path.Combine(directory, name);

    [Fact]
    public Task Audit021_NormalControl() => WatermarkAsync("DRAFT");

    [Theory]
    [InlineData("help")]
    [InlineData("--help")]
    [InlineData("-h")]
    public Task Audit021_WatermarkLiteralHelp(string text) => WatermarkAsync(text);

    private async Task WatermarkAsync(string text)
    {
        var input = PathFor("source.ofd"); var output = PathFor("watermarked.ofd");
        await WriteSourceAsync(input);
        Assert.Equal(0, await global::Cli.RunAsync(["watermark", input, output, "--text", text, "--font", "Arial", "--font-size", "8"]));
        Assert.True(File.Exists(output), "A successful watermark must publish the requested output.");
        using var stream = File.OpenRead(output);
        var package = await new OfdReader().ReadAsync(stream);
        Assert.Equal(new[] { "Source 中文 ABC", text }, Assert.Single(package.Pages).Elements.OfType<OfdTextElement>().Select(element => element.Text));
        var artifacts = Environment.GetEnvironmentVariable("OFDRW_PLAN021_ARTIFACTS");
        if (artifacts is not null)
        {
            Directory.CreateDirectory(artifacts);
            File.Copy(input, Path.Combine(artifacts, "source.ofd"), true);
            File.Copy(output, Path.Combine(artifacts, text switch { "--help" => "dash-help.ofd", "-h" => "dash-h.ofd", _ => text + ".ofd" }), true);
        }
        Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("input", "help")]
    [InlineData("output", "help")]
    [InlineData("input", "--help")]
    [InlineData("output", "--help")]
    public async Task Audit021_ImageOrOutputNamedHelp(string slot, string name)
    {
        var input = slot == "input" ? name : "input.png";
        var output = slot == "output" ? name : "output.ofd";
        using (var image = new Image<Rgb24>(20, 10, new Rgb24(210, 40, 20))) image.SaveAsPng(PathFor(input));
        var previous = Directory.GetCurrentDirectory();
        try
        {
            Directory.SetCurrentDirectory(directory);
            Assert.Equal(0, await global::Cli.RunAsync(["image-to-ofd", "--input", input, "--output", output]));
        }
        finally { Directory.SetCurrentDirectory(previous); }
        Assert.True(File.Exists(PathFor(output)), "A help-looking path value must produce an OFD.");
        using var stream = File.OpenRead(PathFor(output));
        Assert.IsType<OfdImageElement>(Assert.Single(Assert.Single((await new OfdReader().ReadAsync(stream)).Pages).Elements));
    }

    [Theory]
    [InlineData("watermark")]
    [InlineData("split")]
    [InlineData("mix")]
    [InlineData("clean-signatures")]
    [InlineData("merge")]
    public async Task Audit021_DocumentPathsNamedHelp(string command)
    {
        await WriteSourceAsync(PathFor("help"));
        var previous = Directory.GetCurrentDirectory();
        try
        {
            Directory.SetCurrentDirectory(directory);
            string[] args = command switch
            {
                "watermark" => [command, "help", "--help", "--text", "DRAFT"],
                "split" => [command, "help", "--help", "--pages", "1"],
                "mix" => [command, "result.ofd", "help", "1"],
                "merge" => [command, "result.ofd", "help", "help"],
                _ => [command, "help", "--help"]
            };
            Assert.Equal(0, await global::Cli.RunAsync(args));
        }
        finally { Directory.SetCurrentDirectory(previous); }
        var output = PathFor(command is "mix" or "merge" ? "result.ofd" : "--help");
        Assert.True(File.Exists(output));
        using var stream = File.OpenRead(output);
        Assert.NotEmpty((await new OfdReader().ReadAsync(stream)).Pages);
    }

    [Theory]
    [InlineData("help")]
    [InlineData("--help")]
    [InlineData("-h")]
    public async Task Audit021_GenericOutputOptionNamedHelp(string name)
    {
        await WriteSourceAsync(PathFor("source.ofd"));
        var previous = Directory.GetCurrentDirectory();
        try
        {
            Directory.SetCurrentDirectory(directory);
            Assert.Equal(0, await global::Cli.RunAsync(["extract-text", "--input", "source.ofd", "--output", name]));
        }
        finally { Directory.SetCurrentDirectory(previous); }
        Assert.Contains("Source 中文 ABC", await File.ReadAllTextAsync(PathFor(name)));
    }

    [Theory]
    [InlineData("help")]
    [InlineData("-h")]
    [InlineData("--help")]
    public async Task Audit021_TopLevelHelp(string token) => Assert.Equal(0, await global::Cli.RunAsync([token]));

    [Theory]
    [InlineData("watermark", "--help")]
    [InlineData("split", "-h")]
    [InlineData("mix", "--help")]
    [InlineData("clean-signatures", "-h")]
    [InlineData("merge", "--help")]
    [InlineData("image-to-ofd", "--help")]
    [InlineData("ofd-to-image", "-h")]
    public async Task Audit021_ActualHelp(string command, string token)
        => Assert.Equal(0, await global::Cli.RunAsync([command, token]));

    [Theory]
    [InlineData("watermark")]
    [InlineData("split")]
    [InlineData("image-to-ofd")]
    [InlineData("ofd-to-image")]
    public async Task Audit021_HelpAfterConsumedValuesDoesNotRequireInput(string command)
        => Assert.Equal(0, await global::Cli.RunAsync([command, PathFor("missing.ofd"), PathFor("output.ofd"), "--help"]));

    [Theory]
    [InlineData("watermark", "--text")]
    [InlineData("split", "--pages")]
    [InlineData("image-to-ofd", "--output")]
    [InlineData("ofd-to-image", "--input")]
    public async Task Audit021_MissingValue(string command, string option)
    {
        await File.WriteAllTextAsync(PathFor("output.ofd"), "preserved");
        Assert.Equal(1, await global::Cli.RunAsync([command, PathFor("missing.ofd"), PathFor("output.ofd"), option]));
        Assert.Equal("preserved", await File.ReadAllTextAsync(PathFor("output.ofd")));
        Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("watermark")]
    [InlineData("split")]
    [InlineData("image-to-ofd")]
    [InlineData("ofd-to-image")]
    public async Task Audit021_UnknownOptionCannotBeHiddenByHelpValue(string command)
        => Assert.Equal(1, await global::Cli.RunAsync([command, PathFor("missing.ofd"), PathFor("output.ofd"), "--unknown", "--help"]));

    [Theory]
    [InlineData("help")]
    [InlineData("--help")]
    [InlineData("-h")]
    public void Audit021_ReplacementControl(string value)
    {
        var parsed = TextReplacementArguments.Parse([PathFor("source.ofd"), PathFor("output.ofd"), "--find", "source", "--replace", value,
            "--trusted-font", new string('a', 64) + ":100"]);
        Assert.Equal(value, parsed.Replacement);
    }

    private static async Task WriteSourceAsync(string path)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 148, HeightMillimeters = 105 };
        page.Elements.Add(new OfdTextElement { Text = "Source 中文 ABC", FontName = "Arial Unicode MS", FontSizeMillimeters = 5, XMillimeters = 15, YMillimeters = 15 });
        package.Pages.Add(page);
        using var stream = File.Create(path); await new OfdPackageWriter().WriteAsync(package, stream);
    }

    public void Dispose() => Directory.Delete(directory, true);
}
