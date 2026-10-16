using System.Text;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class PlainTextCommandTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-text-" + Guid.NewGuid().ToString("N"));
    public PlainTextCommandTests() => Directory.CreateDirectory(_directory);

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task Utf8ToNativeOfdUsingBothPathForms(bool named)
    {
        var input = Path.Combine(_directory, "notice.data"); var output = Path.Combine(_directory, "notice.ofd");
        await File.WriteAllTextAsync(input, "\uFEFF中英 Alpha\r\n\r\n尾行\n", new UTF8Encoding(false));
        string[] paths = named ? ["--input", input, "--output", output] : [input, output];
        Assert.Equal(0, await global::Cli.RunAsync(["text-to-ofd", .. paths, "--margin-left", "0", "--font-family", "SimSun"]));
        using var stream = File.OpenRead(output); var package = await new OfdReader().ReadAsync(stream);
        Assert.Equal("中英 Alpha尾行", string.Concat(package.Pages.SelectMany(p => p.Elements).OfType<OfdTextElement>().Select(e => e.Text)));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("--pages", "1")] [InlineData("--font-directory", ".")]
    [InlineData("--docx-ofd-mode", "native")] [InlineData("--font-size-mm", "NaN")]
    [InlineData("--font-size-mm", "Infinity")] [InlineData("--font-size-mm", "4,2")]
    [InlineData("--page-width", "210")] [InlineData("--margin-top", "-1")]
    [InlineData("--max-pages", "0")] [InlineData("--max-characters", "2147483648")]
    public async Task InvalidOptionsFailBeforeInputOrStaging(string option, string value)
    {
        var output = Path.Combine(_directory, "out.ofd"); await File.WriteAllTextAsync(output, "original");
        using var errors = new StringWriter(); var old = Console.Error;
        try
        {
            Console.SetError(errors);
            Assert.Equal(1, await global::Cli.RunAsync(["text-to-ofd", Path.Combine(_directory, "missing.txt"), output, option, value]));
        }
        finally { Console.SetError(old); }
        Assert.DoesNotContain("Could not find", errors.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.Equal("original", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("duplicate")] [InlineData("mixed")] [InlineData("extra")] [InlineData("extension")] [InlineData("same")]
    public async Task PathAndSyntaxFailuresPreserveInput(string kind)
    {
        var input = Path.Combine(_directory, "in.ofd"); var output = Path.Combine(_directory, "out.ofd");
        await File.WriteAllTextAsync(input, "original");
        string[] arguments = kind switch
        {
            "duplicate" => [input, output, "--input", input, "-i", input],
            "mixed" => [input, "--output", output], "extra" => [input, output, "extra"],
            "extension" => [input, output + ".pdf"], _ => [input, input]
        };
        Assert.Equal(1, await global::Cli.RunAsync(["text-to-ofd", .. arguments]));
        Assert.Equal("original", await File.ReadAllTextAsync(input)); Assert.False(File.Exists(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task SymbolicFileAndParentAliasesAreRejected()
    {
        if (OperatingSystem.IsWindows()) return;
        var input = Path.Combine(_directory, "in.ofd"); await File.WriteAllTextAsync(input, "original");
        var alias = Path.Combine(_directory, "alias.ofd"); File.CreateSymbolicLink(alias, input);
        Assert.Equal(1, await global::Cli.RunAsync(["text-to-ofd", input, alias]));
        var parentAlias = _directory + "-link"; Directory.CreateSymbolicLink(parentAlias, _directory);
        try { Assert.Equal(1, await global::Cli.RunAsync(["text-to-ofd", input, Path.Combine(parentAlias, "in.ofd")])); }
        finally { Directory.Delete(parentAlias); }
        Assert.Equal("original", await File.ReadAllTextAsync(input));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task InvalidUtf8OrOutputBudgetPreservesDestination(bool budget)
    {
        var input = Path.Combine(_directory, "input.txt"); var output = Path.Combine(_directory, "out.ofd");
        await File.WriteAllBytesAsync(input, budget ? Encoding.UTF8.GetBytes("A\nB") : [0xC3, 0x28]);
        await File.WriteAllTextAsync(output, "original");
        string[] limits = budget ? ["--page-width", "40", "--page-height", "30", "--margin-left", "5", "--margin-right", "5", "--margin-top", "5", "--margin-bottom", "5", "--font-size-mm", "10", "--max-pages", "1"] : [];
        Assert.Equal(1, await global::Cli.RunAsync(["text-to-ofd", input, output, .. limits]));
        Assert.Equal("original", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task CanceledCommandKeepsDestination()
    {
        var input = Path.Combine(_directory, "input.txt"); var output = Path.Combine(_directory, "out.ofd");
        await File.WriteAllTextAsync(input, "text"); await File.WriteAllTextAsync(output, "original");
        using var cancel = new CancellationTokenSource(); cancel.Cancel();
        Assert.Equal(130, await global::Cli.RunAsync(["text-to-ofd", input, output], cancel.Token));
        Assert.Equal("original", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    public void Dispose() => Directory.Delete(_directory, true);
}
