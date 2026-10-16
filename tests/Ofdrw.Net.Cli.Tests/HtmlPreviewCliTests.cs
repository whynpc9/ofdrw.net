using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class HtmlPreviewCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-html-cli-" + Guid.NewGuid().ToString("N"));
    public HtmlPreviewCliTests() => Directory.CreateDirectory(_directory);
    [Theory]
    [InlineData("")][InlineData(" ")][InlineData("1,,2")][InlineData("0")][InlineData("-1")][InlineData("3-1")][InlineData("1-2147483647")][InlineData("2147483648")]
    public async Task StrictSelectionFailsBeforeReadingInputAndPreservesTarget(string pages)
    {
        var output = Path.Combine(_directory, "existing.html"); await File.WriteAllTextAsync(output, "original");
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-html", "missing.ofd", output, "--pages", pages]));
        Assert.Equal("original", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    [Fact]
    public void BoundedParserPreservesRangesAndDuplicates()
    {
        Assert.Equal(new[] { 2, 0, 2 }, global::Cli.ParseHtmlPages("3,1,3", 3, default));
        Assert.Equal(new[] { 0, 1, 2, 1 }, global::Cli.ParseHtmlPages("1-3,2", 4, default));
        Assert.Throws<ArgumentException>(() => global::Cli.ParseHtmlPages(new string('1', 8193), 200, default));
        Assert.Throws<ArgumentException>(() => global::Cli.ParseHtmlPages("1-201", 200, default));
    }
    [Fact]
    public async Task ExplicitCommandCommitsCompleteHtmlAndDoesNotChangeSingleSvgBehavior()
    {
        var input = await Source(); var output = Path.Combine(_directory, "output.html"); await File.WriteAllTextAsync(output, "old");
        Assert.Equal(0, await global::Cli.RunAsync(["ofd-to-html", input, output, "--pages", "2,1,2", "--title", "<literal>&"]));
        var text = await File.ReadAllTextAsync(output); Assert.Contains("&lt;literal&gt;&amp;", text); Assert.Equal(3, text.Split("<section").Length - 1); Assert.EndsWith("</html>", text);
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-svg", input, Path.Combine(_directory, "unchanged.svg"), "--pages", "2,1,2"]));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    [Theory]
    [InlineData("--pages", "3")][InlineData("--max-pages", "1")][InlineData("--max-page-bytes", "100")][InlineData("--max-pages", "0")]
    public async Task InvalidPageOrBudgetPreservesExistingTarget(string option, string value)
    {
        var input = await Source(); var output = Path.Combine(_directory, "existing.html"); await File.WriteAllTextAsync(output, "original");
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-html", input, output, option, value]));
        Assert.Equal("original", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    [Fact]
    public async Task SamePathAndCancellationAreRejectedWithoutDamage()
    {
        var input = await Source(); var bytes = await File.ReadAllBytesAsync(input);
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-html", input, input])); Assert.Equal(bytes, await File.ReadAllBytesAsync(input));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Equal(130, await global::Cli.RunAsync(["ofd-to-html", input, Path.Combine(_directory, "cancel.html")], cancellation.Token));
        Assert.False(File.Exists(Path.Combine(_directory, "cancel.html"))); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    [Fact]
    public async Task RenameFailureReturnsFailureAndCleansTemporaryFile()
    {
        var input = await Source(); var output = Path.Combine(_directory, "directory.html"); Directory.CreateDirectory(output);
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-html", input, output])); Assert.True(Directory.Exists(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }
    private async Task<string> Source()
    {
        var package = new OfdDocumentPackage();
        for (var i = 0; i < 2; i++) { var page = new OfdPage { Index = i, WidthMillimeters = 100, HeightMillimeters = 120 }; page.Elements.Add(new OfdTextElement { Text = "page " + i }); package.Pages.Add(page); }
        var input = Path.Combine(_directory, "input.ofd"); await using var stream = File.Create(input); await new OfdPackageWriter().WriteAsync(package, stream); return input;
    }
    public void Dispose() => Directory.Delete(_directory, true);
}
