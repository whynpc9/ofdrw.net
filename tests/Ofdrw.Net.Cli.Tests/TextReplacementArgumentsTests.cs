namespace Ofdrw.Net.Cli.Tests;

public sealed class TextReplacementArgumentsTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofdrw-replace-args-" + Guid.NewGuid().ToString("N"));
    private const string Digest = "0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef";
    public TextReplacementArgumentsTests() => Directory.CreateDirectory(directory);
    private string Input => Path.Combine(directory, "input.ofd");
    private string Output => Path.Combine(directory, "output.ofd");
    private string[] Valid(string replacement = "-") => [Input, Output, "--find", "中文A", "--replace", replacement, "--trusted-font", Digest + ":100"];

    [Theory]
    [InlineData("")]
    [InlineData("--help")]
    [InlineData("-中文")]
    public void LiteralOptionLookingValuesAndDeletionAreUnambiguous(string replacement)
    {
        var result = TextReplacementArguments.Parse(Valid(replacement));
        Assert.Equal("中文A", result.Keyword); Assert.Equal(replacement, result.Replacement);
        Assert.Equal(Path.GetFullPath(Input), result.InputPath); Assert.Equal(Path.GetFullPath(Output), result.OutputPath);
        Assert.Null(result.Pages); Assert.Equal(Digest, Assert.Single(result.TrustedFonts).Sha256);
    }

    [Fact]
    public void PageSelectionKeepsRequestedOrderAndUsesZeroBasedPositions()
    {
        var result = TextReplacementArguments.Parse([.. Valid(), "--pages", "3,1-2"]);
        Assert.Equal(new[] { 2, 0, 1 }, result.Pages);
    }

    [Theory]
    [InlineData("1-2147483647")]
    [InlineData("1-1001")]
    [InlineData("1,1")]
    [InlineData("1-2,2")]
    [InlineData("0")]
    [InlineData("2-1")]
    [InlineData("1-")]
    [InlineData("1,,2")]
    public void InvalidOrUnboundedPageSelectionsRejectBeforeAnyFileWork(string pages)
        => Assert.Throws<ArgumentException>(() => TextReplacementArguments.Parse([.. Valid(), "--pages", pages]));

    [Fact]
    public void IdenticalCanonicalPathsAndConflictingTrustProfilesReject()
    {
        var same = Valid(); same[1] = Path.Combine(directory, ".", "input.ofd");
        Assert.Throws<ArgumentException>(() => TextReplacementArguments.Parse(same));
        Assert.Throws<ArgumentException>(() => TextReplacementArguments.Parse([.. Valid(), "--trusted-font", Digest + ":101"]));
        Assert.Throws<ArgumentException>(() => TextReplacementArguments.Parse([.. Valid(), "--find", "B"]));
    }

    [Theory]
    [InlineData("pages")]
    [InlineData("unknown")]
    [InlineData("duplicate")]
    [InlineData("trust-format")]
    [InlineData("trust-conflict")]
    [InlineData("missing-value")]
    [InlineData("empty-find")]
    public async Task InvalidReplacementCommandPreservesInputTargetAndCreatesNoStage(string kind)
    {
        await File.WriteAllTextAsync(Input, "original input sentinel"); await File.WriteAllTextAsync(Output, "existing target sentinel");
        string[] args = kind switch
        {
            "pages" => [.. Valid(), "--pages", "1-2147483647"],
            "unknown" => [.. Valid(), "--unknown", "value"],
            "duplicate" => [.. Valid(), "--replace", "X"],
            "trust-format" => [.. Valid(), "--trusted-font", "not-a-digest:100"],
            "trust-conflict" => [.. Valid(), "--trusted-font", Digest + ":101"],
            "missing-value" => [.. Valid(), "--pages"],
            "empty-find" => [Input, Output, "--find", "", "--replace", "X", "--trusted-font", Digest + ":100"],
            _ => throw new InvalidOperationException()
        };
        Assert.Equal(1, await global::Cli.RunAsync(["replace-text", .. args]));
        Assert.Equal("original input sentinel", await File.ReadAllTextAsync(Input));
        Assert.Equal("existing target sentinel", await File.ReadAllTextAsync(Output));
        Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
    }
    public void Dispose() => Directory.Delete(directory, true);
}
