using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

public sealed class Audit004RegressionTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-audit004-" + Guid.NewGuid().ToString("N"));

    public Audit004RegressionTests() => Directory.CreateDirectory(_directory);

    [SymbolicLinkFact]
    public async Task Audit004_FinalSymlink_ShouldRejectWithoutChangingSourceOrLink()
    {
        var input = await CreateInputAsync();
        var output = Path.Combine(_directory, "output.pdf");
        File.CreateSymbolicLink(output, input);
        await AssertAliasRejectedAsync(input, output);
        Assert.Equal(input, new FileInfo(output).LinkTarget);
    }

    [SymbolicLinkFact]
    public async Task Audit004_ParentSymlink_ShouldRejectWithoutChangingSource()
    {
        var input = await CreateInputAsync();
        var alias = Path.Combine(_directory, "alias");
        Directory.CreateSymbolicLink(alias, _directory);
        await AssertAliasRejectedAsync(input, Path.Combine(alias, Path.GetFileName(input)));
    }

    [SymbolicLinkFact]
    public async Task Audit004_InputSymlink_ShouldRejectWithoutChangingSource()
    {
        var source = await CreateInputAsync();
        var input = Path.Combine(_directory, "linked.ofd");
        File.CreateSymbolicLink(input, source);
        await AssertAliasRejectedAsync(input, source);
    }

    [SymbolicLinkFact]
    public async Task Audit004_NestedRelativeParentSymlink_ShouldRejectWithoutChangingSource()
    {
        var input = await CreateInputAsync();
        Directory.CreateDirectory(Path.Combine(_directory, "links"));
        Directory.CreateSymbolicLink(Path.Combine(_directory, "links", "first"), "..");
        Directory.CreateSymbolicLink(Path.Combine(_directory, "second"), "links/first");
        await AssertAliasRejectedAsync(input, Path.Combine(_directory, "second", Path.GetFileName(input)));
    }

    [SymbolicLinkFact]
    public async Task Audit004_TargetParentAfterDirectoryLink_ShouldResolveBeforeParentTraversal()
    {
        var input = await CreateInputAsync();
        var child = Directory.CreateDirectory(Path.Combine(_directory, "nested", "child")).FullName;
        Directory.CreateSymbolicLink(Path.Combine(_directory, "jump"), child);
        Directory.CreateSymbolicLink(Path.Combine(_directory, "return"), "jump/../..");
        await AssertAliasRejectedAsync(input, Path.Combine(_directory, "return", Path.GetFileName(input)));
    }

    [SymbolicLinkFact]
    public async Task Audit004_LinkCycle_ShouldRejectWithoutResidualStage()
    {
        var input = await CreateInputAsync();
        var output = Path.Combine(_directory, "cycle-a.pdf");
        File.CreateSymbolicLink(output, "cycle-b.pdf");
        File.CreateSymbolicLink(Path.Combine(_directory, "cycle-b.pdf"), "cycle-a.pdf");
        await AssertAliasRejectedAsync(input, output);
    }

    [Fact]
    public async Task Audit004_MissingOutputLeaf_ShouldCreateIndependentPdf()
    {
        var input = await CreateInputAsync();
        await AssertIndependentOutputAsync(input, Path.Combine(_directory, "new.pdf"));
    }

    [SymbolicLinkFact]
    public async Task Audit004_MissingOutputLeafInLinkedParent_ShouldCreateIndependentPdf()
    {
        var input = await CreateInputAsync();
        var real = Directory.CreateDirectory(Path.Combine(_directory, "real-output")).FullName;
        var alias = Path.Combine(_directory, "output-alias");
        Directory.CreateSymbolicLink(alias, real);
        await AssertIndependentOutputAsync(input, Path.Combine(alias, "new.pdf"));
        Assert.True(File.Exists(Path.Combine(real, "new.pdf")));
    }

    [NativeCaseFact]
    public async Task Audit004_NativeCaseAlias_ShouldRejectOnActualNativePlatform()
    {
        var input = await CreateInputAsync();
        await AssertAliasRejectedAsync(input, Path.Combine(_directory, "INPUT.OFD"));
    }

    [Fact]
    public async Task Audit004_IntentionalInPlace_ShouldStillReorder()
    {
        var input = await CreateInputAsync();
        Assert.Equal(0, await global::Cli.RunAsync(["reorder", input, input, "--pages", "2,1"]));
        await using var stream = File.OpenRead(input);
        var package = await new OfdReader().ReadAsync(stream);
        Assert.Equal("second", Assert.IsType<OfdTextElement>(Assert.Single(package.Pages[0].Elements)).Text);
        AssertNoStage();
    }

    private async Task<string> CreateInputAsync()
    {
        var input = Path.Combine(_directory, "input.ofd");
        var package = new OfdDocumentPackage();
        foreach (var text in new[] { "first", "second" })
        {
            var page = new OfdPage { Index = package.Pages.Count, WidthMillimeters = 210, HeightMillimeters = 297 };
            page.Elements.Add(new OfdTextElement { Text = text, XMillimeters = 20, YMillimeters = 20, WidthMillimeters = 100, HeightMillimeters = 10 });
            package.Pages.Add(page);
        }
        await using var stream = File.Create(input);
        await new OfdPackageWriter().WriteAsync(package, stream);
        return input;
    }

    private async Task AssertAliasRejectedAsync(string input, string output)
    {
        var before = await File.ReadAllBytesAsync(input);
        Assert.Equal(1, await global::Cli.RunAsync(["ofd-to-pdf", input, output]));
        Assert.Equal(before, await File.ReadAllBytesAsync(input));
        AssertNoStage();
    }

    private async Task AssertIndependentOutputAsync(string input, string output)
    {
        var before = await File.ReadAllBytesAsync(input);
        Assert.Equal(0, await global::Cli.RunAsync(["ofd-to-pdf", input, output]));
        Assert.StartsWith("%PDF-", await File.ReadAllTextAsync(output));
        Assert.Equal(before, await File.ReadAllBytesAsync(input));
        AssertNoStage();
    }

    private void AssertNoStage() => Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp", new EnumerationOptions { RecurseSubdirectories = true, AttributesToSkip = FileAttributes.ReparsePoint }));

    public void Dispose()
    {
        // Remove directory links before recursive cleanup; never traverse the alias back to this directory.
        foreach (var path in Directory.GetFileSystemEntries(_directory))
            if (new DirectoryInfo(path).LinkTarget is not null && Directory.Exists(path)) Directory.Delete(path);
        Directory.Delete(_directory, recursive: true);
    }

    public sealed class SymbolicLinkFactAttribute : FactAttribute
    {
        public SymbolicLinkFactAttribute()
        {
            var directory = Path.Combine(Path.GetTempPath(), "ofdrw-link-probe-" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(directory);
            try
            {
                File.WriteAllText(Path.Combine(directory, "target"), "probe");
                File.CreateSymbolicLink(Path.Combine(directory, "file-link"), "target");
                Directory.CreateSymbolicLink(Path.Combine(directory, "directory-link"), directory);
            }
            catch (Exception error) when (error is UnauthorizedAccessException or IOException or PlatformNotSupportedException)
            {
                Skip = $"Symbolic links unavailable on {System.Runtime.InteropServices.RuntimeInformation.OSDescription}: {error.Message}";
            }
            finally
            {
                var link = Path.Combine(directory, "directory-link");
                if (new DirectoryInfo(link).LinkTarget is not null) Directory.Delete(link);
                Directory.Delete(directory, recursive: true);
            }
        }
    }

    public sealed class NativeCaseFactAttribute : FactAttribute
    {
        public NativeCaseFactAttribute()
        {
            if (!OperatingSystem.IsWindows() && !OperatingSystem.IsMacOS())
                Skip = "Requires the actual Windows/macOS case comparison branch.";
        }
    }
}
