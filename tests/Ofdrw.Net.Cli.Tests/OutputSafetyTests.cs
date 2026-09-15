using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Cli.Tests;

public sealed class OutputSafetyTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofdrw-cli-tests-" + Guid.NewGuid().ToString("N"));

    public OutputSafetyTests() => Directory.CreateDirectory(_directory);

    [Fact]
    public async Task InvalidDocxPageOption_ShouldPreserveExistingOutput()
    {
        var input = Path.Combine(_directory, "input.docx");
        var output = Path.Combine(_directory, "existing.pdf");
        await File.WriteAllTextAsync(input, "unused: argument validation must run first");
        await File.WriteAllTextAsync(output, "existing output");
        var exitCode = await global::Cli.RunAsync(["docx-to-pdf", input, output, "--pages", "1"]);
        Assert.Equal(1, exitCode);
        Assert.Equal("existing output", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task InvalidInput_ShouldPreserveExistingOutput()
    {
        var input = Path.Combine(_directory, "invalid.ofd");
        var output = Path.Combine(_directory, "existing.svg");
        await File.WriteAllTextAsync(input, "not a ZIP package");
        await File.WriteAllTextAsync(output, "existing output");
        var exitCode = await global::Cli.RunAsync(["ofd-to-svg", input, output]);
        Assert.Equal(1, exitCode);
        Assert.Equal("existing output", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task CanceledCommit_ShouldDiscardPartialOutput()
    {
        var outputPath = Path.Combine(_directory, "output.ofd");
        await File.WriteAllTextAsync(outputPath, "original");
        await using (var output = new AtomicOutput(outputPath))
        {
            await output.Stream.WriteAsync(new byte[] { 1, 2, 3 });
            using var cancellation = new CancellationTokenSource();
            cancellation.Cancel();
            await Assert.ThrowsAnyAsync<OperationCanceledException>(() => output.CommitAsync(cancellation.Token));
        }

        Assert.Equal("original", await File.ReadAllTextAsync(outputPath));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task Reorder_ShouldAtomicallySupportTheSameInputAndOutputPath()
    {
        var path = Path.Combine(_directory, "document.ofd");
        var package = new OfdDocumentPackage();
        for (var index = 0; index < 2; index++)
        {
            var page = new OfdPage { Index = index, WidthMillimeters = 210, HeightMillimeters = 297 };
            page.Elements.Add(new OfdTextElement { Text = index == 0 ? "first" : "second" });
            package.Pages.Add(page);
        }

        await using (var output = File.Create(path))
        {
            await new OfdPackageWriter().WriteAsync(package, output);
        }

        Assert.Equal(0, await global::Cli.RunAsync(["reorder", path, path, "--pages", "2,1"]));
        await using var input = File.OpenRead(path);
        var result = await new OfdReader().ReadAsync(input);
        Assert.Equal("second", Assert.IsType<OfdTextElement>(Assert.Single(result.Pages[0].Elements)).Text);
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task CanceledCommand_ShouldReturnCancellationExitCode()
    {
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Equal(130, await global::Cli.RunAsync(["ofd-to-pdf", "unused", "unused"], cancellation.Token));
    }

    public void Dispose() => Directory.Delete(_directory, recursive: true);
}
