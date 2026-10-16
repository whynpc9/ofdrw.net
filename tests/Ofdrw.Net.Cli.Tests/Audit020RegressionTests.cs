using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;

namespace Ofdrw.Net.Cli.Tests;

[Collection("Owned CLI console")]
public sealed class Audit020RegressionTests : IDisposable
{
    private readonly string directory = Path.Combine(Path.GetTempPath(), "ofdrw-audit020-" + Guid.NewGuid().ToString("N"));

    public Audit020RegressionTests() => Directory.CreateDirectory(directory);

    [Fact]
    public async Task Audit020_NormalControl()
    {
        var input = await WriteInputAsync();
        var output = PathFor("control.pdf");
        using var stdout = new StringWriter();
        var result = await RunAsync(["ofd-to-pdf", input, output], stdout);
        Assert.Equal(0, result.Code);
        Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(await File.ReadAllBytesAsync(output)));
        Assert.Contains($"Converted {input} -> {output}", stdout.ToString());
        Assert.Equal("", result.Error);
        AssertNoStagingFiles();
    }

    [Theory]
    [InlineData("ofd-to-pdf")]
    [InlineData("ofd-to-svg")]
    [InlineData("extract-text")]
    [InlineData("merge")]
    [InlineData("image-to-ofd")]
    [InlineData("ofd-to-image")]
    public async Task Audit020_ConvertedThenThrowingStdout(string command)
    {
        var input = await WriteInputAsync();
        var output = PathFor("output" + (command switch
        {
            "ofd-to-pdf" => ".pdf", "ofd-to-svg" => ".svg", "extract-text" => ".txt",
            "ofd-to-image" => ".png", _ => ".ofd"
        }));
        string[] args = [command, input, output];
        if (command == "merge") args = [command, output, input, input];
        if (command == "image-to-ofd")
        {
            input = PathFor("input.png");
            using var image = new Image<Rgb24>(20, 10, new Rgb24(255, 0, 0));
            image.SaveAsPng(input);
            args = [command, input, output];
        }
        var original = await File.ReadAllBytesAsync(input);
        await File.WriteAllTextAsync(output, "existing target sentinel");
        var stdout = new ThrowingWriter();
        var result = await RunAsync(args, stdout);
        // Inspect actual committed output before the exit-code assertion so a red
        // baseline distinguishes failed notification from a failed conversion.
        if (command == "ofd-to-pdf")
            Assert.StartsWith("%PDF-", Encoding.ASCII.GetString(await File.ReadAllBytesAsync(output)));
        else if (command == "ofd-to-svg")
            Assert.Equal("svg", XDocument.Load(output).Root!.Name.LocalName);
        else if (command == "extract-text")
            Assert.Contains("Page 1", await File.ReadAllTextAsync(output));
        else if (command == "ofd-to-image")
        {
            using var image = Image.Load<Rgb24>(output);
            Assert.True(image.Width > 0 && image.Height > 0);
        }
        else
        {
            using var stream = File.OpenRead(output);
            var package = await new OfdReader().ReadAsync(stream);
            Assert.Equal(command == "merge" ? 4 : 1, package.Pages.Count);
            if (command == "image-to-ofd") Assert.IsType<OfdImageElement>(Assert.Single(package.Pages[0].Elements));
        }
        Assert.Equal(original, await File.ReadAllBytesAsync(input));
        Assert.Equal(1, stdout.Attempts);
        AssertNoStagingFiles();
        Assert.Equal(0, result.Code);
        Assert.Equal("", result.Error);
    }

    [Fact]
    public async Task Audit020_InPlaceReorderOnce()
    {
        var input = await WriteInputAsync();
        var stdout = new ThrowingWriter();
        var result = await RunAsync(["reorder", input, input, "--pages", "2,1"], stdout);
        using var stream = File.OpenRead(input);
        var package = await new OfdReader().ReadAsync(stream);
        Assert.Equal(new[] { "Page 2", "Page 1" }, package.Pages.Select(page => Assert.IsType<OfdTextElement>(Assert.Single(page.Elements)).Text));
        Assert.Equal(1, stdout.Attempts);
        AssertNoStagingFiles();
        Assert.Equal(0, result.Code);
        Assert.Equal("", result.Error);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task Audit020_PreCommitFailure(bool publicationFailure)
    {
        var input = await WriteInputAsync();
        var output = PathFor("target.pdf");
        if (publicationFailure) Directory.CreateDirectory(output);
        else await File.WriteAllTextAsync(input, "invalid OFD sentinel");
        var sentinel = publicationFailure ? Path.Combine(output, "sentinel") : output;
        await File.WriteAllTextAsync(sentinel, "unchanged target");
        var original = await File.ReadAllBytesAsync(input);
        var stdout = new ThrowingWriter();
        var result = await RunAsync(["ofd-to-pdf", input, output], stdout);
        Assert.Equal(1, result.Code);
        Assert.Contains("Error:", result.Error);
        Assert.Equal(0, stdout.Attempts);
        Assert.Equal("unchanged target", await File.ReadAllTextAsync(sentinel));
        Assert.Equal(original, await File.ReadAllBytesAsync(input));
        AssertNoStagingFiles();
    }

    [Fact]
    public async Task Audit020_StdoutPayloadFailure()
    {
        var input = await WriteInputAsync();
        var original = await File.ReadAllBytesAsync(input);
        var stdout = new ThrowingWriter();
        var result = await RunAsync(["extract-text", input], stdout);
        Assert.Equal(1, result.Code);
        Assert.Contains("Error: closed success output", result.Error);
        Assert.Equal(1, stdout.Attempts);
        Assert.Equal(original, await File.ReadAllBytesAsync(input));
        AssertNoStagingFiles();
    }

    [Fact]
    public async Task Audit020_CancellationBeforeCommit()
    {
        var input = await WriteInputAsync();
        var output = PathFor("target.pdf");
        await File.WriteAllTextAsync(output, "unchanged target");
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        var stdout = new ThrowingWriter();
        var result = await RunAsync(["ofd-to-pdf", input, output], stdout, cancellation.Token);
        Assert.Equal(130, result.Code);
        Assert.Contains("Conversion canceled.", result.Error);
        Assert.Equal(0, stdout.Attempts);
        Assert.Equal("unchanged target", await File.ReadAllTextAsync(output));
        AssertNoStagingFiles();
    }

    [Fact]
    public async Task Audit020_CancellationAfterCommitNotification()
    {
        var input = await WriteInputAsync();
        var output = PathFor("target.pdf");
        using var cancellation = new CancellationTokenSource();
        var observedCommittedPdf = false;
        var stdout = new ThrowingWriter(() =>
        {
            observedCommittedPdf = Encoding.ASCII.GetString(File.ReadAllBytes(output)).StartsWith("%PDF-", StringComparison.Ordinal);
            cancellation.Cancel();
            throw new OperationCanceledException(cancellation.Token);
        });
        var result = await RunAsync(["ofd-to-pdf", input, output], stdout, cancellation.Token);
        Assert.True(observedCommittedPdf);
        Assert.True(cancellation.IsCancellationRequested);
        Assert.Equal(1, stdout.Attempts);
        AssertNoStagingFiles();
        Assert.Equal(0, result.Code);
        Assert.Equal("", result.Error);
    }

    private async Task<string> WriteInputAsync()
    {
        var path = PathFor("input.ofd");
        var package = new OfdDocumentPackage();
        for (var index = 1; index <= 2; index++)
        {
            var page = new OfdPage { WidthMillimeters = 40, HeightMillimeters = 30 };
            page.Elements.Add(new OfdTextElement
            {
                Text = "Page " + index, FontName = "Arial", FontSizeMillimeters = 4,
                XMillimeters = 5, YMillimeters = 5, WidthMillimeters = 30, HeightMillimeters = 10
            });
            package.Pages.Add(page);
        }
        using var stream = File.Create(path);
        await new OfdPackageWriter().WriteAsync(package, stream);
        return path;
    }

    private static async Task<(int Code, string Error)> RunAsync(string[] args, TextWriter stdout, CancellationToken token = default)
    {
        var savedOut = Console.Out;
        var savedError = Console.Error;
        using var stderr = new StringWriter();
        try
        {
            Console.SetOut(stdout);
            Console.SetError(stderr);
            return (await global::Cli.RunAsync(args, token), stderr.ToString());
        }
        finally
        {
            Console.SetOut(savedOut);
            Console.SetError(savedError);
        }
    }

    private string PathFor(string name) => Path.Combine(directory, name);
    private void AssertNoStagingFiles() => Assert.Empty(Directory.GetFiles(directory, ".ofdrw-*.tmp"));
    public void Dispose() => Directory.Delete(directory, true);

    private sealed class ThrowingWriter(Action? onWrite = null) : TextWriter
    {
        internal int Attempts { get; private set; }
        public override Encoding Encoding => Encoding.UTF8;
        public override void WriteLine(string? value)
        {
            Attempts++;
            onWrite?.Invoke();
            throw new IOException("closed success output");
        }
    }
}
