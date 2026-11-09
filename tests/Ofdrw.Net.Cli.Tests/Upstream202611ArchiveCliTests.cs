using System.Diagnostics;
using Ofdrw.Net.TestSupport;
using SixLabors.ImageSharp;
using SixLabors.ImageSharp.PixelFormats;
using UglyToad.PdfPig;
using UglyToad.PdfPig.Content;

namespace Ofdrw.Net.Cli.Tests;

public sealed class Upstream202611ArchiveCliTests : IDisposable
{
    private readonly string directory = Directory.CreateTempSubdirectory("ofdrw-plan019-").FullName;

    [Theory]
    [InlineData("plain")]
    [InlineData("adjusted")]
    public async Task ConvertsMixedCaseBranchesFromDifferentWorkingDirectories(string layout)
    {
        var input = await WriteInput(layout, "none");
        var output = Path.Combine(directory, "output.pdf");
        var alternate = Path.Combine(directory, "alternate.pdf");

        await AssertConversion(input, output, directory);
        var otherDirectory = Directory.CreateDirectory(Path.Combine(directory, "unrelated-cwd")).FullName;
        await AssertConversion(input, alternate, otherDirectory);

        SaveEvidence(input, layout == "plain" ? "archive-template-case" : layout, output, alternate);
    }

    [Theory]
    [InlineData("unadjusted", "none")]
    [InlineData("plain", "collision")]
    [InlineData("plain", "missing-page")]
    [InlineData("plain", "traversal")]
    [InlineData("plain", "truncated")]
    [InlineData("plain", "false-local-header")]
    public async Task RejectsInvalidFileWithoutPublishingPdf(string layout, string defect)
    {
        var input = await WriteInput(layout, defect);
        var output = Path.Combine(directory, "output.pdf");
        var (code, error) = await ConvertFile(input, output, directory);
        Assert.Equal(1, code);
        Assert.NotEmpty(error);
        Assert.False(File.Exists(output));
        SaveEvidence(input, defect == "none" ? layout : defect);
    }

    private static async Task AssertConversion(string input, string output, string workingDirectory)
    {
        var (code, error) = await ConvertFile(input, output, workingDirectory);
        Assert.True(code == 0, error);
        using var pdf = PdfDocument.Open(output);
        Assert.Equal(1, pdf.NumberOfPages);
        var page = pdf.GetPage(1);
        var images = page.GetImages().OrderBy(image => image.BoundingBox.Left).ToArray();
        Assert.Equal(2, images.Length);
        AssertImage(images[0], 10, new Rgb24(0, 0, 255));
        AssertImage(images[1], 60, new Rgb24(255, 0, 0));
    }

    private static void AssertImage(IPdfImage image, int leftMm, Rgb24 color)
    {
        const double pointsPerMillimeter = 72d / 25.4d;
        Assert.InRange(image.BoundingBox.Left / pointsPerMillimeter, leftMm - 0.1, leftMm + 0.1);
        Assert.InRange(image.BoundingBox.Bottom / pointsPerMillimeter, 19.9, 20.1);
        Assert.InRange(image.BoundingBox.Width / pointsPerMillimeter, 29.9, 30.1);
        Assert.InRange(image.BoundingBox.Height / pointsPerMillimeter, 29.9, 30.1);
        Assert.True(image.TryGetPng(out var png));
        using var decoded = Image.Load<Rgb24>(png);
        Assert.Equal(color, decoded[0, 0]);
    }

    private static async Task<(int ExitCode, string Error)> ConvertFile(string input, string output, string workingDirectory)
    {
        var dll = Path.GetFullPath(typeof(global::Cli).Assembly.Location);
        var dotnet = Environment.GetEnvironmentVariable("DOTNET_HOST_PATH") ?? "dotnet";
        var start = new ProcessStartInfo(dotnet)
        {
            WorkingDirectory = workingDirectory,
            RedirectStandardError = true,
            UseShellExecute = false
        };
        foreach (var argument in new[] { dll, "ofd-to-pdf", "-i", input, "-o", output })
            start.ArgumentList.Add(argument);

        using var process = Process.Start(start)!;
        var errors = process.StandardError.ReadToEndAsync();
        using var timeout = new CancellationTokenSource(TimeSpan.FromSeconds(60));
        try
        {
            await process.WaitForExitAsync(timeout.Token);
        }
        catch (OperationCanceledException)
        {
            if (!process.HasExited)
                process.Kill(entireProcessTree: true);
            await process.WaitForExitAsync();
            throw;
        }
        return (process.ExitCode, await errors);
    }

    private async Task<string> WriteInput(string layout, string defect)
    {
        var path = Path.Combine(directory, "input.ofd");
        await File.WriteAllBytesAsync(path, Upstream202611ArchiveFixture.Create(layout, defect));
        return path;
    }

    private static void SaveEvidence(string input, string name, params string[] outputs)
    {
        var root = Environment.GetEnvironmentVariable("OFDRW_120_ARCHIVE_ARTIFACT_DIR");
        if (string.IsNullOrWhiteSpace(root))
            return;

        var destination = Directory.CreateDirectory(Path.Combine(root, name)).FullName;
        foreach (var path in outputs.Prepend(input))
            File.Copy(path, Path.Combine(destination, Path.GetFileName(path)));
    }

    public void Dispose() => Directory.Delete(directory, recursive: true);
}
