using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using System.Diagnostics;
using System.Globalization;

var repoRoot = ResolveRepoRoot();
var outputDir = Path.Combine(repoRoot, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "output");
var testDataDir = Path.Combine(repoRoot, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "testdata", "upstream-ofdrw");
Directory.CreateDirectory(outputDir);

var sourceOfdPath = Path.Combine(outputDir, "source.ofd");
var convertedPdfPath = Path.Combine(outputDir, "converted.pdf");
var roundtripOfdPath = Path.Combine(outputDir, "roundtrip.ofd");

Console.WriteLine($"[E2E] Output directory: {outputDir}");

var builder = new OfdDocumentBuilder();
builder.SetOptions(new OfdDocumentOptions
{
    DocType = "OFD-H",
    DocumentId = "Doc_0",
    Metadata = new OfdMetadata
    {
        Title = "Ofdrw.Net Package E2E",
        Creator = "Ofdrw.Net.Converter.Pdf.E2E",
        CreationDate = DateTimeOffset.UtcNow,
        ModificationDate = DateTimeOffset.UtcNow
    }
});

builder.AddPage(new OfdPage
{
    Index = 0,
    WidthMillimeters = 210,
    HeightMillimeters = 297,
    Elements =
    {
        new OfdTextElement
        {
            Text = "Package install E2E: OFD -> PDF -> OFD",
            FontName = "SimSun",
            FontSizeMillimeters = 4,
            XMillimeters = 10,
            YMillimeters = 12,
            WidthMillimeters = 160,
            HeightMillimeters = 12
        }
    }
});

var package = builder.Build();
var writer = new OfdPackageWriter();
await using (var sourceOfdStream = File.Create(sourceOfdPath))
{
    await writer.WriteAsync(package, sourceOfdStream);
}

var ofdToPdf = new OfdToPdfConverter();
await using (var ofdInput = File.OpenRead(sourceOfdPath))
await using (var pdfOutput = File.Create(convertedPdfPath))
{
    await ofdToPdf.ConvertAsync(ofdInput, pdfOutput);
}

var pdfToOfd = new PdfToOfdConverter();
await using (var pdfInput = File.OpenRead(convertedPdfPath))
await using (var ofdOutput = File.Create(roundtripOfdPath))
{
    await pdfToOfd.ConvertAsync(pdfInput, ofdOutput);
}

var reader = new OfdReader();
await using var roundtripStream = File.OpenRead(roundtripOfdPath);
var roundtrip = await reader.ReadAsync(roundtripStream);

if (roundtrip.Pages.Count == 0)
{
    throw new InvalidOperationException("Roundtrip OFD has no pages.");
}

Console.WriteLine($"[E2E] Source OFD:    {sourceOfdPath}");
Console.WriteLine($"[E2E] Converted PDF: {convertedPdfPath}");
Console.WriteLine($"[E2E] Roundtrip OFD: {roundtripOfdPath}");
Console.WriteLine($"[E2E] Roundtrip page count: {roundtrip.Pages.Count}");
Console.WriteLine("[E2E] Running upstream sample validation...");

foreach (var sampleOfd in Directory.EnumerateFiles(testDataDir, "*.ofd").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
{
    await using var stream = File.OpenRead(sampleOfd);
    var parsed = await reader.ReadAsync(stream);
    if (parsed.Pages.Count == 0)
    {
        throw new InvalidOperationException($"Sample OFD has no pages: {sampleOfd}");
    }

    Console.WriteLine($"[E2E] Upstream OFD parsed: {Path.GetFileName(sampleOfd)} ({parsed.Pages.Count} page(s))");
}

foreach (var samplePdf in Directory.EnumerateFiles(testDataDir, "*.pdf").OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
{
    await ValidatePdfSampleAsync(samplePdf, outputDir);
}

Console.WriteLine("[E2E] Success: package installation and conversion flow is working.");

static string ResolveRepoRoot()
{
    var configuredRoot = Environment.GetEnvironmentVariable("OFDRW_REPO_ROOT");
    if (!string.IsNullOrWhiteSpace(configuredRoot))
    {
        return configuredRoot;
    }

    var current = new DirectoryInfo(AppContext.BaseDirectory);
    while (current is not null)
    {
        if (File.Exists(Path.Combine(current.FullName, "Ofdrw.Net.sln")))
        {
            return current.FullName;
        }

        current = current.Parent;
    }

    throw new InvalidOperationException("Unable to locate repository root. Set OFDRW_REPO_ROOT.");
}

static async Task ValidatePdfSampleAsync(string samplePdf, string outputDir)
{
    var sampleName = Path.GetFileNameWithoutExtension(samplePdf);
    var sampleOutputDir = Path.Combine(outputDir, "upstream-ofdrw", sampleName);
    Directory.CreateDirectory(sampleOutputDir);

    var ofdPath = Path.Combine(sampleOutputDir, $"{sampleName}.ofd");
    var pdfPath = Path.Combine(sampleOutputDir, $"{sampleName}.roundtrip.pdf");

    var pdfToOfd = new PdfToOfdConverter();
    await using (var pdfInput = File.OpenRead(samplePdf))
    await using (var ofdOutput = File.Create(ofdPath))
    {
        await pdfToOfd.ConvertAsync(pdfInput, ofdOutput);
    }

    var reader = new OfdReader();
    OfdDocumentPackage parsed;
    await using (var ofdInput = File.OpenRead(ofdPath))
    {
        parsed = await reader.ReadAsync(ofdInput);
    }

    if (parsed.Pages.Count == 0)
    {
        throw new InvalidOperationException($"Converted OFD has no pages: {ofdPath}");
    }

    if (parsed.Pages.Any(page => page.Elements.OfType<OfdImageElement>().All(image => image.Data.Length == 0)))
    {
        throw new InvalidOperationException($"Converted OFD has a page without renderable image data: {ofdPath}");
    }

    var ofdToPdf = new OfdToPdfConverter();
    await using (var ofdInput = File.OpenRead(ofdPath))
    await using (var pdfOutput = File.Create(pdfPath))
    {
        await ofdToPdf.ConvertAsync(ofdInput, pdfOutput);
    }

    var sourcePng = await RenderFirstPageAsync(samplePdf, sampleOutputDir, "source");
    var roundtripPng = await RenderFirstPageAsync(pdfPath, sampleOutputDir, "roundtrip");
    await AssertImageHasContentAsync(sourcePng);
    await AssertImageHasContentAsync(roundtripPng);
    await AssertImagesSameSizeAsync(sourcePng, roundtripPng);

    Console.WriteLine($"[E2E] PDF visual smoke passed: {Path.GetFileName(samplePdf)} -> {ofdPath}");
}

static async Task<string> RenderFirstPageAsync(string pdfPath, string outputDir, string name)
{
    var outputPrefix = Path.Combine(outputDir, name);
    var result = await RunProcessAsync("pdftoppm", $"-f 1 -l 1 -png -singlefile \"{pdfPath}\" \"{outputPrefix}\"");
    if (result.ExitCode != 0)
    {
        throw new InvalidOperationException($"pdftoppm failed for {pdfPath}: {result.Error}");
    }

    var pngPath = outputPrefix + ".png";
    if (!File.Exists(pngPath))
    {
        throw new InvalidOperationException($"pdftoppm did not produce {pngPath}");
    }

    return pngPath;
}

static async Task AssertImageHasContentAsync(string imagePath)
{
    var result = await RunProcessAsync("magick", $"\"{imagePath}\" -colorspace Gray -format \"%[fx:standard_deviation]\" info:");
    if (result.ExitCode != 0)
    {
        Console.WriteLine($"[E2E] Skipping image content check because ImageMagick failed: {result.Error}");
        return;
    }

    if (!double.TryParse(result.Output, NumberStyles.Float, CultureInfo.InvariantCulture, out var standardDeviation) || standardDeviation <= 0.0001d)
    {
        throw new InvalidOperationException($"Rendered image appears blank: {imagePath}");
    }
}

static async Task AssertImagesSameSizeAsync(string expectedPath, string actualPath)
{
    var expected = await IdentifySizeAsync(expectedPath);
    var actual = await IdentifySizeAsync(actualPath);
    if (Math.Abs(expected.Width - actual.Width) > 1 || Math.Abs(expected.Height - actual.Height) > 1)
    {
        throw new InvalidOperationException($"Rendered image size mismatch. expected={expected.Width} x {expected.Height}, actual={actual.Width} x {actual.Height}");
    }
}

static async Task<(int Width, int Height)> IdentifySizeAsync(string imagePath)
{
    var result = await RunProcessAsync("magick", $"identify -format \"%w x %h\" \"{imagePath}\"");
    if (result.ExitCode != 0)
    {
        throw new InvalidOperationException($"ImageMagick identify failed for {imagePath}: {result.Error}");
    }

    var parts = result.Output.Split('x', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    if (parts.Length != 2 || !int.TryParse(parts[0], out var width) || !int.TryParse(parts[1], out var height))
    {
        throw new InvalidOperationException($"Unable to parse ImageMagick size output for {imagePath}: {result.Output}");
    }

    return (width, height);
}

static async Task<(int ExitCode, string Output, string Error)> RunProcessAsync(string fileName, string arguments)
{
    var process = new Process
    {
        StartInfo = new ProcessStartInfo
        {
            FileName = fileName,
            Arguments = arguments,
            CreateNoWindow = true,
            UseShellExecute = false,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        }
    };

    try
    {
        process.Start();
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException($"{fileName} is required for visual validation.", ex);
    }

    var stdout = await process.StandardOutput.ReadToEndAsync();
    var stderr = await process.StandardError.ReadToEndAsync();
    await process.WaitForExitAsync();
    return (process.ExitCode, stdout.Trim(), stderr.Trim());
}
