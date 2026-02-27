using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

var repoRoot = ResolveRepoRoot();
var outputDir = Path.Combine(repoRoot, "e2e", "Ofdrw.Net.Converter.Pdf.E2E", "output");
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
