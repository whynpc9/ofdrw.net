using System.Diagnostics;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

PdfFontRegistry.EnsureInstalled();
if (args[0] == "generate")
{
    Directory.CreateDirectory(args[1]);
    foreach (var count in new[] { 1, 10, 100 })
    {
        using var document = new PdfDocument();
        document.Info.Title = $"Deterministic {count}-page conversion benchmark";
        document.Info.CreationDate = new DateTime(2026, 1, 1, 0, 0, 0, DateTimeKind.Utc);
        for (var index = 0; index < count; index++)
        {
            var page = document.AddPage(); page.Width = 432; page.Height = 576;
            using var graphics = XGraphics.FromPdfPage(page);
            graphics.DrawString($"BENCHMARK PAGE {index + 1:000}", new XFont("Arial", 18), XBrushes.DarkBlue, new XPoint(36, 54));
            for (var row = 0; row < 12; row++)
            {
                graphics.DrawRectangle(row % 2 == 0 ? XBrushes.LightGray : XBrushes.White, 36, 80 + row * 28, 360, 28);
                graphics.DrawRectangle(XPens.DarkGray, 36, 80 + row * 28, 360, 28);
                graphics.DrawString($"Row {row + 1:00} - original document content - {index + 1:000}", new XFont("Arial", 10), XBrushes.Black, new XPoint(44, 99 + row * 28));
            }
        }
        document.Save(Path.Combine(args[1], $"pages-{count}.pdf"));
    }
    return;
}

var kind = args[0];
var inputPath = args[1];
var reportPath = args[2];
var countToMeasure = int.Parse(args[3]);
var label = args[4];
Func<Stream, Stream, Task> convert = kind == "pdf"
    ? (input, output) => new PdfToOfdConverter(new PdfToOfdOptions { Dpi = 150, TextLayerMode = PdfTextLayerMode.None }).ConvertAsync(input, output)
    : (input, output) => new DocxToOfdConverter(new DocxConversionOptions { Engine = DocxConversionEngine.BuiltIn, OfdMode = DocxToOfdMode.DualLayer }, new PdfToOfdOptions { Dpi = 150 }).ConvertAsync(input, output);
// Warm the native/font libraries outside the timed trials.
var warmup = kind == "pdf" ? Path.Combine(Path.GetDirectoryName(inputPath)!, "pages-1.pdf") : inputPath;
using (var input = File.OpenRead(warmup))
using (var output = new MemoryStream()) await convert(input, output);
var trials = new List<object>();
for (var trial = 0; trial < countToMeasure; trial++)
{
    GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
    using var input = File.OpenRead(inputPath);
    using var output = new MemoryStream();
    var allocatedBefore = GC.GetTotalAllocatedBytes(precise: true);
    var timer = Stopwatch.StartNew();
    await convert(input, output);
    timer.Stop();
    trials.Add(new { elapsedMs = timer.Elapsed.TotalMilliseconds, allocatedBytes = GC.GetTotalAllocatedBytes(precise: true) - allocatedBefore, outputBytes = output.Length });
    if (trial == countToMeasure - 1) await File.WriteAllBytesAsync(Path.ChangeExtension(reportPath, ".ofd"), output.ToArray());
}
var report = new
{
    label, kind, source = Path.GetFileName(inputPath), sourceSha256 = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(inputPath))),
    libraryVersion = typeof(PdfToOfdConverter).Assembly.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
    architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString(),
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    dpi = 150, trials, peakWorkingSetBytes = Process.GetCurrentProcess().PeakWorkingSet64
};
await File.WriteAllTextAsync(reportPath, JsonSerializer.Serialize(report, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"{label}: {kind} {Path.GetFileName(inputPath)} -> {reportPath}");
