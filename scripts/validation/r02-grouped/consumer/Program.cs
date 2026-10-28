using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Reader.Readers;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

if (args.Length != 2) throw new ArgumentException("input-directory output-directory");
var output = Path.GetFullPath(args[1]); Directory.CreateDirectory(output);
var cases = new[] {
    ("second-only", "ABCD", new[] {36,37,40,39}),
    ("both-local-zero", "ABCD", new[] {40,37,41,39}),
    ("unequal-both", "flA", new[] {1653,36,37}),
    ("mixed-second", "ABCflD", new[] {36,37,38,1653,39})
};
var rows = new List<object>();
foreach (var (name, logical, gids) in cases)
foreach (var prefix in new[] { "grouped-", "go-" })
{
    var stem = prefix + name;
    var path = Path.Combine(args[0], stem + ".ofd");
    if (!File.Exists(path)) throw new FileNotFoundException("Required producer fixture", path);
    var input = await File.ReadAllBytesAsync(path);
    var loaded = await new OfdReader().ReadAsync(new MemoryStream(input));
    var text = loaded.Pages[0].Elements.OfType<OfdTextElement>().Single();
    if (text.Text != logical || text.Runs.Count != 2) throw new Exception("Reader source mismatch: " + stem);
    var beforeXml = text.SourceXml;
    var expectedReject = stem == "go-unequal-both";
    foreach (var outlet in new[] { "pdf", "svg", "html" })
    {
        using var target = new MemoryStream();
        try
        {
            if (outlet == "pdf") await new OfdToPdfConverter(new OfdToPdfOptions { GlyphMappingStrategy = OfdGlyphMappingStrategy.GroupedLocal })
                .ConvertAsync(new MemoryStream(input), target);
            else if (outlet == "svg") await new OfdToSvgConverter(OfdGlyphMappingStrategy.GroupedLocal).ConvertAsync(new MemoryStream(input), target);
            else await new OfdToHtmlConverter(new OfdHtmlOptions { GlyphMappingStrategy = OfdGlyphMappingStrategy.GroupedLocal })
                .ConvertAsync(new MemoryStream(input), target);
        }
        catch (NotSupportedException error) when (expectedReject)
        {
            if (target.Length != 0) throw new Exception("Partial rejected output: " + stem);
            rows.Add(new { stem, outlet, outcome = "outputless-refusal", reason = error.Message });
            continue;
        }
        if (expectedReject) throw new Exception("Unexpectedly accepted writer source-count DeltaX: " + stem);
        var bytes = target.ToArray();
        if (outlet == "pdf")
        {
            using var parsed = PdfPigDocument.Open(bytes);
            if (parsed.NumberOfPages != 1 || parsed.GetPage(1).Text != logical) throw new Exception("PDF source text mismatch: " + stem);
        }
        else
        {
            var root = XDocument.Parse(Encoding.UTF8.GetString(bytes));
            var actual = root.Descendants().Where(n => n.Attribute("data-original-gid") is not null)
                .Select(n => (int)n.Attribute("data-original-gid")!).ToArray();
            if (!actual.SequenceEqual(gids)) throw new Exception("Original GID mismatch: " + stem + outlet);
            var content = string.Concat(root.Descendants().Where(n => n.Attribute("data-ofd-text") is not null)
                .Select(n => n.Attribute("data-ofd-text")!.Value));
            if (content != logical) throw new Exception("Accessible text mismatch: " + stem + outlet);
        }
        if (text.SourceXml != beforeXml) throw new Exception("Source changed: " + stem);
        await File.WriteAllBytesAsync(Path.Combine(output, stem + "." + outlet), bytes);
        rows.Add(new { stem, outlet, outcome = "pass", bytes = bytes.Length,
            inputSha256 = Convert.ToHexString(SHA256.HashData(input)).ToLowerInvariant(),
            outputSha256 = Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant() });
    }
}
var assemblies = AppDomain.CurrentDomain.GetAssemblies().Where(a => a.GetName().Name!.StartsWith("Ofdrw.Net.", StringComparison.Ordinal))
    .OrderBy(a => a.GetName().Name).Select(a => new { name = a.GetName().Name,
        location = a.Location, version = a.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
        sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(a.Location))).ToLowerInvariant() }).ToArray();
await File.WriteAllTextAsync(Path.Combine(output, "receipt.json"), JsonSerializer.Serialize(new { rows, assemblies }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine($"Grouped package probe: {rows.Count} outlet checks passed ({rows.Count - 3} accepted, 3 expected writer-delta refusals).");
