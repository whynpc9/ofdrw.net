using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Crypto.Password;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

var root = Path.GetFullPath(args[0]);
var output = Path.GetFullPath(args[1]);
var fonts = Path.GetFullPath(args[2]);
Directory.CreateDirectory(output);
var font = Path.Combine(fonts, "Ofdrw-CI-NotoSansCJKsc-Regular.ttf");
if (Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(font))).ToLowerInvariant() != "3012a9b63f5eca3e3b38f23a1be5ed504675e394abf8e7a4fa981506582c04aa") throw new Exception("Pinned fixture font mismatch.");
File.Copy(Path.Combine(fonts, "Ofdrw-CI-Noto-OFL.txt"), Path.Combine(output, "Ofdrw-CI-Noto-OFL.txt"), true);
var docx = Path.Combine(output, "licensed-layout.docx");
File.Copy(Path.Combine(root, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx"), docx, true);
using (var zip = ZipFile.Open(docx, ZipArchiveMode.Update))
foreach (var path in new[] { "word/document.xml", "word/styles.xml" })
{
    var entry = zip.GetEntry(path)!; XDocument xml;
    using (var input = entry.Open()) xml = XDocument.Load(input);
    var ns = xml.Root!.Name.Namespace;
    foreach (var run in xml.Descendants(ns + "r"))
    {
        var properties = run.Element(ns + "rPr");
        if (properties is null) { properties = new XElement(ns + "rPr"); run.AddFirst(properties); }
        if (properties.Element(ns + "rFonts") is null) properties.AddFirst(new XElement(ns + "rFonts"));
    }
    foreach (var node in xml.Descendants(ns + "rFonts"))
    {
        node.RemoveAttributes(); foreach (var name in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) node.SetAttributeValue(ns + name, "Noto Sans CJK SC");
    }
    entry.Delete(); using var target = zip.CreateEntry(path).Open(); xml.Save(target);
}
var results = new Dictionary<string, object>();
foreach (var mode in new[] { "native", "default" })
{
    var baselinePath = Path.Combine(output, "baseline-" + mode + ".ofd");
    var options = new DocxConversionOptions { FontDirectories = { fonts } };
    if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
    using (var input = File.OpenRead(docx)) using (var target = File.Create(baselinePath))
        await new DocxToOfdConverter(options).ConvertAsync(input, target);
    // Embed the pinned licensed face into this synthetic sample, using the existing public Writer.
    using (var input = File.OpenRead(baselinePath))
    {
        var package = await new OfdReader().ReadAsync(input);
        foreach (var item in package.Fonts) { item.Data = File.ReadAllBytes(font); item.FileName = Path.GetFileName(font); }
        using var target = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, target);
        input.Dispose(); File.WriteAllBytes(baselinePath, target.ToArray());
    }
    var original = ReadEntries(baselinePath);
    using (var input = File.OpenRead(baselinePath))
    {
        var package = await new OfdReader().ReadAsync(input);
        if (package.Pages.Count != 2) throw new Exception("Deterministic fixture must have two pages.");
    }
    await Pdf(baselinePath);
    foreach (var scope in new[] { "partial", "all" })
    {
        var selection = new OfdPasswordOptions { LoadOptions = new() { MaxInputBytes = 128L * 1024 * 1024, MaxEntryUncompressedBytes = 32L * 1024 * 1024, MaxTotalUncompressedBytes = 128L * 1024 * 1024 }, MaxOutputBytes = 160L * 1024 * 1024 };
        if (scope == "partial") selection.PageIndices.Add(1);
        else foreach (var path in original.Keys) selection.EntryNames.Add(path);
        var name = mode + "-" + scope;
        var encryptedPath = Path.Combine(output, name + "-encrypted.ofd");
        var restoredPath = Path.Combine(output, name + "-restored.ofd");
        await OfdPasswordEnvelope.EncryptFileAsync(baselinePath, encryptedPath, "synthetic-test-password-口令", selection);
        await OfdPasswordEnvelope.DecryptFileAsync(encryptedPath, restoredPath, "synthetic-test-password-口令", selection);
        var restored = ReadEntries(restoredPath);
        if (!original.Keys.Order().SequenceEqual(restored.Keys.Order()) || original.Any(pair => !pair.Value.SequenceEqual(restored[pair.Key])))
            throw new Exception("Restored entry names/bytes differ from source.");
        var pdfPath = await Pdf(restoredPath);
        using var input = File.OpenRead(restoredPath);
        var package = await new OfdReader().ReadAsync(input);
        var text = new OfdTextExtractor().Extract(package, includeTemplates: true);
        File.WriteAllText(Path.Combine(output, name + "-restored.txt"), text);
        if (!text.Contains("Generated DOCX") || !text.Contains("第二页")) throw new Exception("Restored original text is incomplete.");
        results[name] = new { original_entries = original.Count, restored_entries = restored.Count, entry_bytes_equal = true,
            pages = package.Pages.Count, source_bytes = new FileInfo(baselinePath).Length, encrypted_bytes = new FileInfo(encryptedPath).Length,
            restored_bytes = new FileInfo(restoredPath).Length, pdf = Path.GetFileName(pdfPath), text_complete = true };
    }
}
File.WriteAllText(Path.Combine(output, "functional.json"), JsonSerializer.Serialize(results, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("Native/default x partial/all: exact entry recovery, two-page text and PDF exports passed.");

Dictionary<string, byte[]> ReadEntries(string path)
{
    using var archive = ZipFile.OpenRead(path);
    return archive.Entries.Where(entry => entry.Name.Length > 0).ToDictionary(entry => entry.FullName, entry =>
    { using var input = entry.Open(); using var target = new MemoryStream(); input.CopyTo(target); return target.ToArray(); });
}
async Task<string> Pdf(string path)
{
    var pdf = Path.ChangeExtension(path, ".pdf"); using var input = File.OpenRead(path); using var outputFile = File.Create(pdf);
    await new OfdToPdfConverter().ConvertAsync(input, outputFile); return pdf;
}
