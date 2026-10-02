using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;

var root = Path.GetFullPath(args[0]); var output = Path.GetFullPath(args[1]); var fonts = Path.GetFullPath(args[2]);
string Target(string name) => Path.Combine(output, name);
var fontBytes = File.ReadAllBytes(Path.Combine(fonts, "Ofdrw-CI-NotoSansCJKsc-Regular.ttf"));
Directory.CreateDirectory(Target("fonts"));
foreach (var path in Directory.GetFiles(fonts)) File.Copy(path, Target("fonts/" + Path.GetFileName(path)), true);
async Task<OfdDocumentPackage> Read(string name)
{ await using var input = File.OpenRead(Target(name + ".ofd")); return await new OfdReader().ReadAsync(input); }
async Task Save(OfdDocumentPackage package, string name)
{ await using var destination = File.Create(Target(name + ".ofd")); await new OfdPackageWriter().WriteAsync(package, destination); }
async Task Export(string name)
{
    var package = await Read(name);
    File.WriteAllText(Target(name + ".txt"), new OfdTextExtractor().Extract(package));
    await using (var source = File.OpenRead(Target(name + ".ofd")))
    await using (var pdf = File.Create(Target(name + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(source, pdf);
    for (var page = 1; page <= package.Pages.Count; page++)
    {
        await using var source = File.OpenRead(Target(name + ".ofd")); await using var svg = File.Create(Target($"{name}-{page}.svg"));
        await new OfdToSvgConverter().ConvertAsync(source, svg, page - 1);
    }
}
var native = await Read("graphics");
if (native.Pages.Count != 2 || native.Fonts.Count != 1 || native.Pages.SelectMany(page => page.Elements).Any(element => element is not (OfdTextElement or OfdPathElement))) throw new Exception("Expected only native objects and one font.");
await Save(native, "graphics-roundtrip");
var reread = await Read("graphics-roundtrip");
if (new OfdTextExtractor().Extract(native) != new OfdTextExtractor().Extract(reread)) throw new Exception("Roundtrip text changed.");
await Export("graphics"); await Export("graphics-roundtrip");

// Reuse the deterministic DOCX fixture with explicitly licensed font names.
var docx = Target("licensed-layout.docx"); File.Copy(Path.Combine(root, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx"), docx, true);
using (var archive = ZipFile.Open(docx, ZipArchiveMode.Update))
{
    foreach (var entryName in new[] { "word/document.xml", "word/styles.xml" })
    {
        var entry = archive.GetEntry(entryName)!; XDocument xml;
        using (var input = entry.Open()) xml = XDocument.Load(input);
        XNamespace ns = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        foreach (var run in xml.Descendants(ns + "r"))
        {
            var properties = run.Element(ns + "rPr");
            if (properties is null) { properties = new XElement(ns + "rPr"); run.AddFirst(properties); }
            if (properties.Element(ns + "rFonts") is null) properties.AddFirst(new XElement(ns + "rFonts"));
        }
        foreach (var font in xml.Descendants(ns + "rFonts"))
        {
            font.RemoveAttributes(); foreach (var key in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) font.SetAttributeValue(ns + key, "Noto Sans CJK SC");
        }
        entry.Delete(); using var target = archive.CreateEntry(entryName).Open(); var bytes = Encoding.UTF8.GetBytes(xml.ToString()); target.Write(bytes);
    }
}
string expected;
using (var archive = ZipFile.OpenRead(docx))
using (var source = archive.GetEntry("word/document.xml")!.Open())
    expected = string.Concat(XDocument.Load(source).Descendants(XName.Get("t", "http://schemas.openxmlformats.org/wordprocessingml/2006/main")).Select(node => node.Value));
foreach (var mode in new[] { "native", "default" })
{
    var name = "baseline-" + mode;
    var options = new DocxConversionOptions { FontDirectories = { fonts } };
    if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
    await using (var source = File.OpenRead(docx))
    await using (var target = File.Create(Target(name + ".ofd"))) await new DocxToOfdConverter(options).ConvertAsync(source, target);
    var package = await Read(name);
    var actual = string.Concat(package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>().SelectMany(text => text.Runs).Select(run => run.Text));
    if (actual != expected) throw new Exception("Native DOCX original Unicode text changed.");
    foreach (var font in package.Fonts) { font.Data = fontBytes; font.FileName = "Noto-Regular.ttf"; }
    await Save(package, name); await Export(name);
}
Console.WriteLine("Graphics native-object roundtrip and explicit Native/default DOCX text checks passed; all pages exported to PDF/SVG.");
