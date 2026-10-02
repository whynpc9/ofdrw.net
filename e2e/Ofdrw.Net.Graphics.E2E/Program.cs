using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Layout.Graphics;
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

// Native name-only font declarations deliberately have no FontFile. The PDF
// viewer is explicitly supplied the same licensed font, never a system font.
PdfFontRegistry.RegisterFont("Noto Sans CJK SC", fontBytes);
var nameOnly = new OfdDocumentPackage();
nameOnly.Fonts.Add(new OfdFontResource { Id = "name-only", FontName = "Noto Sans CJK SC" });
var namePage = new OfdPage { WidthMillimeters = 148, HeightMillimeters = 210 }; nameOnly.Pages.Add(namePage);
var nameGraphics = new OfdGraphics(nameOnly, namePage);
var nameBrush = new OfdBrush(new OfdColor(30, 93, 166));
var namePen = new OfdPen(new OfdColor(170, 185, 200), 0.15);
OfdFont NameFont(double size, bool italic = false) => new("Noto Sans CJK SC", size, italic: italic, resourceId: "name-only");
nameGraphics.DrawString("Name-only 原生斜体与用户矩阵", NameFont(5.2), nameBrush, 12, 22);
nameGraphics.DrawLine(namePen, 12, 42, 135, 42);
nameGraphics.DrawString("正体 Upright control", NameFont(5), nameBrush, 12, 42);
nameGraphics.DrawLine(namePen, 12, 62, 135, 62);
nameGraphics.DrawString("斜体 Italic baseline", NameFont(5, true), nameBrush, 12, 62);
nameGraphics.Save(); nameGraphics.Translate(25, 88); nameGraphics.Rotate(-10); nameGraphics.Scale(1.2, 0.9);
nameGraphics.MultiplyTransform(new OfdMatrix(1, 0, 0.15, 1, 0, 0));
nameGraphics.DrawLine(namePen, 0, 10, 80, 10);
nameGraphics.DrawString("组合 CTM + italic", NameFont(5.333, true), nameBrush, 0, 10);
nameGraphics.Restore();
nameGraphics.Save(); nameGraphics.IntersectClip(new OfdGraphicsPath().AddRectangle(15, 120, 115, 25));
nameGraphics.DrawLine(namePen, 15, 135, 130, 135);
nameGraphics.DrawString("裁剪基线 Clipped italic", NameFont(5.333, true), nameBrush, 15, 135);
nameGraphics.Restore();
nameGraphics.DrawString("Native CTM = user M * generated F; PDF/SVG factor once", NameFont(2.7), nameBrush, 12, 176);
nameGraphics.DrawString("OFD keeps font name only; exported PDF uses pinned OFL bytes", NameFont(2.5), nameBrush, 12, 190);
await Save(nameOnly, "graphics-name-only");
var nameRoundtrip = await Read("graphics-name-only");
if (nameRoundtrip.Fonts.Any(font => font.Data.Length != 0)) throw new Exception("Name-only OFD unexpectedly embedded a font.");
foreach (var text in nameRoundtrip.Pages[0].Elements.OfType<OfdTextElement>().Where(text => text.Italic))
    if (!XElement.Parse(text.SourceXml!).Attributes().Any(attribute => attribute.Name.LocalName == "FauxItalicMatrixV1")) throw new Exception("Native name-only italic has no generated factor.");
await Export("graphics-name-only");

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
