using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

var root = Path.GetFullPath(args.Length > 0 ? args[0] : ".");
var output = Path.GetFullPath(args.Length > 1 ? args[1] : Path.Combine(root, "docs/evidence/document-tools/files"));
Directory.CreateDirectory(output);
string PathFor(string name) => Path.Combine(output, name);
var fontDirectory = Path.GetFullPath(args.Length > 2 ? args[2] : Path.Combine(root, "artifacts/document-tools-fonts"));
var fontPath = Path.Combine(fontDirectory, "Ofdrw-CI-NotoSansCJKsc-Regular.ttf");
if (!File.Exists(fontPath)) throw new FileNotFoundException("Install the pinned CI Noto font with scripts/install-ci-fonts.py --directory artifacts/document-tools-fonts.", fontPath);
Directory.CreateDirectory(PathFor("fonts"));
File.Copy(fontPath, PathFor("fonts/Ofdrw-CI-NotoSansCJKsc-Regular.ttf"), true);
File.Copy(Path.Combine(fontDirectory, "Ofdrw-CI-Noto-OFL.txt"), PathFor("fonts/Ofdrw-CI-Noto-OFL.txt"), true);
var inputDocx = PathFor("licensed-layout.docx");
File.Copy(Path.Combine(root, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx"), inputDocx, true);
using (var zip = ZipFile.Open(inputDocx, ZipArchiveMode.Update))
{
    foreach (var path in new[] { "word/document.xml", "word/styles.xml" })
    {
        var entry = zip.GetEntry(path)!; XDocument xml;
        using (var stream = entry.Open()) xml = XDocument.Load(stream);
        var wordNs = xml.Root!.Name.Namespace;
        foreach (var run in xml.Descendants(wordNs + "r"))
        {
            var properties = run.Element(wordNs + "rPr");
            if (properties is null) { properties = new XElement(wordNs + "rPr"); run.AddFirst(properties); }
            if (properties.Element(wordNs + "rFonts") is null) properties.AddFirst(new XElement(wordNs + "rFonts"));
        }
        foreach (var fonts in xml.Descendants(wordNs + "rFonts"))
        {
            fonts.RemoveAttributes();
            foreach (var name in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) fonts.SetAttributeValue(wordNs + name, "Noto Sans CJK SC");
        }
        entry.Delete(); using var outputEntry = zip.CreateEntry(path).Open(); var bytes = Bytes(xml); outputEntry.Write(bytes);
    }
}
var fontBytes = File.ReadAllBytes(fontPath);
foreach (var mode in new[] { "native", "default" })
{
    await using var input = File.OpenRead(inputDocx); await using var target = File.Create(PathFor("baseline-" + mode + ".ofd"));
    var options = new DocxConversionOptions { FontDirectories = { fontDirectory } };
    if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
    var converter = new DocxToOfdConverter(options);
    await converter.ConvertAsync(input, target);
    await target.DisposeAsync();
    // Native's viewer-local CJK contract is name-only. Make this review fixture
    // self-contained by embedding the same licensed face used for measurement.
    var embedded = await Read("baseline-" + mode);
    foreach (var font in embedded.Fonts)
    {
        font.Data = fontBytes; font.FileName = "Ofdrw-CI-NotoSansCJKsc-Regular.ttf";
    }
    await Save(embedded, "baseline-" + mode);
}
var source = await Read("baseline-native");
if (source.Pages.Count < 2) throw new Exception("Sample requires multiple pages.");
source.Attachments.Add(new OfdAttachment { Name = "public-note.txt", MediaType = "text/plain", Data = Encoding.UTF8.GetBytes("Synthetic public sample attachment.\n") });
await Save(source, "rich");
var ns = XNamespace.Get(source.Options.Namespace);
Mutate("rich", entries =>
{
    var document = Xml(entries["Doc_0/Document.xml"]);
    var first = document.Root!.Element(ns + "Pages")!.Elements().First();
    var pagePath = "Doc_0/" + first.Attribute("BaseLoc")!.Value;
    var page = Xml(entries[pagePath]);
    page.Root!.Add(new XElement(ns + "Template", new XAttribute("TemplateID", "999001"), new XAttribute("ZOrder", "Background")));
    entries[pagePath] = Bytes(page);
    document.Root.Element(ns + "CommonData")!.Add(new XElement(ns + "TemplatePage", new XAttribute("ID", "999001"), new XAttribute("BaseLoc", "Templates/Content.xml")));
    document.Root.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml"));
    entries["Doc_0/Document.xml"] = Bytes(document);
    var font = source.Fonts.First().Id;
    entries["Doc_0/Templates/Content.xml"] = Encoding.UTF8.GetBytes($"<Page xmlns='{ns}'><Content><Layer ID='999002' Type='Background'><PathObject ID='999003' Boundary='8 8 3 278' Fill='true' Stroke='false'><FillColor Value='220 230 250'/><AbbreviatedData>M 0 0 L 3 0 L 3 278 L 0 278 C</AbbreviatedData></PathObject><PathObject ID='999011' Boundary='18 170 80 20' Fill='true' Stroke='false'><FillColor Value='190 215 250'/><AbbreviatedData>M 0 0 L 80 0 L 80 20 L 0 20 C</AbbreviatedData></PathObject><TextObject ID='999012' Font='{font}' Size='4' Boundary='23 175 75 10'><TextCode X='0' Y='4'>UNDER LAYER 下层</TextCode></TextObject><TextObject ID='999004' Font='{font}' Size='3' Boundary='15 277 120 10'><TextCode X='0' Y='3'>TEMPLATE 模板</TextCode></TextObject></Layer></Content></Page>");
    entries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{first.Attribute("ID")!.Value}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
    entries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='999005' Type='Stamp' Visible='true'><Appearance ID='999006' Boundary='135 10 60 12' CTM='1 0 0 1 0 0'><PageBlock ID='999008'><TextObject ID='999007' Font='{font}' Size='3' Boundary='0 0 60 12'><FillColor Value='30 100 180'/><TextCode X='0' Y='3'>NOTE 注释</TextCode></TextObject><TextObject ID='999013' Name='hidden-text' Visible='false' Font='{font}' Size='3' Boundary='0 0 60 12'><TextCode X='0' Y='3'>HIDDEN</TextCode></TextObject><PathObject ID='999014' Name='hidden-path' Visible='false' Boundary='0 0 60 12' Fill='true' Stroke='false'><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 60 0 L 60 12 L 0 12 C</AbbreviatedData></PathObject></PageBlock></Appearance></Annot></PageAnnot>");
});
source = await Read("rich");
File.Copy(PathFor("rich.ofd"), PathFor("annotation-metadata.ofd"), true);
Mutate("annotation-metadata", entries =>
{
    var xml = Xml(entries["Doc_0/Annots/Page.xml"]);
    xml.Descendants(ns + "Appearance").Single().SetAttributeValue(XNamespace.Get("urn:vendor:metadata") + "Style", "review-fixture");
    entries["Doc_0/Annots/Page.xml"] = Bytes(xml);
});
try
{
    OfdDocumentMixer.Mix([new(await Read("annotation-metadata"), 0)]);
    throw new Exception("Mix must reject unmodeled annotation metadata.");
}
catch (NotSupportedException) { }
var mark = File.ReadAllBytes(Path.Combine(root, "e2e/Ofdrw.Net.DocumentTools.E2E/mark.png"));
await using (var input = File.OpenRead(PathFor("rich.ofd")))
await using (var target = File.Create(PathFor("signed.ofd")))
    await new OfdSignatureService().SignAsync(input, target, new SyntheticProvider(mark));
Mutate("signed", entries =>
{
    var description = entries.Keys.Single(path => path.EndsWith("/Signature.xml"));
    var signature = Xml(entries[description]);
    var signNs = signature.Root!.Name.Namespace;
    signature.Root.Element(signNs + "SignedInfo")!.Add(new XElement(signNs + "StampAnnot", new XAttribute("ID", "999010"), new XAttribute("PageRef", source.Pages[0].Id!), new XAttribute("Boundary", "175 240 15 15")));
    entries[description] = Bytes(signature);
});
source = await Read("signed");
OfdWatermark.AddText(source, [0], "DRAFT 草稿", new OfdWatermarkOptions { XMillimeters = 60, YMillimeters = 250, WidthMillimeters = 100, HeightMillimeters = 16, LayerId = "watermark", LayerType = "Foreground" }, fontName: source.Fonts.First().FontName, fontSizeMillimeters: 7);
OfdWatermark.AddImage(source, [0], mark, "image/png", new OfdWatermarkOptions { XMillimeters = 175, YMillimeters = 215, WidthMillimeters = 15, HeightMillimeters = 15, LayerId = "watermark", LayerType = "Foreground" });
await Save(source, "watermark");
await Save(OfdDocumentMerger.Merge([source, new OfdDocumentPackage { Fonts = { source.Fonts.First() }, Pages = { new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80, Elements = { new OfdTextElement { Text = "MERGE END", FontName = source.Fonts.First().FontName, FontResourceId = source.Fonts.First().Id, XMillimeters = 15, YMillimeters = 20 } } } } }]), "watermark-merged");
var signed = await Read("signed");
await Save(OfdDocumentSplitter.Split(signed, [1, 0]), "split");
var overlay = new OfdDocumentPackage();
overlay.Fonts.Add(signed.Fonts.First());
var overlayPage = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 250 };
overlayPage.Elements.Add(new OfdPathElement { LayerId = "cover", XMillimeters = 18, YMillimeters = 170, WidthMillimeters = 80, HeightMillimeters = 20, Fill = true, Stroke = false, FillColor = new OfdColor(230, 245, 220), AbbreviatedData = "M 0 0 L 80 0 L 80 20 L 0 20 C" });
overlayPage.Elements.Add(new OfdTextElement { LayerId = "cover", Text = "TOP LAYER 上层", FontName = signed.Fonts.First().FontName, FontResourceId = signed.Fonts.First().Id, XMillimeters = 23, YMillimeters = 175, WidthMillimeters = 75, HeightMillimeters = 10, FontSizeMillimeters = 4, FillColor = new OfdColor(40, 120, 30) });
overlay.Pages.Add(overlayPage);
await Save(overlay, "overlay");
File.Copy(PathFor("signed.ofd"), PathFor("scoped-annotations-input.ofd"), true);
Mutate("scoped-annotations-input", entries =>
{
    var list = Xml(entries["Doc_0/Annots/Annotations.xml"]);
    list.Root!.Add(new XElement(ns + "Page", new XAttribute("PageID", signed.Pages[1].Id!), new XElement(XNamespace.Get("urn:vendor") + "FileLoc", "unmodeled.bin")));
    entries["Doc_0/Annots/Annotations.xml"] = Bytes(list);
});
var scopedInput = await Read("scoped-annotations-input");
try { OfdDocumentMixer.Mix([new(scopedInput, 1)]); throw new Exception("Affected page must reject unmodeled annotation records."); }
catch (NotSupportedException) { }
await Save(OfdDocumentMixer.Mix([new(scopedInput, 0), new(overlay, 0)]), "mix");
foreach (var negative in new[] { "annotation-unknown-media", "annotation-missing-payload" })
{
    File.Copy(PathFor("rich.ofd"), PathFor(negative + ".ofd"), true);
    Mutate(negative, entries =>
    {
        var annotation = Xml(entries["Doc_0/Annots/Page.xml"]);
        var appearance = annotation.Descendants(ns + "Appearance").Single(); appearance.RemoveNodes();
        appearance.Add(new XElement(ns + "ImageObject", new XAttribute("ID", "999015"), new XAttribute("ResourceID", "999016"), new XAttribute("Boundary", "0 0 60 12")));
        entries["Doc_0/Annots/Page.xml"] = Bytes(annotation);
        if (negative == "annotation-missing-payload")
        {
            var resources = Xml(entries["Doc_0/PublicRes.xml"]);
            var media = resources.Root!.Element(ns + "MultiMedias"); if (media is null) { media = new XElement(ns + "MultiMedias"); resources.Root.Add(media); }
            media.Add(new XElement(ns + "MultiMedia", new XAttribute("ID", "999016"), new XAttribute("Type", "Image"), new XAttribute("Format", "PNG"), new XElement(ns + "MediaFile", "missing-annotation.png")));
            entries["Doc_0/PublicRes.xml"] = Bytes(resources);
        }
    });
    using var input = File.OpenRead(PathFor(negative + ".ofd")); using var target = new MemoryStream();
    try { await new OfdToPdfConverter().ConvertAsync(input, target); throw new Exception("Missing annotation image must not export a partial PDF."); }
    catch (NotSupportedException exception) { if (target.Length != 0) throw new Exception("Rejected export wrote partial PDF bytes."); File.WriteAllText(PathFor(negative + ".rejection.txt"), exception.Message); }
}
await using (var input = File.OpenRead(PathFor("signed.ofd")))
await using (var target = File.Create(PathFor("clean.ofd"))) await OfdSignatureCleaner.CleanAsync(input, target);
var clipped = new OfdDocumentPackage();
clipped.Fonts.Add(source.Fonts.First(font => !font.Bold && !font.Italic));
clipped.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100,
    Elements = { new OfdImageElement { Data = mark, Alpha = 0, WidthMillimeters = 1, HeightMillimeters = 1 } } });
await Save(clipped, "annotation-clipped");
var clippedRead = await Read("annotation-clipped");
Mutate("annotation-clipped", entries =>
{
    var document = Xml(entries["Doc_0/Document.xml"]);
    document.Root!.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml")); entries["Doc_0/Document.xml"] = Bytes(document);
    var font = clippedRead.Fonts.Single().Id; var media = clippedRead.Pages[0].Elements.OfType<OfdImageElement>().Single().ResourceId;
    entries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{clippedRead.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
    entries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='800'><Appearance Boundary='10 10 20 10' CTM='1 0 0.5 1 0 0'><PageBlock ID='801'><ImageObject ID='802' ResourceID='{media}' Boundary='0 0 200 10'/></PageBlock></Appearance></Annot><Annot ID='810'><Appearance Boundary='50 15 35 15' CTM='0 1 -1 0 0 0'><PageBlock ID='811'><TextObject ID='812' Font='{font}' Size='4' Boundary='0 0 30 8'><TextCode X='0' Y='4'>ROTATE</TextCode></TextObject></PageBlock></Appearance></Annot><Annot ID='820'><Appearance Boundary='50 60 20 10' CTM='1.5 0 0 1.5 0 0'><TextObject ID='821' Font='{font}' Size='4' Boundary='0 0 20 8'><TextCode X='0' Y='4'>SCALE</TextCode></TextObject></Appearance></Annot></PageAnnot>");
});
await Save(OfdDocumentMixer.Mix([new(await Read("annotation-clipped"), 0)]), "annotation-clipped-mix");
OfdDocumentPackage ItalicSample(bool embed, double[]? matrix)
{
    var package = new OfdDocumentPackage();
    package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "OFD Example Noto", Data = embed ? fontBytes : Array.Empty<byte>() });
    package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 60, Elements = {
        new OfdTextElement { Text = "ABCD Test italic", FontName = "OFD Example Noto", FontResourceId = "10", Italic = true,
            FontSizeMillimeters = 6, XMillimeters = 20, YMillimeters = 20, WidthMillimeters = 70, HeightMillimeters = 12, Transform = matrix } } });
    return package;
}
await Save(ItalicSample(false, null), "italic-marked");
var markedItalic = await Read("italic-marked");
foreach (var font in markedItalic.Fonts) { font.Data = fontBytes; font.FileName = "Noto.ttf"; }
await Save(markedItalic, "italic-marked");
await Save(ItalicSample(true, new double[] { 1, 0, 0, 1, 0, 0 }), "italic-control");
await Save(ItalicSample(true, new double[] { 1, 0, -0.2, 1, 1.2, 0 }), "italic-user-matrix");
foreach (var name in new[] { "baseline-native", "baseline-default", "rich", "annotation-metadata", "signed", "watermark", "watermark-merged", "split", "mix", "clean", "overlay", "annotation-clipped", "annotation-clipped-mix", "italic-marked", "italic-control", "italic-user-matrix" })
{
    await using (var input = File.OpenRead(PathFor(name + ".ofd")))
    await using (var target = File.Create(PathFor(name + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(input, target);
    var package = await Read(name);
    for (var page = 0; page < package.Pages.Count; page++)
    {
        await using var input = File.OpenRead(PathFor(name + ".ofd")); await using var target = File.Create(PathFor($"{name}-{page + 1}.svg"));
        await new OfdToSvgConverter().ConvertAsync(input, target, page);
    }
    File.WriteAllText(PathFor(name + ".txt"), new OfdTextExtractor().Extract(package, includeTemplates: true));
}
await using (var input = File.OpenRead(PathFor("clean.ofd")))
{
    var report = await new OfdSignatureVerifier().VerifyAsync(input);
    if (report.HasSignatureDeclarations || report.HasSignatures) throw new Exception("Clean output still has signature declarations.");
}
Console.WriteLine($"Generated document-tool API sample matrix: {output}");

async Task<OfdDocumentPackage> Read(string name) { await using var input = File.OpenRead(PathFor(name + ".ofd")); return await new OfdReader().ReadAsync(input); }
async Task Save(OfdDocumentPackage package, string name) { await using var target = File.Create(PathFor(name + ".ofd")); await new OfdPackageWriter().WriteAsync(package, target); }
static XDocument Xml(byte[] data) { using var input = new MemoryStream(data); return XDocument.Load(input); }
static byte[] Bytes(XDocument xml) => Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting));
void Mutate(string name, Action<Dictionary<string, byte[]>> action)
{
    Dictionary<string, byte[]> entries;
    using (var zip = ZipFile.OpenRead(PathFor(name + ".ofd"))) entries = zip.Entries.ToDictionary(entry => entry.FullName, entry => { using var input = entry.Open(); using var data = new MemoryStream(); input.CopyTo(data); return data.ToArray(); });
    action(entries);
    using var outputZip = ZipFile.Open(PathFor(name + ".ofd") + ".tmp", ZipArchiveMode.Create);
    foreach (var entry in entries) { using var target = outputZip.CreateEntry(entry.Key).Open(); target.Write(entry.Value); }
    outputZip.Dispose(); File.Move(PathFor(name + ".ofd") + ".tmp", PathFor(name + ".ofd"), true);
}
sealed class SyntheticProvider(byte[] mark) : IOfdSignatureProvider
{
    public string ProviderName => "Synthetic sample only"; public string Company => "Test"; public string Version => "1";
    public string SignatureMethod => "urn:ofdrw:test-only"; public string SignatureType => "Seal"; public byte[] SealData => mark;
    public Task<byte[]> SignAsync(byte[] xml, string properties, CancellationToken token = default) => Task.FromResult(Encoding.UTF8.GetBytes("synthetic-value-no-cryptographic-validity"));
}
