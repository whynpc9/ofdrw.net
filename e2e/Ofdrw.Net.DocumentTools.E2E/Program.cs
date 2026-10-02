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
using Ofdrw.Net.Packaging.Archive;
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
    xml.Root!.SetAttributeValue(XNamespace.Get("urn:vendor:metadata") + "Payload", "Attachs/public.bin");
    xml.Descendants(ns + "Annot").Single().SetAttributeValue(XNamespace.Get("urn:vendor:metadata") + "Payload", "Attachs/public.bin");
    xml.Descendants(ns + "Appearance").Single().SetAttributeValue(XNamespace.Get("urn:vendor:metadata") + "Style", "review-fixture");
    entries["Doc_0/Annots/Page.xml"] = Bytes(xml);
    var vendor = XNamespace.Get("urn:vendor:metadata"); var document = Xml(entries["Doc_0/Document.xml"]);
    document.Root!.Element(ns + "Annotations")!.SetAttributeValue(vendor + "Payload", "Attachs/public.bin");
    var template = document.Root.Element(ns + "CommonData")!.Element(ns + "TemplatePage")!;
    template.SetAttributeValue(vendor + "Payload", "Attachs/public.bin"); template.Add(new XElement(vendor + "Extra", "keep"));
    entries["Doc_0/Document.xml"] = Bytes(document);
    var index = Xml(entries["Doc_0/Annots/Annotations.xml"]); var record = index.Root!.Element(ns + "Page")!;
    index.Root.SetAttributeValue(vendor + "Payload", "../Attachs/public.bin");
    record.Element(ns + "FileLoc")!.SetAttributeValue(vendor + "Payload", "../Attachs/public.bin");
    record.SetAttributeValue(vendor + "Payload", "../Attachs/public.bin"); record.Add(new XElement(vendor + "Extra", "keep"));
    entries["Doc_0/Annots/Annotations.xml"] = Bytes(index);
});
try
{
    OfdDocumentMixer.Mix([new(await Read("annotation-metadata"), 0)]);
    throw new Exception("Mix must reject unmodeled annotation metadata.");
}
catch (NotSupportedException) { }
var boxModel = new OfdDocumentPackage(); boxModel.Fonts.Add(source.Fonts.First());
boxModel.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80, Elements = { new OfdTextElement { Text = "BOX 100 x 80", FontName = source.Fonts.First().FontName, FontResourceId = source.Fonts.First().Id, XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 10, FontSizeMillimeters = 4 } } });
await Save(boxModel, "box-whitespace-input");
Mutate("box-whitespace-input", entries =>
{
    var document = Xml(entries["Doc_0/Document.xml"]); document.Descendants(ns + "PageArea").Single().Element(ns + "PhysicalBox")!.Value = "0 0 210 297"; entries["Doc_0/Document.xml"] = Bytes(document);
    var pagePath = entries.Keys.Single(path => path.EndsWith("/Content.xml")); var page = Xml(entries[pagePath]); page.Root!.Element(ns + "Area")!.Element(ns + "PhysicalBox")!.Value = "0\t0\n100\r80"; entries[pagePath] = Bytes(page);
});
await Save(OfdDocumentMixer.Mix([new(await Read("box-whitespace-input"), 0)]), "box-whitespace");
foreach (var inherited in new[] { false, true })
{
    var name = inherited ? "box-inexact-inherited" : "box-inexact-page";
    File.Copy(PathFor("box-whitespace.ofd"), PathFor(name + ".ofd"), true);
    Mutate(name, entries =>
    {
        var path = entries.Keys.Single(path => path.EndsWith("/Content.xml")); var page = Xml(entries[path]);
        if (inherited)
        {
            page.Root!.Element(ns + "Area")!.Remove(); var document = Xml(entries["Doc_0/Document.xml"]);
            document.Descendants(ns + "PageArea").Single().Element(ns + "PhysicalBox")!.Value = "0 0 100 80 SECRET"; entries["Doc_0/Document.xml"] = Bytes(document);
        }
        else page.Root!.Element(ns + "Area")!.Element(ns + "PhysicalBox")!.Value = "0 0 100 80 SECRET";
        entries[path] = Bytes(page);
    });
    var read = await Read(name); if (read.Pages[0].WidthMillimeters != 100 || read.Pages[0].HeightMillimeters != 80) throw new Exception("Inexact display changed page geometry.");
    try { OfdDocumentMixer.Mix([new(read, 0)]); throw new Exception("Inexact box must refuse Mix."); }
    catch (NotSupportedException exception) { File.WriteAllText(PathFor(name + ".rejection.txt"), exception.Message); }
}
foreach (var primitive in new[] { false, true })
{
    var name = primitive ? "annotation-inexact-primitive" : "annotation-inexact-appearance";
    File.Copy(PathFor("box-whitespace.ofd"), PathFor(name + ".ofd"), true); var read = await Read(name);
    Mutate(name, entries =>
    {
        var document = Xml(entries["Doc_0/Document.xml"]); document.Root!.Add(new XElement(ns + "Annotations", "Annots/Annotations.xml")); entries["Doc_0/Document.xml"] = Bytes(document);
        entries["Doc_0/Annots/Annotations.xml"] = Encoding.UTF8.GetBytes($"<Annotations xmlns='{ns}'><Page PageID='{read.Pages[0].Id}'><FileLoc>Page.xml</FileLoc></Page></Annotations>");
        var appearanceBox = primitive ? "10 30 20 20" : "10 30 20 20 SECRET"; var primitiveBox = primitive ? "0 0 10 10 SECRET" : "0 0 10 10";
        entries["Doc_0/Annots/Page.xml"] = Encoding.UTF8.GetBytes($"<PageAnnot xmlns='{ns}'><Annot ID='900'><Appearance Boundary='{appearanceBox}'><PathObject ID='901' Boundary='{primitiveBox}' Fill='true' Stroke='false'><FillColor Value='255 0 0'/><AbbreviatedData>M 0 0 L 10 0 L 10 10 L 0 10 C</AbbreviatedData></PathObject></Appearance></Annot></PageAnnot>");
    });
    read = await Read(name);
    try { OfdDocumentMixer.Mix([new(read, 0)]); throw new Exception("Inexact annotation must refuse Mix."); }
    catch (NotSupportedException exception) { File.WriteAllText(PathFor(name + ".rejection.txt"), exception.Message); }
}
File.Copy(PathFor("box-whitespace.ofd"), PathFor("tiny-page-input.ofd"), true);
Mutate("tiny-page-input", entries => { var path = entries.Keys.Single(path => path.EndsWith("/Content.xml")); var page = Xml(entries[path]); page.Descendants(ns + "PhysicalBox").Single().Value = "0 0 0.0004 80"; entries[path] = Bytes(page); });
var tiny = await Read("tiny-page-input");
foreach (var operation in new[] { "split", "mix", "watermark", "save" })
{
    using var target = new MemoryStream(); var count = tiny.Pages[0].Elements.Count;
    try
    {
        if (operation == "split") OfdDocumentSplitter.Split(tiny, [0]);
        if (operation == "mix") OfdDocumentMixer.Mix([new(tiny, 0)]);
        if (operation == "watermark") OfdWatermark.AddText(tiny, [0], "DRAFT");
        if (operation == "save") await new OfdPackageWriter().WriteAsync(tiny, target);
        throw new Exception("Tiny side must refuse rewrite.");
    }
    catch (NotSupportedException exception) { if (target.Length != 0 || count != tiny.Pages[0].Elements.Count) throw new Exception("Tiny rejection mutated content/output."); File.WriteAllText(PathFor("tiny-page-" + operation + ".rejection.txt"), exception.Message); }
}
var withoutArchive = new OfdDocumentPackage(); withoutArchive.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
withoutArchive.Pages[0].PreservedPageElements.Add($"<Actions xmlns='{ns}'><Action Event='CLICK'/></Actions>");
try { OfdDocumentSplitter.Split(withoutArchive, [0]); throw new Exception("Unarchived page XML must refuse Split."); }
catch (NotSupportedException exception) { File.WriteAllText(PathFor("unarchived-page-xml.rejection.txt"), exception.Message); }
var mark = File.ReadAllBytes(Path.Combine(root, "e2e/Ofdrw.Net.DocumentTools.E2E/mark.png"));
var blank = new OfdDocumentPackage(); blank.Fonts.Add(source.Fonts.First()); blank.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80 });
await Save(blank, "blank-watermark-input"); blank = await Read("blank-watermark-input");
OfdWatermark.AddText(blank, [0], "BLANK PAGE WATERMARK", new OfdWatermarkOptions { XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 10, LayerId = "watermark" }, fontName: blank.Fonts.First().FontName, fontSizeMillimeters: 4);
OfdWatermark.AddImage(blank, [0], mark, "image/png", new OfdWatermarkOptions { XMillimeters = 40, YMillimeters = 35, WidthMillimeters = 20, HeightMillimeters = 20, LayerId = "watermark" });
await Save(blank, "blank-watermark");

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
File.Copy(PathFor("signed.ofd"), PathFor("binary-xml-input.ofd"), true);
Mutate("binary-xml-input", entries =>
{
    var description = entries.Keys.Single(path => path.EndsWith("/Signature.xml")); var signature = Xml(entries[description]); var location = signature.Root!.Element(signature.Root.Name.Namespace + "SignedValue")!;
    var oldPath = location.Value.TrimStart('/'); if (!entries.Remove(oldPath)) throw new Exception("Generated synthetic SignedValue entry missing.");
    var newPath = description[..description.LastIndexOf('/')] + "/SignedValue.xml"; location.Value = "/" + newPath;
    entries[newPath] = [0, 255, 13, 7]; entries[description] = Bytes(signature);
    var attachments = Xml(entries["Doc_0/Attachs/Attachments.xml"]); attachments.Root!.Add(new XElement(ns + "Attachment", new XAttribute("ID", "binary-public"), new XAttribute("Name", "attachment.xml"), new XAttribute("Format", "application/octet-stream"), new XElement(ns + "FileLoc", "attachment.xml")));
    entries["Doc_0/Attachs/Attachments.xml"] = Bytes(attachments); entries["Doc_0/Attachs/attachment.xml"] = [0, 254, 12, 6];
});
await Save(OfdDocumentSplitter.Split(await Read("binary-xml-input"), [1]), "binary-xml-split");
File.Copy(PathFor("rich.ofd"), PathFor("page-wrapper-metadata.ofd"), true);
Mutate("page-wrapper-metadata", entries =>
{
    var document = Xml(entries["Doc_0/Document.xml"]); var first = document.Root!.Element(ns + "Pages")!.Elements().First(); var path = "Doc_0/" + first.Attribute("BaseLoc")!.Value; var page = Xml(entries[path]);
    foreach (var node in page.Root!.DescendantsAndSelf().Where(node => node.Name.LocalName is "Page" or "Area" or "Content" or "Layer")) node.SetAttributeValue(XNamespace.Get("urn:vendor:wrapper") + "Payload", "public.bin");
    entries[path] = Bytes(page); entries["Doc_0/public.bin"] = Encoding.UTF8.GetBytes("public synthetic wrapper payload");
});
var wrapper = await Read("page-wrapper-metadata");
foreach (var operation in new[] { "split", "save", "watermark" })
{
    using var target = new MemoryStream();
    try
    {
        if (operation == "split") OfdDocumentSplitter.Split(wrapper, [0]);
        if (operation == "save") await new OfdPackageWriter().WriteAsync(wrapper, target);
        if (operation == "watermark") OfdWatermark.AddText(wrapper, [1], "DRAFT");
        throw new Exception("Extended page wrapper must not be discarded.");
    }
    catch (NotSupportedException exception) { if (target.Length != 0) throw new Exception("Rejected rewrite emitted bytes."); File.WriteAllText(PathFor("page-wrapper-metadata-" + operation + ".rejection.txt"), exception.Message); }
}
File.Copy(PathFor("rich.ofd"), PathFor("annotation-processing-instruction.ofd"), true);
Mutate("annotation-processing-instruction", entries =>
{
    var xml = Xml(entries["Doc_0/Annots/Page.xml"]); xml.Descendants(ns + "TextCode").First().Add(new XProcessingInstruction("vendor", "public.bin")); entries["Doc_0/Annots/Page.xml"] = Bytes(xml); entries["Doc_0/public.bin"] = [42];
});
foreach (var format in new[] { "pdf", "svg" })
{
    using var input = File.OpenRead(PathFor("annotation-processing-instruction.ofd")); using var target = new MemoryStream();
    try { if (format == "pdf") await new OfdToPdfConverter().ConvertAsync(input, target); else await new OfdToSvgConverter().ConvertAsync(input, target); throw new Exception("Opaque drawing PI must reject export."); }
    catch (NotSupportedException exception) { if (target.Length != 0) throw new Exception("Rejected PI export wrote bytes."); File.WriteAllText(PathFor("annotation-processing-instruction-" + format + ".rejection.txt"), exception.Message); }
}
source = await Read("signed");
OfdWatermark.AddText(source, [0], "DRAFT 草稿", new OfdWatermarkOptions { XMillimeters = 60, YMillimeters = 250, WidthMillimeters = 100, HeightMillimeters = 16, LayerId = "watermark", LayerType = "Foreground" }, fontName: source.Fonts.First().FontName, fontSizeMillimeters: 7);
OfdWatermark.AddImage(source, [0], mark, "image/png", new OfdWatermarkOptions { XMillimeters = 175, YMillimeters = 215, WidthMillimeters = 15, HeightMillimeters = 15, LayerId = "watermark", LayerType = "Foreground" });
await Save(source, "watermark");
await Save(OfdDocumentMerger.Merge([source, new OfdDocumentPackage { Fonts = { source.Fonts.First() }, Pages = { new OfdPage { WidthMillimeters = 100, HeightMillimeters = 80, Elements = { new OfdTextElement { Text = "MERGE END", FontName = source.Fonts.First().FontName, FontResourceId = source.Fonts.First().Id, XMillimeters = 15, YMillimeters = 20 } } } } }]), "watermark-merged");
var signed = await Read("signed");
File.Copy(PathFor("watermark.ofd"), PathFor("resource-suffix-input.ofd"), true);
Mutate("resource-suffix-input", entries =>
{
    var documents = entries.Where(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).ToDictionary(pair => pair.Key, pair => Xml(pair.Value));
    var document = documents["Doc_0/Document.xml"];
    var declarations = document.Root!.Element(ns + "CommonData")!.Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes").ToList();
    var reservedResources = declarations.Select(declaration => documents["Doc_0/" + declaration.Value].Descendants(ns + (declaration.Name.LocalName == "PublicRes" ? "Font" : "MultiMedia")).First()).ToHashSet();
    var maximum = documents.Values.SelectMany(xml => xml.Descendants().Attributes("ID")).Where(attribute => !reservedResources.Contains(attribute.Parent!))
        .Select(attribute => long.TryParse(attribute.Value, out var id) ? id : 0).Max();
    // Real readers can accept a stale document hint. IDs present only in these
    // descriptors still reserve the next otherwise available object numbers.
    foreach (var declaration in declarations)
    {
        var resource = documents["Doc_0/" + declaration.Value].Descendants(ns + (declaration.Name.LocalName == "PublicRes" ? "Font" : "MultiMedia")).First();
        var oldId = resource.Attribute("ID")!.Value; var id = (++maximum).ToString(System.Globalization.CultureInfo.InvariantCulture); resource.SetAttributeValue("ID", id);
        var referenceName = declaration.Name.LocalName == "PublicRes" ? "Font" : "ResourceID";
        foreach (var reference in documents.Values.SelectMany(xml => xml.Descendants().Attributes(referenceName)).Where(attribute => attribute.Value == oldId)) reference.Value = id;
    }
    document.Root.Element(ns + "CommonData")!.Element(ns + "MaxUnitID")!.Value = (maximum - 2).ToString(System.Globalization.CultureInfo.InvariantCulture);
    foreach (var pair in documents) entries[pair.Key] = Bytes(pair.Value);
    foreach (var declaration in document.Root!.Element(ns + "CommonData")!.Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes"))
    {
        var oldPath = "Doc_0/" + declaration.Value;
        declaration.Value = declaration.Name.LocalName == "PublicRes" ? "PublicResources.dat" : "ImageResources.bin";
        entries["Doc_0/" + declaration.Value] = entries[oldPath]; entries.Remove(oldPath);
    }
    entries["Doc_0/Document.xml"] = Bytes(document);
});
var resourceSuffix = await Read("resource-suffix-input");
OfdWatermark.AddText(resourceSuffix, [0], "RESOURCE SUFFIX", new OfdWatermarkOptions { XMillimeters = 60, YMillimeters = 230 },
    fontName: resourceSuffix.Fonts.First().FontName, fontSizeMillimeters: 7);
await Save(resourceSuffix, "watermark-resource-suffix");
File.Copy(PathFor("rich.ofd"), PathFor("vendor-annotations-input.ofd"), true);
Mutate("vendor-annotations-input", entries =>
{
    var document = Xml(entries["Doc_0/Document.xml"]); var declaration = document.Root!.Element(ns + "Annotations")!; var vendor = XNamespace.Get("urn:vendor");
    declaration.AddBeforeSelf(new XElement(vendor + "Annotations", "../../../external"), new XElement(vendor + "Annotations", declaration.Value));
    var template = document.Root.Element(ns + "CommonData")!.Element(ns + "TemplatePage")!;
    template.AddBeforeSelf(new XElement(vendor + "TemplatePage", new XAttribute("ID", template.Attribute("ID")!.Value), new XAttribute("BaseLoc", "../../../external")));
    entries["Doc_0/Document.xml"] = Bytes(document);
    var pagePath = source.Pages[0].SourceEntryPath!; var page = Xml(entries[pagePath]); var reference = page.Root!.Element(ns + "Template")!;
    reference.AddBeforeSelf(new XElement(vendor + "Template", new XAttribute("TemplateID", reference.Attribute("TemplateID")!.Value), new XAttribute("ZOrder", "Foreground")));
    entries[pagePath] = Bytes(page);
});
var vendorAnnotations = await Read("vendor-annotations-input");
if (!vendorAnnotations.Pages[0].AnnotationAppearances.OfType<OfdTextElement>().Any(text => text.Text == "NOTE 注释")) throw new Exception("Vendor metadata must not suppress standard annotation artwork.");
try { OfdDocumentMixer.Mix([new(vendorAnnotations, 0)]); throw new Exception("Vendor annotation metadata must prevent unsafe Mix."); }
catch (NotSupportedException exception) { File.WriteAllText(PathFor("vendor-annotations-input.rejection.txt"), exception.Message); }
await Save(vendorAnnotations, "vendor-annotations-roundtrip");
var tagged = await Read("rich"); tagged.CustomTags["fixture"] = "public"; await Save(tagged, "custom-tags-input");
await Save(OfdDocumentMixer.Mix([new(await Read("custom-tags-input"), 0)]), "mix-custom-tags");
var conflictingTags = await Read("rich"); conflictingTags.CustomTags["fixture"] = "other"; await Save(conflictingTags, "custom-tags-conflict");
try { OfdDocumentMixer.Mix([new(await Read("custom-tags-input"), 0), new(await Read("custom-tags-conflict"), 0)]); throw new Exception("Mix must reject conflicting CustomTags."); }
catch (NotSupportedException exception) { File.WriteAllText(PathFor("custom-tags-input.rejection.txt"), exception.Message); }
await Save(OfdDocumentSplitter.Split(signed, [1, 0]), "split");
File.Copy(PathFor("signed.ofd"), PathFor("template-liveness-input.ofd"), true);
Mutate("template-liveness-input", entries => entries["Doc_0/Extensions/state.dat"] = Encoding.UTF8.GetBytes("<Wrapper><Extension BaseLoc='../Templates'><File>Content.xml</File></Extension></Wrapper>"));
await Save(OfdDocumentSplitter.Split(await Read("template-liveness-input"), [1]), "split-template-liveness");
File.Copy(PathFor("rich.ofd"), PathFor("template-wrapper-extension.ofd"), true);
Mutate("template-wrapper-extension", entries =>
{
    var page = Xml(entries[source.Pages[0].SourceEntryPath!]);
    page.Root!.Element(ns + "Template")!.SetAttributeValue(XNamespace.Get("urn:vendor") + "Placement", "unsupported");
    entries[source.Pages[0].SourceEntryPath!] = Bytes(page);
});
try { OfdDocumentMixer.Mix([new(await Read("template-wrapper-extension"), 0)]); throw new Exception("Mix must reject unmodeled template wrapper metadata."); }
catch (NotSupportedException exception) { File.WriteAllText(PathFor("template-wrapper-extension.rejection.txt"), exception.Message); }
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
File.Copy(PathFor("rich.ofd"), PathFor("annotation-shared-budget-input.ofd"), true);
Mutate("annotation-shared-budget-input", entries =>
{
    var index = Xml(entries["Doc_0/Annots/Annotations.xml"]); var record = index.Root!.Elements().Single();
    var duplicate = new XElement(record); duplicate.SetAttributeValue("PageID", source.Pages[1].Id); index.Root.Add(duplicate); entries["Doc_0/Annots/Annotations.xml"] = Bytes(index);
});
foreach (var objectBudget in new[] { true, false })
{
    using var zip = ZipFile.OpenRead(PathFor("annotation-shared-budget-input.ofd")); var bytes = zip.GetEntry("Doc_0/Annots/Page.xml")!.Length;
    var options = objectBudget ? new OfdPackageLoadOptions { MaxAnnotationObjectCount = source.Pages[0].AnnotationAppearances.Count } : new OfdPackageLoadOptions { MaxAnnotationXmlBytes = bytes * 2 - 1 };
    using var input = File.OpenRead(PathFor("annotation-shared-budget-input.ofd"));
    try { await new OfdReader().ReadAsync(input, options); throw new Exception("Shared annotation cumulative budget must refuse."); }
    catch (InvalidDataException exception) { File.WriteAllText(PathFor("annotation-shared-" + (objectBudget ? "objects" : "xml") + ".rejection.txt"), exception.Message); }
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
// A portable reproduction of Issue 02's fixed-baseline unmarked CTM.
// These rectangle glyphs are original MIT fixtures, not third-party reading fonts.
var fixedAnchor = new OfdDocumentPackage();
var rectangleFont = File.ReadAllBytes(Path.Combine(root, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
fixedAnchor.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Ofdrw Test Face", Data = rectangleFont });
fixedAnchor.Fonts.Add(new OfdFontResource { Id = "11", FontName = "OFD Example Noto", Data = fontBytes });
var fixedPage = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 70 };
foreach (var (label, x) in new[] { ("IDENTITY", 20d), ("UNMARKED CTM", 60d) })
    fixedPage.Elements.Add(new OfdTextElement { Text = label, FontName = "OFD Example Noto", FontResourceId = "11", XMillimeters = x, YMillimeters = 5, FontSizeMillimeters = 3 });
foreach (var italic in new[] { false, true })
foreach (var shear in new[] { false, true })
{
    var glyph = new OfdTextElement { Text = "I", FontName = "Ofdrw Test Face", FontResourceId = "10", Italic = italic,
        FontSizeMillimeters = 8, XMillimeters = shear ? 60 : 20, YMillimeters = italic ? 40 : 20,
        FillColor = new OfdColor(128, 0, 128), Transform = shear ? [1, 0, -0.2, 1, 1.6, 0] : [1, 0, 0, 1, 0, 0] };
    glyph.Runs.Add(new OfdTextRun { Text = "I", YMillimeters = 8 }); fixedPage.Elements.Add(glyph);
}
fixedAnchor.Pages.Add(fixedPage); await Save(fixedAnchor, "italic-fixed-anchor");
File.Copy(Path.Combine(root, "scripts/generate-font-test-fixtures.py"), PathFor("fonts/generate-font-test-fixtures.py"), true);
File.Copy(Path.Combine(root, "LICENSE"), PathFor("fonts/MIT-rectangle-LICENSE.txt"), true);
File.WriteAllBytes(PathFor("fonts/narrow.ttf"), rectangleFont);
foreach (var name in new[] { "baseline-native", "baseline-default", "rich", "annotation-metadata", "binary-xml-split", "page-wrapper-metadata", "box-whitespace", "box-inexact-page", "box-inexact-inherited", "annotation-inexact-appearance", "annotation-inexact-primitive", "blank-watermark", "signed", "watermark", "watermark-merged", "watermark-resource-suffix", "vendor-annotations-roundtrip", "mix-custom-tags", "split", "split-template-liveness", "mix", "clean", "overlay", "annotation-clipped", "annotation-clipped-mix", "italic-marked", "italic-control", "italic-user-matrix", "italic-fixed-anchor" })
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
