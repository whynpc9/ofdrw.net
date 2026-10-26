using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;

if (args.Length < 3) throw new ArgumentException("Usage: output-directory font-directory checkerboard.png [original-v1.ofd]");
var output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
var package = new OfdDocumentPackage { CompositeResources = new() };
foreach (var item in new[] { ("Noto Sans", "NotoSans-Regular.ttf"), ("Noto Sans CJK SC", "Ofdrw-CI-NotoSansCJKsc-Regular.ttf") })
    package.Fonts.Add(new OfdFontResource { FontName = item.Item1, FileName = item.Item2, Data = File.ReadAllBytes(Path.Combine(args[1], item.Item2)) });
var imageData = File.ReadAllBytes(args[2]);
var leaf = new OfdCompositeResource { WidthMillimeters = 70, HeightMillimeters = 45 };
leaf.Elements.Add(Shape(0, 0, 70, 45, "M 0 0 L 70 0 L 70 45 L 0 45 C", new(230, 241, 250)));
leaf.Elements.Add(Text("共享组合 / shared vector", 3, 2, 64, "Noto Sans CJK SC", new(25, 80, 130), true));
leaf.Elements.Add(Text("Proportional Wiim: ABC xyz", 3, 13, 64, "Noto Sans", new(40, 40, 40), italic: true));
leaf.Elements.Add(new OfdImageElement { Data = imageData, FileName = "checkerboard.png", XMillimeters = 3, YMillimeters = 27, WidthMillimeters = 25, HeightMillimeters = 14 });
leaf.Elements.Add(Shape(40, 27, 25, 14, "M 0 0 L 25 0 L 25 14 L 0 14 C M 0 7 L 25 7 M 12 0 L 12 14", new(240, 224, 170)));
leaf.Elements.Add(Shape(68, 20, 12, 12, "M 0 0 L 12 0 L 12 12 L 0 12 C", new(220, 70, 60)));
package.CompositeResources.Items.Add(leaf);
var parent = new OfdCompositeResource { WidthMillimeters = 80, HeightMillimeters = 55 };
parent.Elements.Add(Place(leaf, 4, 5, 70, 45)); package.CompositeResources.Items.Add(parent);
var page1 = new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 };
page1.Elements.Add(Text("Typed clips and composites / 强类型图形", 15, 12, 185, "Noto Sans CJK SC", new(20, 70, 110), true));
page1.Elements.Add(Text("One definition, two placements; red overflow is clipped", 15, 26, 185, "Noto Sans", new(60, 60, 60)));
page1.Elements.Add(Place(parent, 15, 45, 80, 55));
page1.Elements.Add(Place(parent, 110, 50, 80, 70, new double[] { .95, .16, -.12, 1, 5, 0 }));
page1.Elements.Add(Text("Only the left placement moves (+8,+6) after save", 15, 132, 180, "Noto Sans", new(80, 40, 150)));
page1.Elements.Add(Text("Original multi-M compound table reference", 15, 153, 180, "Noto Sans", new(40, 40, 40)));
page1.Elements.Add(Shape(20, 171, 170, 30, "M 0 0 L 170 0 L 170 30 L 0 30 C M 0 10 L 170 10 M 85 0 L 85 30", new(220, 235, 245)));
page1.Elements.Add(Text("项目 / Item", 23, 172, 75, "Noto Sans CJK SC", new(0, 0, 0), true));
page1.Elements.Add(Text("数量 / Count", 108, 172, 75, "Noto Sans CJK SC", new(0, 0, 0), true));
page1.Elements.Add(Text("Alpha", 23, 184, 75, "Noto Sans", new(0, 0, 0))); page1.Elements.Add(Text("128", 108, 184, 75, "Noto Sans", new(0, 0, 0)));
page1.Elements.Add(Text("Page 1 / 第一页", 80, 278, 80, "Noto Sans CJK SC", new(80, 80, 80))); package.Pages.Add(page1);
var page2 = new OfdPage { Index = 1, WidthMillimeters = 210, HeightMillimeters = 297, Annotations = new() };
page2.Elements.Add(Text("Clips: one Path Area per Clip; successive intersection", 15, 12, 185, "Noto Sans", new(20, 70, 110), true));
var image = new OfdImageElement { Data = imageData, FileName = "checkerboard.png", XMillimeters = 20, YMillimeters = 40, WidthMillimeters = 75, HeightMillimeters = 50,
    Transform = new double[] { 75, 8, 0, 50, 0, 0 }, Clips = Clips("M 0 0 L 65 0 L 65 42 L 0 42 C") };
image.Clips.Items[0].Areas[0].Transform = new double[] { 1, .08, .05, 1, 3, 2 };
image.Clips.Items.Add(Clips("M 0 0 L 55 0 L 55 45 L 0 45 C").Items[0]); page2.Elements.Add(image);
page2.Elements.Add(Text("Image: explicit CTM + Area CTM + two Clip intersection", 15, 96, 185, "Noto Sans", new(50, 50, 50)));
var clippedText = Text("Clipped text 中文 ABC xyz", 105, 40, 80, "Noto Sans CJK SC", new(130, 40, 80), true);
clippedText.Clips = Clips("M 0 0 L 38 0 L 38 8 L 0 8 C"); page2.Elements.Add(clippedText);
page2.Elements.Add(Text("Intentional text crop at 38 mm", 105, 58, 85, "Noto Sans", new(50, 50, 50)));
for (var i = 0; i < 2; i++)
{
    var path = Shape(25 + i * 90, 125, 65, 45, "M 0 0 L 65 0 L 65 45 L 0 45 C M 15 10 L 50 10 L 50 35 L 15 35 C", new(70, 155, 185));
    path.FillRule = i == 0 ? OfdFillRule.NonZero : OfdFillRule.EvenOdd;
    path.Clips = Clips(path.AbbreviatedData); path.Clips.Items[0].Areas[0].Path.FillRule = path.FillRule;
    page2.Elements.Add(path); page2.Elements.Add(Text(i == 0 ? "NonZero: same winding, solid" : "EvenOdd: visible hole", 20 + i * 90, 177, 85, "Noto Sans", new(40, 40, 40)));
}
var annotation = new OfdAnnotation { Type = OfdAnnotationType.Stamp, Creator = "OFL typed fixture", LastModifiedDate = new DateTime(2026, 10, 15), Appearance = new()
    { XMillimeters = 25, YMillimeters = 215, WidthMillimeters = 140, HeightMillimeters = 35, Transform = new double[] { 1, .08, 0, 1, 0, 0 } } };
annotation.Appearance.Elements.Add(Text("局部批注 / local annotation once", 3, 2, 125, "Noto Sans CJK SC", new(30, 110, 60), true));
var notePath = Shape(3, 14, 110, 12, "M 0 0 L 110 0 L 110 12 L 0 12 C", new(170, 220, 170)); notePath.Clips = Clips("M 0 0 L 75 0 L 75 12 L 0 12 C");
annotation.Appearance.Elements.Add(notePath); page2.Annotations.Items.Add(annotation);
page2.Elements.Add(Text("Page 2 / 第二页", 80, 278, 80, "Noto Sans CJK SC", new(80, 80, 80))); package.Pages.Add(page2);
// The same rich deterministic source exercises cross-document ownership before
// navigation is added below (navigation migration remains a separate feature).
await Save(package, "r05-source");
var r05Merged = OfdDocumentMerger.Merge(new[] { package, package });
Require(r05Merged.Pages.Count == 4 && r05Merged.CompositeResources!.Items.Count == 4,
    "R05 Merge lost per-occurrence composite closure");
Require(r05Merged.Pages.Sum(page => page.Annotations?.Items.Count ?? 0) == 2,
    "R05 Merge lost typed annotations");
Require(Count(new OfdTextExtractor().Extract(r05Merged), "共享组合 / shared vector") == 4 &&
    Count(new OfdTextExtractor().Extract(r05Merged), "局部批注 / local annotation once") == 2,
    "R05 Merge duplicated or lost visible text");
await Save(r05Merged, "r05-merged");
var r05Mixed = OfdDocumentMixer.Mix(new[] { new OfdMixSource(package, 0), new OfdMixSource(package, 1) });
Require(r05Mixed.Pages.Count == 1 && r05Mixed.CompositeResources!.Items.Count == 2 &&
    r05Mixed.Pages[0].Annotations is null, "R05 Mix lost paint-order flattening");
Require(Count(new OfdTextExtractor().Extract(r05Mixed), "共享组合 / shared vector") == 2 &&
    Count(new OfdTextExtractor().Extract(r05Mixed), "局部批注 / local annotation once") == 1,
    "R05 Mix duplicated or lost visible text");
await Save(r05Mixed, "r05-mixed");
package.Bookmarks = new(); var bookmark = new OfdBookmark { Name = "第二页", Destination = new OfdDestination { Page = page2, Type = OfdDestinationType.Fit } }; package.Bookmarks.Items.Add(bookmark);
package.Outlines = new(); var outline = new OfdOutline { Title = "图形样例 / graphics", Expanded = true, Actions = new() };
outline.Actions.Items.Add(new OfdGoToAction { Destination = new OfdDestination { Page = page1, Type = OfdDestinationType.Fit } });
var childOutline = new OfdOutline { Title = "裁剪 / clips", Actions = new() }; childOutline.Actions.Items.Add(new OfdGoToAction { Bookmark = bookmark }); outline.Children.Add(childOutline); package.Outlines.Items.Add(outline);
package.Actions = new(); package.Actions.Items.Add(new OfdGoToAction { Event = "DO", Destination = new OfdDestination { Page = page1, Type = OfdDestinationType.Fit } });
await Save(package, "before");
using var originalInput = File.OpenRead(Path.Combine(output, "before.ofd")); var original = await new OfdReader().ReadAsync(originalInput);
var edited = OfdDocumentSplitter.Split(original, new[] { 0, 1 });
var left = (OfdCompositeElement)edited.Pages[0].Elements[2]; left.XMillimeters += 8; left.YMillimeters += 6;
edited.Pages[1].Elements.OfType<OfdImageElement>().Single().Clips!.Items[0].Areas[0].Transform![4] += 5;
await Save(edited, "after");
using var readbackInput = File.OpenRead(Path.Combine(output, "after.ofd")); var readback = await new OfdReader().ReadAsync(readbackInput);
var placements = readback.Pages[0].Elements.OfType<OfdCompositeElement>().ToArray();
Require(readback.Outlines!.Items[0].Children.Count == 1 && ReferenceEquals(readback.Bookmarks!.Items[0].Destination.Page, readback.Pages[1]) && readback.Actions!.Items[0] is OfdGoToAction, "Navigation/actions lost on roundtrip");
Require(ReferenceEquals(placements[0].Resource, placements[1].Resource), "Shared resource copied per occurrence");
Require(placements[0].XMillimeters == 23 && placements[1].XMillimeters == 110, "Placement edit changed the wrong occurrence");
Require(original.Pages[0].Elements.OfType<OfdCompositeElement>().First().XMillimeters == 15, "Clone borrowed source mutable state");
Require(readback.Fonts.Select(f => Hash(f.Data)).Order().SequenceEqual(original.Fonts.Select(f => Hash(f.Data)).Order()), "Font payload changed");
Require(readback.Pages[1].Elements.OfType<OfdImageElement>().Single().Data.SequenceEqual(imageData), "Image payload changed");
var extracted = new OfdTextExtractor().Extract(readback); Require(Count(extracted, "共享组合 / shared vector") == 2, "Definition plus occurrence text duplication");
Require(Count(extracted, "局部批注 / local annotation once") == 1, "Annotation projection painted twice");
File.WriteAllText(Path.Combine(output, "extracted.txt"), extracted);
foreach (var stem in new[] { "before", "after" }) await Export(stem);
await Export("r05-source"); await Export("r05-merged", 4); await Export("r05-mixed", 1);
if (args.Length > 3) { File.Copy(args[3], Path.Combine(output, "original-v1.ofd"), true); await Export("original-v1"); }
var files = Directory.GetFiles(output).Where(p => Path.GetExtension(p) is ".ofd" or ".pdf" or ".svg").Select(p => new { Name = Path.GetFileName(p), Bytes = new FileInfo(p).Length, Sha256 = Hash(File.ReadAllBytes(p)) });
File.WriteAllText(Path.Combine(output, "results.json"), JsonSerializer.Serialize(new { Runtime = Environment.Version.ToString(), Source = typeof(OfdPage).Assembly.GetCustomAttributes(false).OfType<System.Reflection.AssemblyInformationalVersionAttribute>().Single().InformationalVersion,
    Passed = true, PreviewCompleted = false, ResourceDefinitions = readback.CompositeResources!.Items.Count, Placements = placements.Length,
    Fonts = readback.Fonts.Select(f => new { f.FontName, Bytes = f.Data.Length, Sha256 = Hash(f.Data) }), Files = files }, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine(File.ReadAllText(Path.Combine(output, "results.json")));
async Task Save(OfdDocumentPackage value, string stem) { using var stream = File.Create(Path.Combine(output, stem + ".ofd")); await new OfdPackageWriter().WriteAsync(value, stream); }
async Task Export(string stem, int pageCount = 2)
{
    using (var input = File.OpenRead(Path.Combine(output, stem + ".ofd"))) using (var pdf = File.Create(Path.Combine(output, stem + ".pdf"))) await new OfdToPdfConverter().ConvertAsync(input, pdf);
    for (var page = 0; page < pageCount; page++)
        using (var input = File.OpenRead(Path.Combine(output, stem + ".ofd"))) using (var svg = File.Create(Path.Combine(output, stem + $"-{page + 1}.svg")))
            await new OfdToSvgConverter().ConvertAsync(input, svg, page);
}
OfdCompositeElement Place(OfdCompositeResource resource, double x, double y, double w, double h, double[]? matrix = null) => new() { Resource = resource, XMillimeters = x, YMillimeters = y, WidthMillimeters = w, HeightMillimeters = h, Transform = matrix };
OfdTextElement Text(string value, double x, double y, double w, string font, OfdColor color, bool bold = false, bool italic = false) => new() { Text = value, XMillimeters = x, YMillimeters = y, WidthMillimeters = w, HeightMillimeters = 9, FontName = font, FontSizeMillimeters = 4, FillColor = color, Weight = bold ? 700 : 400, Italic = italic };
OfdPathElement Shape(double x, double y, double w, double h, string data, OfdColor fill) => new() { XMillimeters = x, YMillimeters = y, WidthMillimeters = w, HeightMillimeters = h, AbbreviatedData = data, Fill = true, FillColor = fill, StrokeColor = new(50, 75, 90), LineWidthMillimeters = .25 };
OfdClipCollection Clips(string data) { var result = new OfdClipCollection(); var clip = new OfdClip(); clip.Areas.Add(new OfdClipArea { Path = new OfdPathElement { AbbreviatedData = data, Fill = true, Stroke = false } }); result.Items.Add(clip); return result; }
string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
int Count(string text, string needle) => (text.Length - text.Replace(needle, "").Length) / needle.Length;
void Require(bool condition, string message) { if (!condition) throw new InvalidDataException(message); }
