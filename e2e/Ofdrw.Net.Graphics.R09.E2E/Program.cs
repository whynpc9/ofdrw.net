using System.IO.Compression;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

if (args.Length != 1) throw new ArgumentException("Usage: r09-e2e output-directory");
var output = Path.GetFullPath(args[0]); Directory.CreateDirectory(output);
var quadrants = Convert.FromBase64String(
    "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAAIUlEQVR4nGP4z8DwHx8mAo4aMAwMIKAfCBjw4lEDhoUBAFfoPd+1WG54AAAAAElFTkSuQmCC");
var package = new OfdDocumentPackage();
OfdGraphics Page()
{
    var page = new OfdPage { Index = package.Pages.Count, WidthMillimeters = 148, HeightMillimeters = 210 };
    package.Pages.Add(page); return new OfdGraphics(package, page);
}

var images = Page();
images.DrawImage(quadrants, 10, 10, 42, 42);
images.DrawImage(quadrants, 65, 10, 32, 32, 8, 0, 8, 8);
images.DrawImage(quadrants, 105, 10, 32, 32, 0, 8, 8, 8, 128);
images.Save();
images.IntersectClip(new OfdGraphicsPath().AddRectangle(15, 66, 100, 49));
images.Translate(61, 87); images.Rotate(24); images.Scale(-1.2, 0.85);
images.DrawImage(quadrants, -35, -20, 50, 40, 0, 0, 8, 8, 220);
images.Restore();
images.DrawRectangle(new OfdPen(new OfdColor(45, 52, 62), 0.4), 15, 66, 100, 49);
images.DrawImage(quadrants, 10, 116, 40, 10, 7.75, 0, 1, 8);
images.FillImagePatternRectangle(quadrants, 10, 130, 55, 35, 11, 11);
images.FillLinearGradientRectangle(new OfdColor(220, 26, 35), new OfdColor(30, 72, 229), 74, 130, 62, 35);
images.FillLinearGradientRectangle(new OfdColor(20, 147, 85, 220), new OfdColor(255, 225, 50, 80), 10, 176, 126, 24,
    horizontal: false);

var pens = Page();
var gray = new OfdPen(new OfdColor(180, 188, 198), 0.25);
foreach (var y in new[] { 30d, 65, 100, 145, 180 }) pens.DrawLine(gray, 10, y, 138, y);
pens.DrawLine(new OfdPen(new OfdColor(220, 26, 35), 4, OfdLineCap.Butt), 20, 23, 90, 23);
pens.DrawLine(new OfdPen(new OfdColor(30, 140, 76), 4, OfdLineCap.Round), 20, 58, 90, 58);
pens.DrawLine(new OfdPen(new OfdColor(30, 72, 229), 4, OfdLineCap.Square), 20, 93, 90, 93);
pens.DrawPath(new OfdGraphicsPath().MoveTo(15, 138).LineTo(38, 111).LineTo(61, 138),
    new OfdPen(new OfdColor(220, 26, 35), 4, OfdLineCap.Butt, join: OfdLineJoin.Miter, miterLimit: 10));
pens.DrawPath(new OfdGraphicsPath().MoveTo(68, 138).LineTo(91, 111).LineTo(114, 138),
    new OfdPen(new OfdColor(30, 140, 76), 4, OfdLineCap.Butt, join: OfdLineJoin.Round));
pens.DrawLine(new OfdPen(new OfdColor(30, 72, 229), 2, new[] { 8d, 4d }, dashOffset: 2), 16, 171, 132, 171);
pens.DrawLine(new OfdPen(new OfdColor(190, 60, 153), 2, new[] { 6d, 3d, 2d }, dashOffset: 1), 16, 195, 132, 195);

var ofdPath = Path.Combine(output, "r09-graphics.ofd");
await using (var destination = File.Create(ofdPath)) await new OfdPackageWriter().WriteAsync(package, destination);
await using (var source = File.OpenRead(ofdPath))
{
    var read = await new OfdReader().ReadAsync(source);
    if (read.Pages.Count != 2 || read.Pages[0].Elements.OfType<OfdImageElement>().Count() < 4)
        throw new InvalidDataException("R09 graphics OFD round-trip lost images or pages.");
}
using (var archive = ZipFile.OpenRead(ofdPath))
    if (archive.Entries.Count(entry => entry.Name.StartsWith("Image_", StringComparison.Ordinal)) != 3)
        throw new InvalidDataException("Expected one original tile and two gradient resources.");
await using (var source = File.OpenRead(ofdPath))
await using (var destination = File.Create(Path.Combine(output, "r09-graphics.pdf")))
    await new OfdToPdfConverter().ConvertAsync(source, destination);
for (var index = 0; index < 2; index++)
{
    await using (var pngSource = File.OpenRead(ofdPath))
    await using (var pngDestination = File.Create(Path.Combine(output, $"r09-graphics-page-{index + 1}.png")))
        await new OfdToImageConverter().ConvertAsync(pngSource, pngDestination, index);
    await using var source = File.OpenRead(ofdPath);
    await using var destination = File.Create(Path.Combine(output, $"r09-graphics-page-{index + 1}.svg"));
    await new OfdToSvgConverter().ConvertAsync(source, destination, index);
}
await using (var source = File.OpenRead(ofdPath))
await using (var destination = File.Create(Path.Combine(output, "r09-graphics.html")))
    await new OfdToHtmlConverter().ConvertAsync(source, destination);
Console.WriteLine($"R09 generated {package.Pages.Count} OFD pages with PDF, independent SVG and HTML exports at {output}.");
