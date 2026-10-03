using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;

if (args.Length != 2) throw new ArgumentException("Usage: sample output.ofd licensed-Noto-Regular.ttf");
var package = new OfdDocumentPackage();
// Register one caller-owned, licensed resource. OfdFont only binds this identity.
package.Fonts.Add(new OfdFontResource { Id = "ticket-font", FontName = "Noto Sans CJK SC", FileName = "Noto-Regular.ttf", Data = File.ReadAllBytes(args[1]) });
var black = new OfdBrush(new OfdColor(32, 44, 58));
var blue = new OfdBrush(new OfdColor(30, 93, 166));
var green = new OfdBrush(new OfdColor(26, 133, 112));
var red = new OfdBrush(new OfdColor(186, 58, 65));
var grid = new OfdPen(new OfdColor(108, 131, 153), 0.25);
var strong = new OfdPen(new OfdColor(30, 93, 166), 0.75);
OfdFont Font(double size = 3.5, int weight = 400, bool italic = false) => new("Noto Sans CJK SC", size, weight, italic, "ticket-font");
OfdGraphics Page()
{
    var page = new OfdPage { Index = package.Pages.Count, WidthMillimeters = 148, HeightMillimeters = 210 };
    package.Pages.Add(page); return new OfdGraphics(package, page);
}
var invoice = Page();
invoice.FillRectangle(new OfdBrush(new OfdColor(232, 241, 249)), 10, 10, 128, 27);
invoice.DrawString("原生绘图票面 / NATIVE TICKET", Font(5.2, 700), blue, 15, 22);
invoice.DrawString("OFD paths + searchable Unicode text", Font(3), black, 15, 30);
invoice.DrawString("No. 2026-004", Font(), black, 12, 48);
invoice.DrawString("日期 Date: 2026-10-02", Font(), black, 73, 48);
invoice.FillRectangle(new OfdBrush(new OfdColor(225, 236, 246)), 12, 58, 124, 13);
invoice.DrawRectangle(grid, 12, 58, 124, 70);
foreach (var y in new[] { 71d, 90, 109 }) invoice.DrawLine(grid, 12, y, 136, y);
foreach (var x in new[] { 82d, 108 }) invoice.DrawLine(grid, x, 58, x, 128);
invoice.DrawString("项目 Item", Font(3.5, 700), black, 16, 67);
invoice.DrawString("数量 Qty", Font(3.5, 700), black, 85, 67);
invoice.DrawString("金额 CNY", Font(3.5, 700), black, 111, 67);
invoice.DrawString("技术文档 Technical notes", Font(3.4), black, 16, 83);
invoice.DrawString("2", Font(), black, 92, 83);
invoice.DrawString("128.00", Font(), black, 113, 83);
invoice.DrawString("矢量示意 Vector diagram", Font(3.4, italic: true), blue, 16, 102);
invoice.DrawString("1", Font(), black, 92, 102);
invoice.DrawString("64.00", Font(), black, 113, 102);
invoice.DrawString("合计 Total", Font(3.5, 700), black, 16, 121);
invoice.DrawString("192.00", Font(3.5, 700), red, 111, 121);
invoice.DrawLine(new OfdPen(new OfdColor(181, 196, 208), 0.15), 12, 148, 136, 148);
invoice.DrawString("基线 Baseline: Wi fi / 中文", Font(4), black, 12, 148);
invoice.DrawString("局部红色", Font(), red, 12, 160);
invoice.DrawString("Regular remains regular", Font(), black, 38, 160);
invoice.Save(); invoice.Translate(108, 173); invoice.Rotate(-12);
invoice.DrawRectangle(strong, -23, -8, 46, 16);
invoice.DrawString("已付款 PAID", Font(4, 700), blue, -20, 1);
invoice.Restore();
invoice.DrawString("Native PathObject / TextObject - page 1 / 2", Font(2.8), black, 12, 199);

var diagram = Page();
diagram.DrawString("变换与裁剪 / TRANSFORMS", Font(5.2, 700), blue, 12, 22);
diagram.DrawString("Full affine strokes, curves and fill rules", Font(3), black, 12, 31);
diagram.DrawString("Scale (2, 1): both pens are 1 mm in user space", Font(2.8), black, 12, 43);
diagram.Save(); diagram.Translate(15, 50); diagram.Scale(2, 1);
diagram.DrawLine(new OfdPen(new OfdColor(30, 93, 166), 1), 0, 0, 35, 0);
diagram.DrawLine(new OfdPen(new OfdColor(30, 93, 166), 1), 42, 0, 42, 16);
diagram.Restore();
diagram.Save(); diagram.Translate(22, 86); diagram.Rotate(-15);
diagram.DrawPath(new OfdGraphicsPath().MoveTo(0, 0).BezierTo(8, -15, 28, 15, 38, 0).QuadraticTo(46, -10, 55, 0), strong);
diagram.DrawString("曲线 Bezier", Font(3.4, italic: true), green, 0, 9);
diagram.Restore();
// Acute miter ratio ~5.75: SVG's default 4 would bevel this tip.
diagram.DrawPath(new OfdGraphicsPath().MoveTo(108, 97).LineTo(111, 80).LineTo(114, 97), new OfdPen(new OfdColor(30, 93, 166), 1.4));
diagram.DrawString("Miter=10", Font(2.5), black, 101, 105);
diagram.DrawString("EvenOdd", Font(3), black, 17, 114);
diagram.DrawString("NonZero", Font(3), black, 78, 114);
diagram.FillPath(green, new OfdGraphicsPath(OfdFillRule.EvenOdd).AddRectangle(16, 120, 40, 26).AddRectangle(26, 127, 20, 12));
diagram.FillPath(green, new OfdGraphicsPath().AddRectangle(77, 120, 40, 26).AddRectangle(87, 127, 20, 12));
diagram.DrawString("Frozen clip + translated/sheared content", Font(2.8), black, 12, 158);
diagram.Save(); diagram.IntersectClip(new OfdGraphicsPath().AddRectangle(18, 164, 100, 19));
diagram.Translate(10, 0); diagram.MultiplyTransform(new OfdMatrix(1, 0, 0.2, 1, 0, 0));
diagram.FillRectangle(new OfdBrush(new OfdColor(234, 241, 247)), -40, 160, 190, 30);
diagram.DrawString("裁剪固定 Clip stays in page space", Font(5, 700), blue, -25, 178);
diagram.Restore(); diagram.DrawRectangle(grid, 18, 164, 100, 19);
diagram.DrawString("Restored: untransformed, unclipped - page 2 / 2", Font(2.8), black, 12, 199);

await using var output = File.Create(args[0]);
await new OfdPackageWriter().WriteAsync(package, output);
Console.WriteLine($"Wrote {package.Pages.Count} pages; {package.Pages.Sum(page => page.Elements.Count)} native path/text objects, one font, no images.");
