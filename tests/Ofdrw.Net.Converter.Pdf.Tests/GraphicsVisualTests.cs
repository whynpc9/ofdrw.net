using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class GraphicsVisualTests
{
    [Fact]
    public async Task AnisotropicCtm_TransformsStrokeOutlineAlongBothAxes()
    {
        var (package, _, graphics) = Create();
        var pen = new OfdPen(OfdColor.Black, 1);
        graphics.Scale(2, 1);
        graphics.DrawLine(pen, 10, 10, 10, 30); // physical x=20; horizontal stroke width=2 mm
        graphics.DrawLine(pen, 25, 50, 35, 50); // physical y=50; vertical stroke width=1 mm
        var raster = await Render(package);
        var pixelMillimeters = 25.4 / 72 / 2;
        var verticalWidth = Enumerable.Range(-20, 41).Count(offset => Gray(raster, 20 + offset * pixelMillimeters, 20) < 120);
        var horizontalWidth = Enumerable.Range(-20, 41).Count(offset => Gray(raster, 60, 50 + offset * pixelMillimeters) < 120);
        Assert.InRange(verticalWidth, 9, 14);
        Assert.InRange(horizontalWidth, 4, 8);
        Assert.True(verticalWidth >= horizontalWidth + 4, $"X-stroke={verticalWidth}px Y-stroke={horizontalWidth}px");
    }

    [Theory]
    [InlineData(OfdFillRule.EvenOdd, true)]
    [InlineData(OfdFillRule.NonZero, false)]
    public async Task NestedSameDirectionContours_RespectFillRule(OfdFillRule rule, bool hasHole)
    {
        var (package, _, graphics) = Create();
        var path = new OfdGraphicsPath(rule).AddRectangle(10, 10, 50, 50).AddRectangle(20, 20, 30, 30);
        graphics.FillPath(new OfdBrush(OfdColor.Black), path);
        var raster = await Render(package);
        Assert.InRange(Gray(raster, 15, 15), 0, 20);
        Assert.InRange(Gray(raster, 35, 35), hasHole ? 240 : 0, hasHole ? 255 : 20);
        Assert.InRange(Gray(raster, 70, 35), 240, 255);
    }

    [Fact]
    public async Task ClipSnapshot_RemainsAtOriginalPagePositionAcrossSaveAndRestore()
    {
        var (package, _, graphics) = Create();
        var clip = new OfdGraphicsPath().AddRectangle(10, 10, 20, 20);
        graphics.IntersectClip(clip);
        graphics.Save();
        graphics.Translate(30, 0);
        clip.AddRectangle(60, 60, 10, 10); // mutation after clip must not change output
        graphics.FillRectangle(new OfdBrush(OfdColor.Black), -20, 10, 40, 20); // physical 10..50, clipped at 30
        graphics.ResetClip();
        graphics.FillRectangle(new OfdBrush(new OfdColor(255, 0, 0)), 5, 10, 10, 20); // physical 35..45
        graphics.Restore();
        graphics.FillRectangle(new OfdBrush(OfdColor.Black), 50, 10, 20, 20); // clipped out again
        var raster = await Render(package);
        Assert.InRange(Gray(raster, 15, 20), 0, 20);
        Assert.InRange(Gray(raster, 33, 20), 240, 255);
        Assert.InRange(Red(raster, 40, 20), 240, 255);
        Assert.InRange(Gray(raster, 60, 20), 240, 255);
    }

    [Fact]
    public async Task FractionalScale_RightStrokeStaysInsideFrozenClipOnNarrowPage()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 1001, HeightMillimeters = 20 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        graphics.Scale(0.9996, 1);
        graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(0, 0, 1000, 10));
        graphics.DrawRectangle(new OfdPen(OfdColor.Black, 0.5), 0, 0, 1000, 10);
        var raster = await Render(package);
        Assert.InRange(Gray(raster, 999.45, 5), 0, 100);
        Assert.InRange(Gray(raster, 1000.3, 5), 240, 255);
    }

    private static (OfdDocumentPackage Package, OfdPage Page, OfdGraphics Graphics) Create()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        return (package, page, new OfdGraphics(package, page));
    }

    private static async Task<(byte[] Pixels, int Width)> Render(OfdDocumentPackage package)
    {
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var document = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(2d));
        using var page = document.GetPageReader(0);
        return (page.GetImage(), page.GetPageWidth());
    }

    private static int Red((byte[] Pixels, int Width) raster, double x, double y) => Pixel(raster, x, y, 2);
    private static int Gray((byte[] Pixels, int Width) raster, double x, double y) => Pixel(raster, x, y, 1);
    private static int Pixel((byte[] Pixels, int Width) raster, double x, double y, int channel)
    {
        var offset = ((int)Math.Round(y * 72 / 25.4 * 2) * raster.Width + (int)Math.Round(x * 72 / 25.4 * 2)) * 4;
        var alpha = raster.Pixels[offset + 3];
        return (int)Math.Round(raster.Pixels[offset + channel] * alpha / 255d + 255 - alpha);
    }
}
