using Docnet.Core;
using Docnet.Core.Models;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class PublicTableVisualTests
{
    [Fact]
    public async Task TableGrid_ShouldNotPaintDiagonalBridgesAcrossCellInteriors()
    {
        var document = new FlowDocument(); document.Options.PageWidthMillimeters = 80; document.Options.PageHeightMillimeters = 60;
        document.Options.MarginLeftMillimeters = document.Options.MarginRightMillimeters = 10;
        document.Options.MarginTopMillimeters = document.Options.MarginBottomMillimeters = 10;
        var table = new Table { BorderWidthMillimeters = 0.3 }; var row = new Row { MinimumHeightMillimeters = 30 };
        row.Cells.Add(new Cell { BackgroundColor = new OfdColor(220, 230, 250) });
        row.Cells.Add(new Cell { BackgroundColor = new OfdColor(255, 246, 218) }); table.Rows.Add(row); document.Blocks.Add(table);
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(document.Render(), ofd); ofd.Position = 0;
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var reader = DocLib.Instance.GetDocReader(pdf.ToArray(), new PageDimensions(3d)); using var page = reader.GetPageReader(0);
        var pixels = page.GetImage(); var width = page.GetPageWidth(); var height = page.GetPageHeight();
        var darkInteriorPixels = 0;
        for (var y = (int)(12 * height / 60d); y < (int)(38 * height / 60d); y++)
        for (var x = (int)(12 * width / 80d); x < (int)(38 * width / 80d); x++)
        {
            var i = 4 * (y * width + x);
            if (pixels[i] < 100 && pixels[i + 1] < 100 && pixels[i + 2] < 100) darkInteriorPixels++;
        }
        Assert.Equal(0, darkInteriorPixels);
        int[] Pixel(double x, double y)
        {
            var i = 4 * ((int)(y * height / 60d) * width + (int)(x * width / 80d));
            return new[] { (int)pixels[i + 2], pixels[i + 1], pixels[i], pixels[i + 3] }; // BGRA → RGBA.
        }
        Assert.Equal(new[] { 220, 230, 250, 255 }, Pixel(20, 25));
        Assert.Equal(new[] { 255, 246, 218, 255 }, Pixel(50, 25));
        // Search the stroke neighborhood; rasterization can place the exact geometric coordinate on an antialiased edge.
        var divider = Enumerable.Range(-2, 5).Select(offset => Pixel(40 + offset * 80d / width, 25))
            .OrderBy(value => value.Take(3).Sum()).First();
        Assert.All(divider.Take(3), channel => Assert.InRange(channel, 0, 40));
        Assert.Equal(255, divider[3]);
    }
}
