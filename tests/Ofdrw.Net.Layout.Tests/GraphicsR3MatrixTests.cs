using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class GraphicsR3MatrixTests
{
    [Theory]
    [InlineData(0.186, 2.074, -0.93, -10.37)]
    [InlineData(-0.22, 1.15, -3.542, 18.515)]
    public void DecimalSingularMatrix_IsRejectedWithoutChangingCurrentOrSavedState(double a, double b, double c, double d)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        graphics.Rotate(90); // must remain allowed despite floating point cosine residue
        var rotated = graphics.Transform;
        var singular = new OfdMatrix(a, b, c, d, 0, 0);
        Assert.Throws<ArgumentException>(() => graphics.MultiplyTransform(singular));
        Assert.Same(rotated, graphics.Transform);
        graphics.SetTransform(OfdMatrix.Translation(4, 5));
        var original = graphics.Transform;
        graphics.Save();
        Assert.Throws<ArgumentException>(() => graphics.SetTransform(singular));
        Assert.Same(original, graphics.Transform);
        Assert.Throws<ArgumentException>(() => graphics.MultiplyTransform(singular));
        Assert.Same(original, graphics.Transform);
        graphics.Restore();
        Assert.Same(original, graphics.Transform);
        Assert.Empty(page.Elements);
    }

    [Fact]
    public async Task NativePathAndFrozenClip_KeepSameFractionalScaleAfterRoundTrip()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 1001, HeightMillimeters = 20 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        graphics.Scale(0.9996, 1);
        graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(0, 0, 1000, 10));
        graphics.DrawRectangle(new OfdPen(OfdColor.Black, 0.5), 0, 0, 1000, 10);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var saved = await new OfdReader().ReadAsync(stream);
        var path = Assert.IsType<OfdPathElement>(Assert.Single(saved.Pages[0].Elements));
        Assert.Equal(0.9996, path.Transform![0], 10);
        Assert.Contains("0.9996", path.ClippingXml);
    }
}
