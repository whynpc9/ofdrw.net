using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class GraphicsBudgetAndStrokeTests
{
    [Fact]
    public void TinyAdvances_StopAtGeometryBudgetWithoutEnumeratingEntireInputOrConsumingDrawBudget()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page,
            new OfdGraphicsOptions { MaxGeometryCharacters = 1000, MaxTextCharacters = 10001, MaxPageElements = 1 });
        var advances = new CountedAdvances(10000);
        Assert.Throws<InvalidOperationException>(() => graphics.DrawString(new string('A', 10001),
            new OfdFont("Plain", 4), new OfdBrush(OfdColor.Black), 0, 10, advances));
        Assert.InRange(advances.Reads, 3, 4); // four 326-character subnormal literals cannot fit
        Assert.Empty(page.Elements);
        graphics.DrawString("OK", new OfdFont("Plain", 4), new OfdBrush(OfdColor.Black), 0, 10);
        Assert.Single(page.Elements);
    }

    [Fact]
    public async Task WriterRoundedThinPen_KeepsConservativeAcuteMiterBoundary()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 12000, HeightMillimeters = 20000 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page);
        graphics.Scale(1000, 1000);
        var acute = new OfdGraphicsPath().MoveTo(9, 19).LineTo(10, 10).LineTo(11, 19);
        graphics.DrawPath(acute, new OfdPen(OfdColor.Black, 0.0006));
        var native = Assert.IsType<OfdPathElement>(Assert.Single(page.Elements));
        Assert.True(native.XMillimeters <= 9000 - 5, $"Boundary left {native.XMillimeters}");
        Assert.True(native.XMillimeters + native.WidthMillimeters >= 11000 + 5);
        Assert.True(native.YMillimeters <= 10000 - 5);
        Assert.True(native.YMillimeters + native.HeightMillimeters >= 19000 + 5);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var saved = await new OfdReader().ReadAsync(stream);
        var path = Assert.IsType<OfdPathElement>(Assert.Single(saved.Pages[0].Elements));
        Assert.Equal(0.001, path.LineWidthMillimeters, 3);
        Assert.True(path.XMillimeters <= 9000 - 5);
        Assert.True(path.XMillimeters + path.WidthMillimeters >= 11000 + 5);
    }

    [Fact]
    public void PathDataBudget_RejectsLongDecimalWithoutAppendingOrConsumingFailedDraw()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page,
            new OfdGraphicsOptions { MaxGeometryCharacters = 100, MaxPageElements = 1 });
        var tiny = new OfdGraphicsPath().MoveTo(double.Epsilon, 0).LineTo(1, 1);
        Assert.Throws<InvalidOperationException>(() => graphics.DrawPath(tiny, new OfdPen(OfdColor.Black)));
        Assert.Empty(page.Elements);
        graphics.DrawLine(new OfdPen(OfdColor.Black), 0, 0, 1, 1);
        Assert.Single(page.Elements);
    }

    [Fact]
    public void ClipXmlMarkupBudget_RejectsAtomicallyAndResetAllowsValidDraw()
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 };
        package.Pages.Add(page);
        var graphics = new OfdGraphics(package, page,
            new OfdGraphicsOptions { MaxGeometryCharacters = 100, MaxPageElements = 1 });
        graphics.IntersectClip(new OfdGraphicsPath().AddRectangle(0, 0, 10, 10));
        Assert.Throws<InvalidOperationException>(() => graphics.DrawLine(new OfdPen(OfdColor.Black), 1, 1, 2, 2));
        Assert.Empty(page.Elements);
        graphics.ResetClip();
        graphics.DrawLine(new OfdPen(OfdColor.Black), 1, 1, 2, 2);
        Assert.Single(page.Elements);
    }

    private sealed class CountedAdvances : IReadOnlyList<double>
    {
        public CountedAdvances(int count) => Count = count;
        public int Count { get; }
        public int Reads { get; private set; }
        public double this[int index] { get { Reads++; return double.Epsilon; } }
        public IEnumerator<double> GetEnumerator()
        {
            for (var index = 0; index < Count; index++) yield return this[index];
        }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
