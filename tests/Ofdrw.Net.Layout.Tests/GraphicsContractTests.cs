using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class GraphicsContractTests
{
    private static (OfdDocumentPackage Package, OfdPage Page, OfdGraphics Graphics) Context(OfdGraphicsOptions? options = null)
    {
        var package = new OfdDocumentPackage();
        var page = new OfdPage { XMillimeters = 7, YMillimeters = 11, WidthMillimeters = 100, HeightMillimeters = 80 };
        package.Pages.Add(page);
        return (package, page, new OfdGraphics(package, page, options));
    }

    [Fact]
    public async Task NativeObjects_RoundTripWithUniqueIdsAndText()
    {
        var (package, page, graphics) = Context();
        graphics.DrawLine(new OfdPen(OfdColor.Black, 0.5), 1, 2, 8, 9);
        graphics.FillRectangle(new OfdBrush(OfdColor.Black), 10, 12, 20, 15);
        graphics.DrawString("中文 ÁB", new OfdFont("Fixture", 4), new OfdBrush(OfdColor.Black), 2, 10);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var roundTrip = await new OfdReader().ReadAsync(stream);
        var elements = roundTrip.Pages.Single().Elements;
        Assert.Equal(3, elements.Count);
        Assert.Equal(2, elements.OfType<OfdPathElement>().Count());
        Assert.Equal("中文 ÁB", Assert.Single(elements.OfType<OfdTextElement>()).Text);
        Assert.All(elements, element => Assert.False(string.IsNullOrWhiteSpace(element.ObjectId)));
        Assert.Equal(3, elements.Select(element => element.ObjectId).Distinct().Count());
    }

    [Fact]
    public void MatrixOrder_AndPhysicalPageOrigin_ArePreserved()
    {
        var (_, page, graphics) = Context();
        graphics.Translate(12, 15);
        graphics.Scale(-2, 3);
        graphics.MultiplyTransform(new OfdMatrix(1, 0.5, 0.25, 1, 0, 0));
        graphics.DrawLine(new OfdPen(OfdColor.Black, 0.5), 1, 2, 4, 6);
        var path = Assert.IsType<OfdPathElement>(Assert.Single(page.Elements));
        var actual = path.Transform!;
        var expected = OfdMatrix.Translation(page.XMillimeters, page.YMillimeters)
            .Multiply(OfdMatrix.Translation(12, 15)).Multiply(OfdMatrix.Scale(-2, 3))
            .Multiply(new OfdMatrix(1, 0.5, 0.25, 1, 0, 0));
        Assert.Equal(new[] { expected.A, expected.B, expected.C, expected.D }, actual.Take(4));
        Assert.Equal(expected.E - path.XMillimeters, actual[4], 9);
        Assert.Equal(expected.F - path.YMillimeters, actual[5], 9);
        Assert.True(path.WidthMillimeters > 0);
        Assert.True(path.HeightMillimeters > 0);
    }

    [Fact]
    public void SaveRestore_AndClips_AreFrozenAtSetTime()
    {
        var (_, page, graphics) = Context();
        var clip = new OfdGraphicsPath().AddRectangle(1, 2, 10, 12);
        graphics.IntersectClip(clip);
        graphics.Save();
        graphics.Translate(20, 0);
        clip.AddRectangle(30, 30, 5, 5);
        graphics.FillRectangle(new OfdBrush(OfdColor.Black), 0, 0, 2, 2);
        var first = Assert.IsType<OfdPathElement>(page.Elements[0]);
        Assert.DoesNotContain("30 30", first.ClippingXml);
        graphics.Restore();
        graphics.FillRectangle(new OfdBrush(OfdColor.Black), 0, 0, 2, 2);
        var second = Assert.IsType<OfdPathElement>(page.Elements[1]);
        Assert.Equal(first.ClippingXml!.Split("AbbreviatedData")[1], second.ClippingXml!.Split("AbbreviatedData")[1]);
        Assert.NotEqual(first.XMillimeters, second.XMillimeters);
        Assert.Throws<InvalidOperationException>(() => graphics.Restore());
        Assert.Equal(2, page.Elements.Count);
    }

    [Fact]
    public void GraphemeAdvances_FontIdentity_AndAtomicFailures()
    {
        var (package, page, graphics) = Context();
        package.Fonts.Add(new OfdFontResource { Id = "21", FontName = "Actual Face" });
        var brush = new OfdBrush(OfdColor.Black);
        graphics.DrawString("Á中B", new OfdFont("Alias", 4, resourceId: "21"), brush, 1, 5, new[] { 2d, 3d });
        var text = Assert.IsType<OfdTextElement>(Assert.Single(page.Elements));
        Assert.Equal("Actual Face", text.FontName);
        Assert.Equal("2 3", Assert.Single(text.Runs).DeltaX);
        Assert.Throws<ArgumentException>(() => graphics.DrawString("Á中B", new OfdFont("Alias", 4), brush, 1, 5, new[] { 2d }));
        Assert.Throws<ArgumentException>(() => graphics.DrawString("X", new OfdFont("Alias", 4, resourceId: "missing"), brush, 1, 5));
        package.Fonts.Add(new OfdFontResource { Id = "21", FontName = "Duplicate" });
        Assert.Throws<ArgumentException>(() => graphics.DrawString("X", new OfdFont("Alias", 4, resourceId: "21"), brush, 1, 5));
        Assert.Single(page.Elements);
    }

    [Theory]
    [InlineData(0.0004, 0, 0, 1)]
    [InlineData(1, 1, 1, 1.0004)]
    public async Task Transform_SmallNonsingularCoefficients_RoundTripAtFullPrecision(double a, double b, double c, double d)
    {
        var (package, page, graphics) = Context();
        graphics.Rotate(90);
        graphics.Save();
        var matrix = new OfdMatrix(a, b, c, d, 0, 0);
        graphics.SetTransform(matrix);
        Assert.Same(matrix, graphics.Transform);
        graphics.Restore();
        graphics.SetTransform(matrix);
        graphics.DrawLine(new OfdPen(OfdColor.Black), 1, 2, 3, 4);
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var saved = await new OfdReader().ReadAsync(stream);
        var path = Assert.IsType<OfdPathElement>(Assert.Single(saved.Pages[0].Elements));
        Assert.Equal(new[] { a, b, c, d }, path.Transform!.Take(4));
    }

    [Fact]
    public void Budgets_Precision_Overflow_AndCancellation_DoNotAppend()
    {
        var (_, page, graphics) = Context(new OfdGraphicsOptions { MaxPageElements = 1, MaxPathCommands = 4, MaxTextCharacters = 2, MaxGeometryCharacters = 80 });
        var brush = new OfdBrush(OfdColor.Black);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfdPen(OfdColor.Black, 0.0001));
        Assert.Throws<ArgumentException>(() => new OfdMatrix(double.PositiveInfinity, 0, 0, 1, 0, 0));
        Assert.Throws<ArgumentException>(() => graphics.SetTransform(new OfdMatrix(0, 0, 0, 1, 0, 0)));
        graphics.Scale(2, 1);
        Assert.Throws<ArgumentException>(() => graphics.DrawLine(new OfdPen(OfdColor.Black), double.MaxValue, 0, double.MaxValue, 1));
        graphics.SetTransform(OfdMatrix.Identity);
        using var canceled = new CancellationTokenSource(); canceled.Cancel();
        Assert.Throws<OperationCanceledException>(() => graphics.DrawString("X", new OfdFont("F", 4), brush, 1, 5, cancellationToken: canceled.Token));
        Assert.Empty(page.Elements);
        graphics.DrawString("OK", new OfdFont("F", 4), brush, 1, 5);
        Assert.Throws<InvalidOperationException>(() => graphics.DrawString("X", new OfdFont("F", 4), brush, 1, 5));
        Assert.Single(page.Elements);
    }
}
