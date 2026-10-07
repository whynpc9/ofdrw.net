using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Graphics.SkiaSharp;
using Ofdrw.Net.Layout.Graphics;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using SkiaSharp;

namespace Ofdrw.Net.Graphics.SkiaSharp.Tests;

public class AdapterTests
{
    private static (OfdDocumentPackage Package, OfdPage Page) Target()
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 148, HeightMillimeters = 210 }; package.Pages.Add(page); return (package, page);
    }
    private static SKPaint Pen() => new() { Style = SKPaintStyle.Stroke, StrokeWidth = 0.5f, Color = SKColors.Blue, StrokeMiter = 10 };
    private static byte[] FontBytes(string name = "narrow") => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", name + ".ttf"));
    private static SkiaDrawEvent Text(OfdDocumentPackage package, string text = "A  A", string file = "narrow", string id = "face")
    {
        var bytes = FontBytes(file); using var data = SKData.CreateCopy(bytes); using var typeface = SKTypeface.FromData(data); using var font = new SKFont(typeface, 4); using var fill = new SKPaint { Color = SKColors.Red };
        package.Fonts.Add(new OfdFontResource { Id = id, FontName = typeface.FamilyName, Data = bytes, Bold = typeface.IsBold, Italic = typeface.IsItalic });
        var advances = Enumerable.Repeat(2f, System.Globalization.StringInfo.ParseCombiningCharacters(text).Length - 1).ToArray();
        return SkiaDrawEvent.Text(text, new SKPoint(12, 20), font, fill, id, advances);
    }
    [Fact]
    public async Task NativeRoundtripPreservesKindsTextAndBindingAfterInputsDisposed()
    {
        var (package, page) = Target(); SkiaDrawEvent line, pathEvent;
        using (var pen = Pen())
        using (var path = new SKPath())
        {
            path.MoveTo(3, 4); path.QuadTo(5, 6, 7, 8); path.CubicTo(9, 10, 11, 12, 13, 14);
            line = SkiaDrawEvent.Line(new SKPoint(1, 2), new SKPoint(3, 4), pen);
            pathEvent = SkiaDrawEvent.Path(path, pen);
            path.Reset(); pen.Color = SKColors.Green;
        }
        var text = Text(package); OfdSkiaAdapter.Append(package, page, new[] { line, pathEvent, text });
        Assert.Equal(3, page.Elements.Count); Assert.Equal(255, Assert.IsType<OfdPathElement>(page.Elements[0]).StrokeColor!.Blue);
        await using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, output); output.Position = 0;
        var read = await new OfdReader().ReadAsync(output);
        Assert.Equal(2, read.Pages[0].Elements.OfType<OfdPathElement>().Count());
        var actual = Assert.Single(read.Pages[0].Elements.OfType<OfdTextElement>()); Assert.Equal("A  A", actual.Text);
        Assert.Equal(Assert.Single(read.Fonts).Id, actual.FontResourceId); Assert.Empty(read.Pages[0].Elements.OfType<OfdImageElement>());
    }
    [Fact]
    public void LateWrongFaceFailsWholeBatchWithoutChangingResourcesOrPriorElements()
    {
        var (package, page) = Target(); var text = Text(package); package.Fonts[0].Data = FontBytes("wide");
        using var pen = Pen(); var prior = new OfdPathElement(); page.Elements.Add(prior);
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen), text }));
        Assert.Same(prior, Assert.Single(page.Elements)); Assert.Single(package.Fonts); Assert.Equal(FontBytes("wide"), package.Fonts[0].Data);
    }
    [Theory]
    [InlineData(false)] [InlineData(true)]
    public void MissingOrDuplicateIdIsAtomic(bool duplicate)
    {
        var (package, page) = Target(); var text = Text(package);
        if (duplicate) package.Fonts.Add(package.Fonts[0]); else package.Fonts.Clear();
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { text })); Assert.Empty(page.Elements);
    }
    [Fact]
    public void StyleMismatchAndAbsentFontDataFail()
    {
        var (package, page) = Target(); var text = Text(package); package.Fonts[0].Bold = !package.Fonts[0].Bold;
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { text }));
        package.Fonts[0].Bold = !package.Fonts[0].Bold; package.Fonts[0].Data = [];
        Assert.Throws<ArgumentException>(() => OfdSkiaAdapter.Append(package, page, new[] { text })); Assert.Empty(page.Elements);
    }
    [Fact]
    public void EventAndExistingPageAndCumulativePathBudgetsAreAtomic()
    {
        using var pen = Pen(); var line = SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen); var (package, page) = Target();
        Assert.Throws<InvalidOperationException>(() => OfdSkiaAdapter.Append(package, page, new[] { line, line }, new() { MaxEvents = 1 })); Assert.Empty(page.Elements);
        Assert.Throws<InvalidOperationException>(() => OfdSkiaAdapter.Append(package, page, new[] { line, line }, new() { MaxInputPathCommands = 3 })); Assert.Empty(page.Elements);
        page.Elements.Add(new OfdPathElement());
        Assert.Throws<InvalidOperationException>(() => OfdSkiaAdapter.Append(package, page, new[] { line }, new() { Graphics = new() { MaxPageElements = 1 } })); Assert.Single(page.Elements);
    }
    [Fact]
    public void EnumeratorFailureAndLateCancellationAppendNothing()
    {
        using var pen = Pen(); var line = SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen); var (package, page) = Target(); using var cancellation = new CancellationTokenSource();
        IEnumerable<SkiaDrawEvent> Failure() { yield return line; throw new IOException("producer failed"); }
        IEnumerable<SkiaDrawEvent> Cancel() { yield return line; cancellation.Cancel(); }
        Assert.Throws<IOException>(() => OfdSkiaAdapter.Append(package, page, Failure())); Assert.Empty(page.Elements);
        Assert.Throws<OperationCanceledException>(() => OfdSkiaAdapter.Append(package, page, Cancel(), cancellationToken: cancellation.Token)); Assert.Empty(page.Elements);
    }
    [Fact]
    public void AffineUsesUtimesMWithFullStrokeAndTextScale()
    {
        var (package, page) = Target(); using var pen = Pen(); var matrix = new SKMatrix(2, .2f, 10, .3f, 1, 20, 0, 0, 1);
        var line = SkiaDrawEvent.Line(new(0, 0), new(20, 0), pen, matrix); OfdSkiaAdapter.Append(package, page, new[] { line }, new() { MillimetersPerUnit = .5 });
        var path = Assert.IsType<OfdPathElement>(page.Elements[0]); Assert.Equal(1, path.Transform![0]); Assert.Equal(.5, path.Transform[3]);
        Assert.Equal(.15, path.Transform[1], 6); Assert.Equal(5, path.Transform[4] + path.XMillimeters, 6); Assert.Equal(10, path.Transform[5] + path.YMillimeters, 6);
    }
    [Fact]
    public void DefaultMiterAcceptedForLineAndRectangleButArbitraryPathRequires10()
    {
        using var pen = Pen(); pen.StrokeMiter = 4;
        SkiaDrawEvent.Line(new(1, 1), new(2, 2), pen); SkiaDrawEvent.Rectangle(new(1, 1, 4, 4), pen);
        using var path = new SKPath(); path.MoveTo(1, 1); path.LineTo(2, 3); path.LineTo(3, 1);
        Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Path(path, pen));
        pen.StrokeMiter = 1; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Rectangle(new(1, 1, 4, 4), pen));
    }
    [Fact]
    public void UnsupportedPaintFieldsNeverSilentlyDrop()
    {
        using var pen = Pen(); var p = new SKPoint(1, 1); var q = new SKPoint(2, 2);
        using var shader = SKShader.CreateLinearGradient(p, q, new[] { SKColors.Red, SKColors.Blue }, SKShaderTileMode.Clamp);
        pen.Shader = shader; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(p, q, pen)); pen.Shader = null;
        using var color = SKColorFilter.CreateColorMatrix(new float[] {1,0,0,0,0,0,1,0,0,0,0,0,1,0,0,0,0,0,1,0});
        pen.ColorFilter = color; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(p, q, pen)); pen.ColorFilter = null;
        using var image = SKImageFilter.CreateBlur(1, 1); pen.ImageFilter = image; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(p, q, pen)); pen.ImageFilter = null;
        using var mask = SKMaskFilter.CreateBlur(SKBlurStyle.Normal, 1); pen.MaskFilter = mask; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(p, q, pen)); pen.MaskFilter = null;
        using var effect = SKPathEffect.CreateDash(new[] {1f, 2f}, 0); pen.PathEffect = effect; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(p, q, pen)); pen.PathEffect = null;
        pen.BlendMode = SKBlendMode.Multiply; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(p, q, pen));
        Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Unsupported("ClipPath"));
    }
    [Fact]
    public void HighPrecisionColorIsNotSilentlyQuantized()
    {
        using var pen = Pen(); pen.ColorF = new SKColorF(.1f, .2f, .3f, 1);
        Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen));
    }
    [Fact]
    public void InverseConicPerspectiveSingularAndHairlineFail()
    {
        using var pen = Pen(); using var path = new SKPath(); path.MoveTo(1, 1); path.ConicTo(2, 3, 4, 5, .5f);
        Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Path(path, pen)); path.FillType = SKPathFillType.InverseEvenOdd; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Path(path, pen));
        Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen, new SKMatrix(1,0,0,0,1,0,.1f,0,1)));
        Assert.Throws<ArgumentException>(() => SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen, SKMatrix.CreateScale(0, 1)));
        pen.StrokeWidth = 0; Assert.Throws<ArgumentOutOfRangeException>(() => SkiaDrawEvent.Line(new(1, 2), new(3, 4), pen));
    }
    [Fact]
    public void TextAdvanceAndSnapshotBudgetsAndSyntheticStylesFailAtEntry()
    {
        using var data = SKData.CreateCopy(FontBytes()); using var face = SKTypeface.FromData(data); using var font = new SKFont(face, 4); using var paint = new SKPaint();
        Assert.Throws<ArgumentException>(() => SkiaDrawEvent.Text("AA", new(1, 2), font, paint, "face", []));
        Assert.Throws<InvalidOperationException>(() => SkiaDrawEvent.Text("AAA", new(1, 2), font, paint, "face", [1, 1], maxTextCharacters: 2));
        Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Text("AA", new(1, 2), font, paint, "face", [1], maxFontBytes: 1));
        font.Embolden = true; Assert.Throws<NotSupportedException>(() => SkiaDrawEvent.Text("AA", new(1, 2), font, paint, "face", [1]));
    }
    [Fact]
    public void PathSnapshotBudgetAndCloseThenLineRetainGeometry()
    {
        using var pen = Pen(); using var path = new SKPath(); path.MoveTo(1, 1); path.LineTo(2, 2); path.LineTo(3, 1); path.Close(); path.LineTo(5, 6);
        Assert.Throws<InvalidOperationException>(() => SkiaDrawEvent.Path(path, pen, maxPathCommands: 2));
        var (package, page) = Target(); OfdSkiaAdapter.Append(package, page, new[] { SkiaDrawEvent.Path(path, pen) });
        Assert.Contains("5 6", Assert.IsType<OfdPathElement>(page.Elements[0]).AbbreviatedData);
    }
}
