using Ofdrw.Net.Converter.Docx.Internal.BuiltIn;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Converter.Docx.Tests;

public sealed class NativeBoldStrokePlannerTests
{
    [Fact]
    public void NestedCounterCannotBorrowTheLargeInnerWhiteRegion()
    {
        var original = Font();
        var planner = new NativeBoldStrokePlanner(default);
        Assert.True(planner.Width(original, false, 4, "中") > .04); // plain rectangle control
        var nested = CounterFont(original, nested: true);
        var snapshot = nested.ToArray();
        // The outer white ring is only10 design units wide. Every point in
        // its9x9 grid lies inside the large nested inner hole. Those points
        // belong to the inner counter and cannot justify thickening the ring.
        Assert.Equal(.04, planner.Width(nested, false, 4, "中"));
        Assert.Equal(snapshot, nested);
    }

    [Fact]
    public void OpticalMinimumOnlyAllowsAnIncreaseWhenTheSerializedSizePreservesItsOpening()
    {
        var face = CounterFont(Font(), nested: false);
        var planner = new NativeBoldStrokePlanner(default);
        const double small = 25.4 * 6 / 72, medium = 25.4 * 12 / 72, large = 25.4 * 18 / 72;
        Assert.Equal(small * .01, planner.Width(face, false, small, "中"));
        var mediumWidth = planner.Width(face, false, medium, "中");
        var largeWidth = planner.Width(face, false, large, "中");
        Assert.InRange(mediumWidth, .139, .141);
        Assert.InRange(largeWidth, .221, .223);
        // The rectangular hole is100/1000em high. These are independent
        // analytic clearances at the serialized12pt/18pt sizes, including the
        // stricter target before the planner subtracts its sampling allowance.
        var mediumRadius = 4.233 * 72 / 25.4 * .05;
        Assert.True(mediumRadius - mediumWidth * 72 / 25.4 / 2 >= .4005);
        Assert.True(.9 - largeWidth * 72 / 25.4 / 2 >= .5405);
        // The6pt hole cannot meet0.4pt even without a stroke. The legacy
        // fallback above does not claim to satisfy the increased-width guard.
    }

    [Fact]
    public void MissingGlyphAndUnsupportedFontKeepTheLegacyStroke()
    {
        var original = Font();
        var planner = new NativeBoldStrokePlanner(default);
        Assert.Equal(.04, planner.Width(original, false, 4, "量"));
        var variable = new OpenTypeFace(original);
        variable.Tables["fvar"] = new byte[16];
        Assert.Equal(.04, planner.Width(variable.Build(), false, 4, "中"));
        Assert.Equal(.04, planner.Width(original, true, 4, "中"));
    }

    [Fact]
    public void CancellationIsNotConvertedIntoAnOpticalFallback()
    {
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        var planner = new NativeBoldStrokePlanner(cancellation.Token);
        Assert.Throws<OperationCanceledException>(() => planner.Width(Font(), false, 4, "中"));
    }

    private static byte[] Font()
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!File.Exists(Path.Combine(root!.FullName, "Ofdrw.Net.sln"))) root = root.Parent;
        return File.ReadAllBytes(Path.Combine(root.FullName, "e2e/Ofdrw.Net.Converter.Pdf.E2E/testdata/fonts/narrow.ttf"));
    }

    private static byte[] CounterFont(byte[] original, bool nested)
    {
        var face = new OpenTypeFace(original); var gid = new OpenTypeCmap(face).DirectGlyph('中');
        (int X, int Y)[] points = nested ? new (int, int)[]
        {
            (0,0),(0,1000),(1000,1000),(1000,0),
            (100,100),(900,100),(900,900),(100,900),
            (110,110),(110,890),(890,890),(890,110),
            (150,150),(850,150),(850,850),(150,850)
        } : new (int, int)[]
        {
            (0,0),(0,1000),(1000,1000),(1000,0),
            (100,450),(900,450),(900,550),(100,550)
        };
        var contours = points.Length / 4;
        var flags = 12 + contours * 2;
        var xCoordinates = flags + points.Length;
        var yCoordinates = xCoordinates + points.Length * 2;
        var glyph = new byte[yCoordinates + points.Length * 2];
        OpenTypeFace.Put16(glyph, 0, contours); OpenTypeFace.Put16(glyph, 6, 1000); OpenTypeFace.Put16(glyph, 8, 1000);
        for (var i = 0; i < contours; i++) OpenTypeFace.Put16(glyph, 10 + i * 2, i * 4 + 3);
        for (var i = 0; i < points.Length; i++)
        {
            glyph[flags + i] = 1;
            OpenTypeFace.Put16(glyph, xCoordinates + i * 2, unchecked((ushort)(points[i].X - (i == 0 ? 0 : points[i - 1].X))));
            OpenTypeFace.Put16(glyph, yCoordinates + i * 2, unchecked((ushort)(points[i].Y - (i == 0 ? 0 : points[i - 1].Y))));
        }
        var loca = face.Table("loca"); var shortLoca = OpenTypeFace.U16(face.Table("head"), 50) == 0;
        var count = OpenTypeFace.U16(face.Table("maxp"), 4);
        int Offset(int index) => shortLoca ? OpenTypeFace.U16(loca, index * 2) * 2 : (int)OpenTypeFace.U32(loca, index * 4);
        var start = Offset(gid); var end = Offset(gid + 1); var old = face.Table("glyf");
        var replacement = new byte[old.Length + glyph.Length - (end - start)];
        Buffer.BlockCopy(old, 0, replacement, 0, start); Buffer.BlockCopy(glyph, 0, replacement, start, glyph.Length);
        Buffer.BlockCopy(old, end, replacement, start + glyph.Length, old.Length - end);
        face.Tables["glyf"] = replacement;
        for (var i = count; i > gid; i--)
        {
            var offset = Offset(i) + glyph.Length - (end - start);
            if (shortLoca) OpenTypeFace.Put16(loca, i * 2, offset / 2);
            else OpenTypeFace.Put32(loca, i * 4, (uint)offset);
        }
        OpenTypeFace.Put16(face.Table("maxp"), 6, Math.Max(OpenTypeFace.U16(face.Table("maxp"), 6), points.Length));
        OpenTypeFace.Put16(face.Table("maxp"), 8, Math.Max(OpenTypeFace.U16(face.Table("maxp"), 8), contours));
        return face.Build();
    }
}
