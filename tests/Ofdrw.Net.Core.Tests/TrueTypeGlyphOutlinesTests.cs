using System;
using System.Buffers.Binary;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Globalization;
using System.Text;
using Ofdrw.Net.Core.Fonts;

namespace Ofdrw.Net.Core.Tests;

public sealed class TrueTypeGlyphOutlinesTests
{
    // Independent fontTools 4.59.2 getCoordinates oracle over the pinned OFL
    // NotoSans-Regular.ttf: signed big-endian x/y, one on-curve byte per point,
    // then big-endian contour end indices. 'fl' is a GSUB GID without a cmap
    // scalar; é also exercises a separate accent component.
    [Theory]
    [InlineData(36, 18, 2, 639, "87b232a19336c343b39518ca59c94dc2743de3a8bd8d00c086a937b1a4595426")]
    [InlineData(1653, 28, 2, 602, "165e09631437760eb62a355f43e5863cca58eb110d05186dabefbd5cbbefd317")]
    [InlineData(171, 43, 3, 564, "55ce309c6236a9598dff429b59668ec927fc9f5c9b1f22af06426e6e98dbfe60")]
    public void OriginalGidReturnsTheIndependentUnhintedOutline(int gid, int pointCount, int contourCount,
        int advance, string expectedSha256)
    {
        var bytes = Fixture();
        var reader = new TrueTypeGlyphOutlines(bytes);
        var outline = reader.Read(gid);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), outline.SourceSha256);
        Assert.Equal(outline.SourceSha256, outline.SelectedFaceSha256);
        Assert.Equal(0, outline.FaceIndex);
        Assert.Equal(gid, outline.GlyphId);
        Assert.Equal(1000, outline.UnitsPerEm);
        Assert.Equal(advance, outline.AdvanceWidth);
        Assert.Equal(contourCount, outline.Contours.Count);
        Assert.Equal(pointCount, outline.Contours.Sum(c => c.Points.Count));
        Assert.Equal(expectedSha256, OracleHash(outline));
        var path = outline.ToPathCommands();
        Assert.Equal(contourCount, path.Count(command => command.Verb == TrueTypePathVerb.Move));
        Assert.Equal(contourCount, path.Count(command => command.Verb == TrueTypePathVerb.Close));
        Assert.All(path, command => Assert.True(double.IsFinite(command.X) && double.IsFinite(command.Y)));
        if (gid == 171) Assert.Contains(path, command => command.Verb == TrueTypePathVerb.Quadratic);
        WriteVisualProbe(gid, path);
    }

    [Fact]
    public void CallerMutationCannotChangeTheSelectedOriginalGid()
    {
        var bytes = Fixture();
        var reader = new TrueTypeGlyphOutlines(bytes);
        bytes.AsSpan().Clear();
        Assert.Equal("165e09631437760eb62a355f43e5863cca58eb110d05186dabefbd5cbbefd317",
            OracleHash(reader.Read(1653)));
    }

    [Fact]
    public void ExplicitCollectionFaceIndexRetainsSourceAndSelectedFaceIdentity()
    {
        var source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "testdata", "style-collection.ttc"));
        var reader = new TrueTypeGlyphOutlines(source, faceIndex: 0);
        var outline = reader.Read(2);
        Assert.Equal(0, outline.FaceIndex);
        Assert.Equal(Convert.ToHexString(SHA256.HashData(source)).ToLowerInvariant(), outline.SourceSha256);
        Assert.NotEqual(outline.SourceSha256, outline.SelectedFaceSha256);
        Assert.Equal(500, outline.AdvanceWidth);
        Assert.Single(outline.Contours);
        Assert.Equal("ccb1267561e0ca4aa464c3c056ec7f2d6c2a5e62917c18bd50db393d8bdc1660", OracleHash(outline));
        Assert.Throws<InvalidDataException>(() => new TrueTypeGlyphOutlines(source, faceIndex: 1));
    }

    [Fact]
    public void UnsupportedPointAttachedCompositeFailsBeforeProducingAnOutline()
    {
        var font = new OpenTypeFace(Fixture());
        var loca = font.Table("loca");
        var glyf = font.Table("glyf");
        var start = checked((int)OpenTypeFace.U32(loca, 1653 * 4));
        var flags = OpenTypeFace.U16(glyf, start + 10);
        Assert.NotEqual(0, flags & 2);
        OpenTypeFace.Put16(glyf, start + 10, flags & ~2);
        var reader = new TrueTypeGlyphOutlines(font.Build());
        Assert.Contains("point-attached", Assert.Throws<NotSupportedException>(() => reader.Read(1653)).Message);

        var transformed = new OpenTypeFace(Fixture());
        OpenTypeFace.Put16(transformed.Table("glyf"), start + 10, flags | 0x0008);
        var scaledReader = new TrueTypeGlyphOutlines(transformed.Build());
        Assert.Contains("Transformed", Assert.Throws<NotSupportedException>(() => scaledReader.Read(1653)).Message);
    }

    [Fact]
    public void MissingOrCancelledGidDoesNotReturnAnOutline()
    {
        var reader = new TrueTypeGlyphOutlines(Fixture());
        Assert.Throws<InvalidDataException>(() => new TrueTypeGlyphOutlines(Fixture(), faceIndex: 1));
        Assert.Throws<NotSupportedException>(() => reader.Read(0)); // .notdef is not a drawable requested GID.
        Assert.Throws<NotSupportedException>(() => reader.Read(reader.GlyphCount));
        var space = reader.Read(3);
        Assert.Empty(space.Contours); // A valid empty glyph differs from .notdef.
        Assert.True(space.AdvanceWidth > 0);
        using var cancel = new System.Threading.CancellationTokenSource();
        cancel.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => reader.Read(36, cancel.Token));
    }

    private static byte[] Fixture() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory,
        "testdata", "NotoSans-Regular.ttf"));

    private static void WriteVisualProbe(int gid, System.Collections.Generic.IReadOnlyList<TrueTypePathCommand> path)
    {
        var folder = Environment.GetEnvironmentVariable("R08_GID_OUTLINE_SVG_DIR");
        if (string.IsNullOrWhiteSpace(folder)) return;
        Directory.CreateDirectory(folder);
        var data = new StringBuilder();
        static string N(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        foreach (var command in path)
        {
            switch (command.Verb)
            {
                case TrueTypePathVerb.Move: data.Append("M ").Append(N(command.X)).Append(' ').Append(N(command.Y)).Append(' '); break;
                case TrueTypePathVerb.Line: data.Append("L ").Append(N(command.X)).Append(' ').Append(N(command.Y)).Append(' '); break;
                case TrueTypePathVerb.Quadratic:
                    data.Append("Q ").Append(N(command.ControlX)).Append(' ').Append(N(command.ControlY)).Append(' ')
                        .Append(N(command.X)).Append(' ').Append(N(command.Y)).Append(' '); break;
                case TrueTypePathVerb.Close: data.Append("Z "); break;
            }
        }
        var svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"120\" height=\"110\" viewBox=\"0 0 120 110\">" +
            "<rect width=\"120\" height=\"110\" fill=\"white\"/>" +
            "<path transform=\"translate(10 90) scale(.1 -.1)\" fill=\"#225c9c\" fill-rule=\"nonzero\" d=\"" +
            data + "\"/></svg>";
        File.WriteAllText(Path.Combine(folder, "gid-" + gid + ".svg"), svg);
    }

    private static string OracleHash(TrueTypeGlyphOutline outline)
    {
        using var stream = new MemoryStream();
        Span<byte> field = stackalloc byte[4];
        var endpoint = -1;
        foreach (var contour in outline.Contours)
        {
            foreach (var point in contour.Points)
            {
                Assert.Equal(Math.Truncate(point.X), point.X);
                Assert.Equal(Math.Truncate(point.Y), point.Y);
                BinaryPrimitives.WriteInt32BigEndian(field, checked((int)point.X)); stream.Write(field);
                BinaryPrimitives.WriteInt32BigEndian(field, checked((int)point.Y)); stream.Write(field);
                stream.WriteByte(point.OnCurve ? (byte)1 : (byte)0);
                endpoint++;
            }
        }
        // FontTools' endPtsOfContours refer to the flattened point stream.
        endpoint = -1;
        foreach (var contour in outline.Contours)
        {
            endpoint += contour.Points.Count;
            BinaryPrimitives.WriteUInt16BigEndian(field, checked((ushort)endpoint));
            stream.Write(field[..2]);
        }
        return Convert.ToHexString(SHA256.HashData(stream.ToArray())).ToLowerInvariant();
    }
}
