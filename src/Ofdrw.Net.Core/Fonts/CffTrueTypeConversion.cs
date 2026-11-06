using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>Builds a TrueType (glyf) rendition of a static OpenType/CFF1 face for consumers that can
/// only embed TrueType programs (PDF FontFile2). GIDs, cmap, horizontal metrics, OS/2 and name are
/// unchanged; each cubic segment becomes quadratic segments within one design unit.</summary>
/// <remarks>The OFD keeps the original CFF bytes. This rendition only serves export.</remarks>
internal static class CffTrueTypeConversion
{
    private const double Tolerance = 1.0;
    private const int MaximumPieces = 16;
    private static readonly string[] DroppedTables = ["CFF ", "VORG", "DSIG", "hdmx", "LTSH", "VDMX"];

    internal static byte[] Convert(byte[] faceBytes, CancellationToken cancellationToken = default)
    {
        var face = new OpenTypeFace(faceBytes);
        var outlines = new CffCharStringOutlines(faceBytes);
        var count = outlines.GlyphCount;
        using var glyf = new MemoryStream();
        var loca = new byte[(count + 1) * 4];
        int maxPoints = 0, maxContours = 0;
        for (var gid = 0; gid < count; gid++)
        {
            if ((gid & 255) == 0) cancellationToken.ThrowIfCancellationRequested();
            OpenTypeFace.Put32(loca, gid * 4, (uint)glyf.Length);
            if (gid == 0) continue; // .notdef stays empty; it is never drawn.
            var outline = outlines.Read(gid, cancellationToken);
            var contours = Quadratic(outline);
            if (contours.Count == 0) continue;
            var points = contours.Sum(contour => contour.Count);
            if (points > 65535 || contours.Count > 32767) throw new NotSupportedException("A CFF glyph is too complex for TrueType.");
            maxPoints = Math.Max(maxPoints, points); maxContours = Math.Max(maxContours, contours.Count);
            var bytes = Encode(contours);
            glyf.Write(bytes, 0, bytes.Length);
            while (glyf.Length % 4 != 0) glyf.WriteByte(0);
        }
        OpenTypeFace.Put32(loca, count * 4, (uint)glyf.Length);
        foreach (var tag in DroppedTables) face.Tables.Remove(tag);
        face.Tables["glyf"] = glyf.ToArray();
        face.Tables["loca"] = loca;
        var head = (byte[])face.Table("head", 54).Clone();
        OpenTypeFace.Put16(head, 50, 1);
        face.Tables["head"] = head;
        var maxp = new byte[32];
        OpenTypeFace.Put32(maxp, 0, 0x00010000);
        OpenTypeFace.Put16(maxp, 4, count);
        OpenTypeFace.Put16(maxp, 6, maxPoints);
        OpenTypeFace.Put16(maxp, 8, maxContours);
        OpenTypeFace.Put16(maxp, 14, 2); // maxZones
        face.Tables["maxp"] = maxp;
        // post format 3: no glyph names, which a converted CID font would not have anyway.
        if (face.Tables.TryGetValue("post", out var post) && post.Length >= 32)
        {
            var header = new byte[32];
            Array.Copy(post, header, 32);
            OpenTypeFace.Put32(header, 0, 0x00030000);
            face.Tables["post"] = header;
        }
        return face.Build(0x00010000);
    }

    // Each cubic becomes n quadratics (n from the cubic's third difference) using the midpoint control
    // (3(P1 + P2) − (P0 + P3)) / 4 of each de Casteljau piece.
    private static List<List<(int X, int Y, bool On)>> Quadratic(TrueTypeGlyphOutline outline)
    {
        var result = new List<List<(int X, int Y, bool On)>>();
        foreach (var command in Group(outline.ToPathCommands()))
        {
            var contour = new List<(int X, int Y, bool On)>();
            (double X, double Y) pen = (0, 0);
            foreach (var segment in command)
            {
                switch (segment.Verb)
                {
                    case TrueTypePathVerb.Move:
                        pen = (segment.X, segment.Y);
                        Add(contour, pen, true);
                        break;
                    case TrueTypePathVerb.Line:
                        pen = (segment.X, segment.Y);
                        Add(contour, pen, true);
                        break;
                    case TrueTypePathVerb.Quadratic:
                        Add(contour, (segment.ControlX, segment.ControlY), false);
                        pen = (segment.X, segment.Y);
                        Add(contour, pen, true);
                        break;
                    case TrueTypePathVerb.Cubic:
                    {
                        var p0 = pen; var p1 = (segment.ControlX, segment.ControlY);
                        var p2 = (segment.Control2X, segment.Control2Y); var p3 = (segment.X, segment.Y);
                        var dx = p3.Item1 - 3 * p2.Item1 + 3 * p1.Item1 - p0.X;
                        var dy = p3.Item2 - 3 * p2.Item2 + 3 * p1.Item2 - p0.Y;
                        var error = Math.Sqrt(3) / 36 * Math.Sqrt(dx * dx + dy * dy);
                        var pieces = Math.Max(1, Math.Min(MaximumPieces, (int)Math.Ceiling(Math.Pow(error / Tolerance, 1.0 / 3))));
                        for (var piece = 0; piece < pieces; piece++)
                        {
                            var (a, b, c, d) = Slice(p0, p1, p2, p3, (double)piece / pieces, (double)(piece + 1) / pieces);
                            Add(contour, ((3 * (b.X + c.X) - a.X - d.X) / 4, (3 * (b.Y + c.Y) - a.Y - d.Y) / 4), false);
                            Add(contour, d, true);
                        }
                        pen = p3;
                        break;
                    }
                }
            }
            // The contour closes implicitly; drop a final on-curve point equal to the start.
            if (contour.Count > 1 && contour[contour.Count - 1].On && contour[contour.Count - 1].X == contour[0].X &&
                contour[contour.Count - 1].Y == contour[0].Y)
                contour.RemoveAt(contour.Count - 1);
            if (contour.Count > 1) result.Add(contour);
        }
        return result;
    }

    private static IEnumerable<List<TrueTypePathCommand>> Group(IReadOnlyList<TrueTypePathCommand> commands)
    {
        List<TrueTypePathCommand>? current = null;
        foreach (var command in commands)
        {
            if (command.Verb == TrueTypePathVerb.Move) { if (current is not null) yield return current; current = new() { command }; }
            else if (command.Verb == TrueTypePathVerb.Close) { if (current is not null) yield return current; current = null; }
            else current?.Add(command);
        }
        if (current is not null) yield return current;
    }

    private static void Add(List<(int X, int Y, bool On)> contour, (double X, double Y) point, bool on)
    {
        var x = (int)Math.Round(point.X); var y = (int)Math.Round(point.Y);
        if (x is < short.MinValue or > short.MaxValue || y is < short.MinValue or > short.MaxValue)
            throw new NotSupportedException("A CFF point exceeds the TrueType coordinate range.");
        // Collapse a repeated on-curve point (zero-length segment after rounding).
        if (on && contour.Count > 0 && contour[contour.Count - 1].On && contour[contour.Count - 1].X == x &&
            contour[contour.Count - 1].Y == y) return;
        contour.Add((x, y, on));
    }

    private static ((double X, double Y), (double X, double Y), (double X, double Y), (double X, double Y)) Slice(
        (double X, double Y) p0, (double X, double Y) p1, (double X, double Y) p2, (double X, double Y) p3, double t0, double t1)
    {
        (double X, double Y) At(double t)
        {
            var u = 1 - t;
            return (u * u * u * p0.X + 3 * u * u * t * p1.X + 3 * u * t * t * p2.X + t * t * t * p3.X,
                u * u * u * p0.Y + 3 * u * u * t * p1.Y + 3 * u * t * t * p2.Y + t * t * t * p3.Y);
        }
        (double X, double Y) Derivative(double t)
        {
            var u = 1 - t;
            return (3 * (u * u * (p1.X - p0.X) + 2 * u * t * (p2.X - p1.X) + t * t * (p3.X - p2.X)),
                3 * (u * u * (p1.Y - p0.Y) + 2 * u * t * (p2.Y - p1.Y) + t * t * (p3.Y - p2.Y)));
        }
        var h = (t1 - t0) / 3;
        var a = At(t0); var d = At(t1); var da = Derivative(t0); var dd = Derivative(t1);
        return (a, (a.X + da.X * h, a.Y + da.Y * h), (d.X - dd.X * h, d.Y - dd.Y * h), d);
    }

    private static byte[] Encode(List<List<(int X, int Y, bool On)>> contours)
    {
        var points = contours.SelectMany(contour => contour).ToArray();
        using var output = new MemoryStream();
        void W16(int value) { output.WriteByte((byte)(value >> 8)); output.WriteByte((byte)value); }
        W16(contours.Count);
        W16(points.Min(point => point.X)); W16(points.Min(point => point.Y));
        W16(points.Max(point => point.X)); W16(points.Max(point => point.Y));
        var end = -1;
        foreach (var contour in contours) { end += contour.Count; W16(end); }
        W16(0); // no instructions
        var flags = new byte[points.Length];
        var xs = new List<byte>(); var ys = new List<byte>();
        int px = 0, py = 0;
        for (var index = 0; index < points.Length; index++)
        {
            var flag = points[index].On ? 1 : 0;
            var dx = points[index].X - px; var dy = points[index].Y - py;
            px = points[index].X; py = points[index].Y;
            if (dx == 0) flag |= 0x10;
            else if (Math.Abs(dx) < 256) { flag |= 0x02 | (dx > 0 ? 0x10 : 0); xs.Add((byte)Math.Abs(dx)); }
            else { xs.Add((byte)(dx >> 8)); xs.Add((byte)dx); }
            if (dy == 0) flag |= 0x20;
            else if (Math.Abs(dy) < 256) { flag |= 0x04 | (dy > 0 ? 0x20 : 0); ys.Add((byte)Math.Abs(dy)); }
            else { ys.Add((byte)(dy >> 8)); ys.Add((byte)dy); }
            flags[index] = (byte)flag;
        }
        output.Write(flags, 0, flags.Length);
        output.Write(xs.ToArray(), 0, xs.Count);
        output.Write(ys.ToArray(), 0, ys.Count);
        return output.ToArray();
    }
}
