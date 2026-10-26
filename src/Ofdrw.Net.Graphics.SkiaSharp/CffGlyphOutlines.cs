using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Threading;
using Ofdrw.Net.Core.Fonts;
using SkiaSharp;

namespace Ofdrw.Net.Graphics.SkiaSharp;

// Optional native backend for one original GID in a standalone static CFF1
// OpenType face. Logical TextCode and glyph positioning belong to the caller.
internal sealed class CffGlyphOutlines : IDisposable
{
    private const int MaximumFaceBytes = 16 * 1024 * 1024;
    private const int MaximumCommands = 32768;
    private static readonly string[] UnsupportedTables =
    [
        "CFF2", "glyf", "loca", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
        "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "EBDT", "EBLC", "sbix"
    ];

    private readonly SKTypeface face;
    private readonly byte[] hmtx;
    private readonly int metricCount;
    private readonly object gate = new();
    private bool disposed;

    internal CffGlyphOutlines(byte[] sourceBytes)
    {
        if (sourceBytes is null) throw new ArgumentNullException(nameof(sourceBytes));
        if (sourceBytes.Length is < 12 or > MaximumFaceBytes)
            throw new NotSupportedException("CFF GID outlines require a bounded standalone OpenType face.");
        var original = (byte[])sourceBytes.Clone();
        var font = new OpenTypeFace(original);
        if (font.Version != 0x4F54544F || !font.Tables.ContainsKey("CFF ") ||
            UnsupportedTables.Any(font.Tables.ContainsKey))
            throw new NotSupportedException("CFF GID outlines require a static monochrome CFF1 OpenType face.");
        var cff = font.Table("CFF ", 4);
        if (cff[0] != 1)
            throw new NotSupportedException("CFF GID outlines require a CFF1 table.");
        var head = font.Table("head", 54);
        UnitsPerEm = OpenTypeFace.U16(head, 18);
        GlyphCount = OpenTypeFace.U16(font.Table("maxp", 6), 4);
        if (UnitsPerEm is < 16 or > 16384 || GlyphCount == 0)
            throw new InvalidDataException("Invalid CFF face metrics.");
        metricCount = OpenTypeFace.U16(font.Table("hhea", 36), 34);
        if (metricCount == 0 || metricCount > GlyphCount)
            throw new InvalidDataException("Invalid CFF horizontal metric count.");
        hmtx = font.Table("hmtx", checked(metricCount * 4 + (GlyphCount - metricCount) * 2));
        using var data = SKData.CreateCopy(original);
        face = SKTypeface.FromData(data) ?? throw new NotSupportedException("Skia could not read the selected CFF face.");
        try
        {
            if (face.GlyphCount != GlyphCount || face.UnitsPerEm != UnitsPerEm)
                throw new InvalidDataException("CFF charstring count or units per em differs from sfnt metrics.");
            using var stream = face.OpenStream(out var index);
            if (index != 0 || stream is null || !stream.HasLength || stream.Length != original.Length)
                throw new NotSupportedException("Skia did not expose the original standalone CFF face.");
            var buffer = new byte[8192]; var offset = 0;
            while (offset < original.Length)
            {
                var count = stream.Read(buffer, Math.Min(buffer.Length, original.Length - offset));
                if (count <= 0) throw new InvalidDataException("Incomplete Skia CFF source stream.");
                for (var i = 0; i < count; i++)
                    if (buffer[i] != original[offset + i])
                        throw new InvalidDataException("Skia CFF source stream differs from the selected face.");
                offset += count;
            }
        }
        catch { face.Dispose(); throw; }
        using var sha = SHA256.Create();
        SourceSha256 = BitConverter.ToString(sha.ComputeHash(original)).Replace("-", "").ToLowerInvariant();
    }

    internal string SourceSha256 { get; }
    internal int UnitsPerEm { get; }
    internal int GlyphCount { get; }

    internal CffGlyphOutline Read(int glyphId, CancellationToken cancellationToken = default)
    {
        lock (gate)
        {
            if (disposed) throw new ObjectDisposedException(nameof(CffGlyphOutlines));
            cancellationToken.ThrowIfCancellationRequested();
            if (glyphId <= 0 || glyphId >= GlyphCount)
                throw new NotSupportedException("The requested original CFF GID is outside the admitted font.");
            using var font = new SKFont(face, UnitsPerEm)
            {
                Hinting = SKFontHinting.None, Edging = SKFontEdging.Alias,
                Subpixel = false, LinearMetrics = true, Embolden = false,
                ScaleX = 1, SkewX = 0
            };
            using var path = font.GetGlyphPath((ushort)glyphId);
            cancellationToken.ThrowIfCancellationRequested();
            // Skia uses null both for a genuine empty glyph (for example,
            // space) and for an unavailable native outline. Until the CFF
            // CharString itself is checked, neither can be returned as a
            // proven empty original-GID outline.
            if (path is null || path.VerbCount == 0)
                throw new NotSupportedException("CFF GID has no verifiable native outline.");
            var commands = new List<CffPathCommand>();
            if (path.VerbCount > MaximumCommands)
                throw new InvalidDataException("CFF glyph command budget exceeded.");
            using var iterator = path.CreateRawIterator();
            var points = new SKPoint[4];
            for (var verb = iterator.Next(points); verb != SKPathVerb.Done; verb = iterator.Next(points))
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (commands.Count >= MaximumCommands)
                    throw new InvalidDataException("CFF glyph command budget exceeded.");
                CffPathCommand command;
                switch (verb)
                {
                    case SKPathVerb.Move:
                        command = new(CffPathVerb.Move, X(points[0]), Y(points[0])); break;
                    case SKPathVerb.Line:
                        command = new(CffPathVerb.Line, X(points[1]), Y(points[1])); break;
                    case SKPathVerb.Cubic:
                        command = new(CffPathVerb.Cubic, X(points[3]), Y(points[3]),
                            X(points[1]), Y(points[1]), X(points[2]), Y(points[2])); break;
                    case SKPathVerb.Close:
                        command = new(CffPathVerb.Close, 0, 0); break;
                    default:
                        throw new NotSupportedException("The CFF outline backend returned a non-cubic path segment.");
                }
                commands.Add(command);
            }
            if (!commands.Any(command => command.Verb is CffPathVerb.Line or CffPathVerb.Cubic))
                throw new NotSupportedException("CFF GID has no drawable native outline.");
            cancellationToken.ThrowIfCancellationRequested();
            var advance = OpenTypeFace.U16(hmtx, Math.Min(glyphId, metricCount - 1) * 4);
            return new CffGlyphOutline(SourceSha256, glyphId, UnitsPerEm, advance,
                new ReadOnlyCollection<CffPathCommand>(commands));
        }
    }

    private static double X(SKPoint point) => Bounded(point.X);
    private static double Y(SKPoint point) => -Bounded(point.Y);
    private static double Bounded(float value)
    {
        if (float.IsNaN(value) || float.IsInfinity(value) || Math.Abs(value) > 1_000_000)
            throw new InvalidDataException("CFF outline exceeds bounded design coordinates.");
        return value;
    }

    public void Dispose()
    {
        lock (gate)
        {
            if (disposed) return;
            disposed = true; face.Dispose();
        }
    }
}

internal sealed class CffGlyphOutline(string sourceSha256, int glyphId, int unitsPerEm,
    int advanceWidth, IReadOnlyList<CffPathCommand> commands)
{
    internal string SourceSha256 { get; } = sourceSha256;
    internal int GlyphId { get; } = glyphId;
    internal int UnitsPerEm { get; } = unitsPerEm;
    internal int AdvanceWidth { get; } = advanceWidth;
    internal IReadOnlyList<CffPathCommand> Commands { get; } = commands;
}

internal enum CffPathVerb { Move, Line, Cubic, Close }

internal readonly struct CffPathCommand(CffPathVerb verb, double x, double y,
    double controlX1 = 0, double controlY1 = 0, double controlX2 = 0, double controlY2 = 0)
{
    internal CffPathVerb Verb { get; } = verb;
    internal double X { get; } = x;
    internal double Y { get; } = y;
    internal double ControlX1 { get; } = controlX1;
    internal double ControlY1 { get; } = controlY1;
    internal double ControlX2 { get; } = controlX2;
    internal double ControlY2 { get; } = controlY2;
}
