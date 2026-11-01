using OfdFillRule = Ofdrw.Net.Layout.Graphics.OfdFillRule;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Xml;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Graphics;
using SkiaSharp;

namespace Ofdrw.Net.Graphics.SkiaSharp;

/// <summary>An immutable snapshot supplied explicitly by a cooperating Skia producer. This is not an intercepted SKCanvas call.</summary>
public sealed partial class SkiaDrawEvent
{
    private readonly OfdGraphicsPath? _path;
    private readonly OfdPen? _pen;
    private readonly OfdBrush? _brush;
    private readonly OfdMatrix _matrix;
    private readonly string? _text, _fontId;
    private readonly double[]? _advances;
    private readonly byte[]? _fontHash;
    private readonly double _fontSize, _x, _y;
    private readonly bool _bold, _italic;
    internal int PathCommands { get; }
    internal int TextCharacters => _text?.Length ?? 0;

    private SkiaDrawEvent(OfdGraphicsPath path, SKPaint paint, SKMatrix? matrix, bool joinsMatter)
    {
        (_pen, _brush) = Paint(paint, joinsMatter);
        _matrix = Matrix(matrix); _path = path; PathCommands = path.CommandCount;
    }
    private SkiaDrawEvent(string text, SKPoint baseline, SKFont font, SKPaint paint, string fontId,
        IReadOnlyList<float> advances, SKMatrix? matrix, int maxTextCharacters, int maxFontBytes,
        Func<FontIdentity>? identityFactory = null, CancellationToken token = default)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        if (text.Length > maxTextCharacters) throw new InvalidOperationException("Text snapshot budget exceeded.");
        if (text.Length == 0 || text.IndexOfAny(new[] { '\r', '\n', '\t' }) >= 0)
            throw new ArgumentException("Original Unicode must be a nonempty single baseline run.", nameof(text));
        XmlConvert.VerifyXmlChars(text);
        if (font is null) throw new ArgumentNullException(nameof(font));
        if (font.Handle == IntPtr.Zero) throw new ObjectDisposedException(nameof(font));
        if (paint is null) throw new ArgumentNullException(nameof(paint));
        if (paint.Style != SKPaintStyle.Fill) throw new NotSupportedException("Text requires a fill paint.");
        (_, _brush) = Paint(paint, false);
        var face = font.Typeface;
        if (face is null || face.Handle == IntPtr.Zero) throw new ArgumentException("An explicit live typeface is required.", nameof(font));
        var readTrace = identityFactory is null ? FontReadTrace.Create(null) : null;
        if (readTrace is not null) { readTrace.Bind(face); ValidateFixedFace(face, readTrace); }
        if (font.ScaleX != 1 || font.SkewX != 0 || font.Embolden)
            throw new NotSupportedException("Synthetic SKFont scale, skew or emboldening is not supported.");
        Positive(font.Size); Point(baseline);
        if (string.IsNullOrWhiteSpace(fontId)) throw new ArgumentException("A destination font resource ID is required.", nameof(fontId));
        if (advances is null) throw new ArgumentNullException(nameof(advances));
        // Use the same .NET text-element contract as OfdGraphics; never recover text from glyph IDs.
        var count = StringInfo.ParseCombiningCharacters(text).Length;
        if (advances.Count != count - 1) throw new ArgumentException("Advances must match grapheme count minus one.", nameof(advances));
        _advances = new double[advances.Count];
        for (var i = 0; i < advances.Count; i++) { Finite(advances[i]); _advances[i] = advances[i]; }
        _matrix = Matrix(matrix); _text = text; _fontId = fontId; _fontSize = font.Size;
        _x = baseline.X; _y = baseline.Y;
        // Public Text authenticates every caller-supplied face. Only the owned
        // producer binding can reuse a privately authenticated identity.
        var identity = identityFactory is null ? ReadIdentity(face, maxFontBytes, null, token, readTrace!) : identityFactory();
        _bold = identity.Bold; _italic = identity.Italic;
        _fontHash = (byte[])identity.Hash.Clone();
    }

    /// <summary>Records a nonzero line. A line has no joins, so any positive miter limit has identical geometry.</summary>
    public static SkiaDrawEvent Line(SKPoint start, SKPoint end, SKPaint paint, SKMatrix? matrix = null)
    {
        Point(start); Point(end);
        if (start == end) throw new ArgumentException("A line must have distinct endpoints.");
        if (paint is null) throw new ArgumentNullException(nameof(paint));
        if (paint.Style != SKPaintStyle.Stroke) throw new NotSupportedException("A line requires a stroke paint.");
        return new SkiaDrawEvent(new OfdGraphicsPath().MoveTo(start.X, start.Y).LineTo(end.X, end.Y), paint, matrix, false);
    }
    /// <summary>Records a positive rectangle. Miter limits at least sqrt(2), including Skia's default 4, retain its right-angle joins.</summary>
    public static SkiaDrawEvent Rectangle(SKRect rectangle, SKPaint paint, SKMatrix? matrix = null)
    {
        Finite(rectangle.Left); Finite(rectangle.Top); Positive(rectangle.Width); Positive(rectangle.Height);
        if (paint is null) throw new ArgumentNullException(nameof(paint));
        if (paint.Style != SKPaintStyle.Fill && paint.StrokeMiter < Math.Sqrt(2))
            throw new NotSupportedException("Rectangle miter limit would bevel a right-angle corner.");
        return new SkiaDrawEvent(new OfdGraphicsPath().AddRectangle(rectangle.Left, rectangle.Top, rectangle.Width, rectangle.Height), paint, matrix, false);
    }
    /// <summary>Records M/L/Q/cubic/close path data under a bounded command budget. Conics and inverse fill explicitly fail.</summary>
    public static SkiaDrawEvent Path(SKPath path, SKPaint paint, SKMatrix? matrix = null, int maxPathCommands = 100_000)
    {
        if (path is null) throw new ArgumentNullException(nameof(path));
        if (path.Handle == IntPtr.Zero) throw new ObjectDisposedException(nameof(path));
        if (maxPathCommands <= 0) throw new ArgumentOutOfRangeException(nameof(maxPathCommands));
        if (path.FillType is not (SKPathFillType.Winding or SKPathFillType.EvenOdd)) throw new NotSupportedException("Inverse fill is not supported.");
        // Reject paint effects before allocating the geometry snapshot.
        Paint(paint, true); Matrix(matrix);
        var result = new OfdGraphicsPath(path.FillType == SKPathFillType.EvenOdd ? OfdFillRule.EvenOdd : OfdFillRule.NonZero, maxPathCommands);
        using var iterator = path.CreateRawIterator(); var points = new SKPoint[4];
        var drawable = false; var closed = false; SKPoint figureStart = default;
        for (var verb = iterator.Next(points); verb != SKPathVerb.Done; verb = iterator.Next(points))
        {
            if (result.CommandCount >= maxPathCommands) throw new InvalidOperationException("Path snapshot budget exceeded.");
            if (closed && verb is SKPathVerb.Line or SKPathVerb.Quad or SKPathVerb.Cubic)
            { result.MoveTo(figureStart.X, figureStart.Y); closed = false; }
            switch (verb)
            {
                case SKPathVerb.Move: Point(points[0]); result.MoveTo(points[0].X, points[0].Y); figureStart = points[0]; closed = false; break;
                case SKPathVerb.Line: Point(points[1]); result.LineTo(points[1].X, points[1].Y); drawable = true; break;
                case SKPathVerb.Quad: Point(points[1]); Point(points[2]); result.QuadraticTo(points[1].X, points[1].Y, points[2].X, points[2].Y); drawable = true; break;
                case SKPathVerb.Cubic: Point(points[1]); Point(points[2]); Point(points[3]); result.BezierTo(points[1].X, points[1].Y, points[2].X, points[2].Y, points[3].X, points[3].Y); drawable = true; break;
                case SKPathVerb.Close: result.Close(); closed = true; break;
                default: throw new NotSupportedException("Unsupported Skia path verb: " + verb);
            }
        }
        if (!drawable) throw new ArgumentException("A path must contain drawable segments.", nameof(path));
        return new SkiaDrawEvent(result, paint, matrix, true);
    }
    /// <summary>Records original Unicode and producer-supplied user-space advances, binding an exact font payload to a destination ID.</summary>
    /// <remarks>No shaping or font resolution is performed. SKFont synthetic styles, TTC and unavailable original font streams fail. Caller may dispose all native inputs after return.</remarks>
    public static SkiaDrawEvent Text(string text, SKPoint baseline, SKFont font, SKPaint paint, string fontResourceId,
        IReadOnlyList<float> advances, SKMatrix? matrix = null, int maxTextCharacters = 1_000_000, int maxFontBytes = 64_000_000)
    {
        if (maxTextCharacters <= 0 || maxFontBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxTextCharacters));
        return new SkiaDrawEvent(text, baseline, font, paint, fontResourceId, advances, matrix, maxTextCharacters, maxFontBytes);
    }
    /// <summary>Fails explicitly when a producer reaches a clip, layer, image, shaped blob or other unsupported operation. No event is emitted.</summary>
    public static SkiaDrawEvent Unsupported(string operation) => throw new NotSupportedException("Unsupported producer operation: " + operation);

    internal void Emit(OfdGraphics graphics, OfdDocumentPackage package, double units, Dictionary<string, byte[]> fontHashes,
        int maxFontBytes, System.Threading.CancellationToken token)
    {
        graphics.SetTransform(OfdMatrix.Scale(units, units).Multiply(_matrix));
        if (_path is not null) { graphics.DrawPath(_path, _pen, _brush, token); return; }
        var matches = package.Fonts.FindAll(value => string.Equals(value.Id, _fontId, StringComparison.Ordinal));
        if (matches.Count != 1 || string.IsNullOrWhiteSpace(matches[0].FontName)) throw new ArgumentException("Font ID must match exactly one named destination resource.");
        var resource = matches[0];
        if (resource.Data is null || resource.Data.Length == 0 || resource.Data.Length > maxFontBytes)
            throw new ArgumentException("Bounded embedded font payload is required.");
        if (resource.Bold != _bold || resource.Italic != _italic) throw new ArgumentException("Font face style does not match the destination resource.");
        if (!fontHashes.TryGetValue(_fontId!, out var actual))
        { using var hash = SHA256.Create(); actual = hash.ComputeHash(resource.Data); fontHashes.Add(_fontId!, actual); }
        for (var i = 0; i < actual.Length; i++) if (actual[i] != _fontHash![i]) throw new ArgumentException("SKFont face payload does not match the destination resource.");
        graphics.DrawString(_text!, new OfdFont(resource.FontName, _fontSize, weight: 400, italic: false, resourceId: _fontId),
            _brush!, _x, _y, _advances, token);
    }
    private static (OfdPen?, OfdBrush?) Paint(SKPaint paint, bool joinsMatter)
    {
        if (paint is null) throw new ArgumentNullException(nameof(paint));
        if (paint.Handle == IntPtr.Zero) throw new ObjectDisposedException(nameof(paint));
        // BlendMode falls back to SrcOver for runtime/arithmetic blenders.
        // Only the SDK's canonical SrcOver mode blender (or default null) is
        // proven representable; never infer custom blender semantics from the enum.
        var blender = paint.Blender;
        if (blender is not null && !ReferenceEquals(blender, SKBlender.CreateBlendMode(SKBlendMode.SrcOver)))
            throw new NotSupportedException("Custom Skia blenders are not supported.");
        if (paint.Shader is not null || paint.ColorFilter is not null || paint.ImageFilter is not null || paint.MaskFilter is not null || paint.PathEffect is not null || paint.BlendMode != SKBlendMode.SrcOver)
            throw new NotSupportedException("Only solid SrcOver paint without shaders, filters or path effects is supported.");
        if (paint.Style is not (SKPaintStyle.Fill or SKPaintStyle.Stroke or SKPaintStyle.StrokeAndFill)) throw new NotSupportedException("Unsupported paint style.");
        var color = paint.Color; var precise = paint.ColorF;
        Finite(precise.Red); Finite(precise.Green); Finite(precise.Blue); Finite(precise.Alpha);
        if (Math.Abs(precise.Red - color.Red / 255f) > 0.0000001f || Math.Abs(precise.Green - color.Green / 255f) > 0.0000001f ||
            Math.Abs(precise.Blue - color.Blue / 255f) > 0.0000001f || Math.Abs(precise.Alpha - color.Alpha / 255f) > 0.0000001f)
            throw new NotSupportedException("Only 8-bit sRGB paint colors are supported; high precision/HDR colors require another profile.");
        var mapped = new OfdColor(color.Red, color.Green, color.Blue, color.Alpha);
        OfdPen? pen = null; OfdBrush? brush = null;
        if (paint.Style != SKPaintStyle.Fill)
        {
            Positive(paint.StrokeWidth); Positive(paint.StrokeMiter);
            if (paint.StrokeCap != SKStrokeCap.Butt || paint.StrokeJoin != SKStrokeJoin.Miter || (joinsMatter && paint.StrokeMiter != 10))
                throw new NotSupportedException("Strokes require butt caps and miter joins; arbitrary paths require miter limit 10.");
            pen = new OfdPen(mapped, paint.StrokeWidth);
        }
        if (paint.Style != SKPaintStyle.Stroke) brush = new OfdBrush(mapped);
        return (pen, brush);
    }
    private static OfdMatrix Matrix(SKMatrix? matrix)
    {
        var value = matrix ?? SKMatrix.Identity;
        Finite(value.Persp0); Finite(value.Persp1); Finite(value.Persp2);
        if (value.Persp0 != 0 || value.Persp1 != 0 || value.Persp2 != 1) throw new NotSupportedException("Perspective is not supported.");
        var result = new OfdMatrix(value.ScaleX, value.SkewY, value.SkewX, value.ScaleY, value.TransX, value.TransY);
        // Reuse 04's exact singular validation rather than duplicating its numeric contract.
        var package = new OfdDocumentPackage(); var page = new OfdPage { WidthMillimeters = 1, HeightMillimeters = 1 }; package.Pages.Add(page);
        new OfdGraphics(package, page).SetTransform(result); return result;
    }
    private static ushort FontWord(SKTypeface face, uint tag, int offset, FontReadTrace trace, string phase, NativeFontTableAccess? access = null)
    {
        var size = NativeTableSize(face, tag, phase, trace, access);
        if (size < offset + 2)
            throw trace.Failure("native-style-table-missing-or-truncated", "Font style table is missing or truncated.");
        var buffer = new byte[1]; var pinned = GCHandle.Alloc(buffer, GCHandleType.Pinned);
        try
        {
            // SDK TryGetTableData only promises a nonzero read. A one-byte
            // request makes success unambiguous and never allocates a whole table.
            if (!ReadByte(offset))
                throw trace.Failure("native-style-byte-unavailable", "Font style flags could not be read.");
            var high = buffer[0];
            if (!ReadByte(offset + 1))
                throw trace.Failure("native-style-byte-unavailable", "Font style flags could not be read.");
            return (ushort)((high << 8) | buffer[0]);
            bool ReadByte(int position)
            {
                bool success;
                var started = FontReadTrace.StartTimestamp();
                try
                {
                    success = access?.Byte is { } read ? read(face, tag, position, pinned.AddrOfPinnedObject(), phase) :
                        face.TryGetTableData(tag, position, 1, pinned.AddrOfPinnedObject());
                }
                catch (Exception error)
                {
                    trace.Record(phase, "byte-threw", tag, offset: position, startedTimestamp: started);
                    trace.Attach(error, "native-table-call-threw"); throw;
                }
                trace.Record(phase, "byte", tag, offset: position, size: 1, success: success, value: success ? buffer[0] : null, startedTimestamp: started);
                return success;
            }
        }
        finally { pinned.Free(); }
    }
    private static void Point(SKPoint value) { Finite(value.X); Finite(value.Y); }
    private static void Finite(double value) { if (double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentException("Values must be finite."); }
    private static void Positive(double value) { Finite(value); if (value <= 0) throw new ArgumentOutOfRangeException(nameof(value), "Value must be positive."); }
}
