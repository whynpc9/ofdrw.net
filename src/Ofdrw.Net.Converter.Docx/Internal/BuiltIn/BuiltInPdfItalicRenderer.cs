using System;
using System.Collections.Generic;
using System.IO;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Core.Fonts;
using PdfSharpCore.Drawing;
using PdfSharpCore.Fonts;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

// PdfSharpCore 1.3.67 drops FontResolverInfo.StyleSimulations in XGlyphTypeface.
// Intercept the exact MigraDoc drawing call, before PDF character encoding. This
// keeps the original layout, glyph selection and single semantic text operation.
internal sealed class BuiltInPdfItalicRenderer : IXGraphicsRenderer
{
    private const double Skew = 0.342020143325669; // PDFsharp's sin(20 degrees).
    private readonly XGraphics _target;
    private readonly IReadOnlyDictionary<string, bool> _italicFaces;
    private readonly List<(XGraphicsState Source, XGraphicsState Target)> _states = new();

    internal BuiltInPdfItalicRenderer(XGraphics target, IReadOnlyDictionary<string, bool> italicFaces)
    {
        _target = target;
        _italicFaces = italicFaces;
    }

    public void DrawString(string text, XFont font, XBrush brush, XRect rectangle, XStringFormat format)
    {
        if (!font.Italic)
        {
            _target.DrawString(text, font, brush, rectangle, format);
            return;
        }
        if (!_italicFaces.TryGetValue(font.Name, out var actualItalic))
            throw new InvalidDataException("BuiltIn cannot prove the physical face of an italic text operation.");
        if (actualItalic)
        {
            _target.DrawString(text, font, brush, rectangle, format);
            return;
        }
        // MigraDoc supplies baseline coordinates for text, fields and tab leaders.
        // Reject unsupported alignment rather than silently shifting an anchor.
        if (format.LineAlignment != XLineAlignment.BaseLine || rectangle.Height != 0)
            throw new InvalidDataException("BuiltIn italic simulation requires an explicit text baseline.");
        var state = _target.Save();
        try
        {
            _target.MultiplyTransform(new XMatrix(1, 0, -Skew, 1, Skew * rectangle.Y, 0));
            _target.DrawString(text, font, brush, rectangle, format);
        }
        finally { _target.Restore(state); }
    }

    public void Save(XGraphicsState state) => _states.Add((state, _target.Save()));
    public void Restore(XGraphicsState state)
    {
        var index = _states.FindLastIndex(pair => ReferenceEquals(pair.Source, state));
        if (index < 0) throw new InvalidDataException("Unbalanced BuiltIn PDF graphics state.");
        _target.Restore(_states[index].Target);
        _states.RemoveRange(index, _states.Count - index);
    }

    // MigraDoc's supported DOCX model uses Save/Restore and intersect clips.
    // Refuse future unsupported drawing operations before the candidate commits.
    public void BeginContainer(XGraphicsContainer container, XRect destination, XRect source, XGraphicsUnit unit) =>
        throw new NotSupportedException("BuiltIn PDF graphics containers are not supported.");
    public void EndContainer(XGraphicsContainer container) =>
        throw new NotSupportedException("BuiltIn PDF graphics containers are not supported.");
    public void ResetClip() => throw new NotSupportedException("BuiltIn PDF clip reset is not supported.");
    public void SetClip(XGraphicsPath path, XCombineMode mode)
    {
        if (mode != XCombineMode.Intersect) throw new NotSupportedException("BuiltIn PDF requires intersect clipping.");
        _target.IntersectClip(path);
    }
    public void Close() { } // The caller owns the target graphics context.
    public void AddTransform(XMatrix matrix, XMatrixOrder order) => _target.MultiplyTransform(matrix, order);
    public void WriteComment(string comment) => _target.WriteComment(comment);
    public void DrawLine(XPen pen, double x1, double y1, double x2, double y2) => _target.DrawLine(pen, x1, y1, x2, y2);
    public void DrawLines(XPen pen, XPoint[] points) => _target.DrawLines(pen, points);
    public void DrawBezier(XPen pen, double x1, double y1, double x2, double y2, double x3, double y3, double x4, double y4) =>
        _target.DrawBezier(pen, x1, y1, x2, y2, x3, y3, x4, y4);
    public void DrawBeziers(XPen pen, XPoint[] points) => _target.DrawBeziers(pen, points);
    public void DrawCurve(XPen pen, XPoint[] points, double tension) => _target.DrawCurve(pen, points, tension);
    public void DrawArc(XPen pen, double x, double y, double width, double height, double start, double sweep) =>
        _target.DrawArc(pen, x, y, width, height, start, sweep);
    public void DrawRectangle(XPen pen, XBrush brush, double x, double y, double width, double height) =>
        _target.DrawRectangle(pen, brush, x, y, width, height);
    public void DrawRectangles(XPen pen, XBrush brush, XRect[] rectangles) => _target.DrawRectangles(pen, brush, rectangles);
    public void DrawRoundedRectangle(XPen pen, XBrush brush, double x, double y, double width, double height, double ellipseWidth, double ellipseHeight) =>
        _target.DrawRoundedRectangle(pen, brush, x, y, width, height, ellipseWidth, ellipseHeight);
    public void DrawEllipse(XPen pen, XBrush brush, double x, double y, double width, double height) =>
        _target.DrawEllipse(pen, brush, x, y, width, height);
    public void DrawPolygon(XPen pen, XBrush brush, XPoint[] points, XFillMode mode) => _target.DrawPolygon(pen, brush, points, mode);
    public void DrawPie(XPen pen, XBrush brush, double x, double y, double width, double height, double start, double sweep) =>
        _target.DrawPie(pen, brush, x, y, width, height, start, sweep);
    public void DrawClosedCurve(XPen pen, XBrush brush, XPoint[] points, double tension, XFillMode mode) =>
        _target.DrawClosedCurve(pen, brush, points, mode, tension);
    public void DrawPath(XPen pen, XBrush brush, XGraphicsPath path) => _target.DrawPath(pen, brush, path);
    public void DrawImage(XImage image, double x, double y, double width, double height) => _target.DrawImage(image, x, y, width, height);
    public void DrawImage(XImage image, XRect destination, XRect source, XGraphicsUnit unit) => _target.DrawImage(image, destination, source, unit);
}

// A conversion-local snapshot of each selected italic face. The physical style
// comes from the original sfnt, never from the host's potentially false flags.
internal sealed class BuiltInPdfItalicFaces
{
    private readonly IFontResolver _resolver;
    private readonly Func<string, byte[]> _readFont;
    private readonly Dictionary<(string Family, bool Bold), string> _families = new();
    private readonly Dictionary<string, bool> _actualItalic = new(StringComparer.Ordinal);

    internal BuiltInPdfItalicFaces(IFontResolver resolver, Func<string, byte[]>? readFont = null)
    {
        _resolver = resolver;
        _readFont = readFont ?? PdfFontRegistry.GetOriginalFont;
    }
    internal IReadOnlyDictionary<string, bool> ActualItalic => _actualItalic;

    internal string Resolve(string family, bool bold)
    {
        var key = (family, bold);
        if (_families.TryGetValue(key, out var cached)) return cached;
        var face = _resolver.ResolveTypeface(family, bold, true);
        if (face is null) throw new InvalidDataException("BuiltIn could not resolve the requested italic font.");
        var bytes = _readFont(face.FaceName);
        // GetOriginalFont delegates to the installed host for unregistered faces.
        var actual = new OpenTypeFace(bytes).Style;
        var alias = PdfFontRegistry.RegisterFontFace(bytes, bold, true);
        var font = new XFont(alias, 10, bold ? XFontStyle.BoldItalic : XFontStyle.Italic);
        _actualItalic[font.Name] = actual.Italic;
        _families.Add(key, alias);
        return alias;
    }
}
