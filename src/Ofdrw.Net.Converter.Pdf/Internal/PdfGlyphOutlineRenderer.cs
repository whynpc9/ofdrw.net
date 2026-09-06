using System.Numerics;
using PdfSharpCore.Drawing;
using SixLabors.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Strokes a fallback font's outlines without emitting a second PDF text layer.
/// </summary>
internal sealed class PdfGlyphOutlineRenderer : IGlyphRenderer
{
    private readonly XGraphics _graphics;
    private readonly XPen _pen;
    private XGraphicsPath _path = new();
    private Vector2 _position;

    internal PdfGlyphOutlineRenderer(XGraphics graphics, XColor color, double strokeWidth)
    {
        _graphics = graphics;
        _pen = new XPen(color, strokeWidth);
    }

    public void BeginText(FontRectangle bounds) { }
    public void EndText() { }
    public bool BeginGlyph(FontRectangle bounds, GlyphRendererParameters parameters)
    {
        _path = new XGraphicsPath();
        return true;
    }

    public void EndGlyph() => _graphics.DrawPath(_pen, new XSolidBrush(_pen.Color), _path);
    public void BeginFigure() => _path.StartFigure();
    public void EndFigure() => _path.CloseFigure();
    public void MoveTo(Vector2 point) => _position = point;

    public void LineTo(Vector2 point)
    {
        _path.AddLine(_position.X, _position.Y, point.X, point.Y);
        _position = point;
    }

    public void QuadraticBezierTo(Vector2 control, Vector2 point)
    {
        CubicBezierTo(_position + (control - _position) * (2f / 3f),
            point + (control - point) * (2f / 3f), point);
    }

    public void CubicBezierTo(Vector2 control1, Vector2 control2, Vector2 point)
    {
        _path.AddBezier(_position.X, _position.Y, control1.X, control1.Y,
            control2.X, control2.Y, point.X, point.Y);
        _position = point;
    }
}
