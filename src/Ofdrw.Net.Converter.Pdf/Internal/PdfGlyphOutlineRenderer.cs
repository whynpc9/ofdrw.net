using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using PdfSharpCore.Drawing;
using SixLabors.Fonts;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>
/// Strokes a fallback font's outlines without emitting a second PDF text layer.
/// </summary>
internal sealed class PdfGlyphOutlineRenderer : IGlyphRenderer
{
    private const int CurveSegments = 16;
    private readonly XGraphics _graphics;
    private readonly XPen? _pen;
    private readonly XColor _color;
    private readonly bool _fill;
    private readonly double _emboldenHalfWidth;
    private readonly List<IReadOnlyList<(double X, double Y)>>? _contours;
    private List<(double X, double Y)>? _contour;
    private XGraphicsPath _path = new();
    private Vector2 _position;

    internal PdfGlyphOutlineRenderer(XGraphics graphics, XColor color, double strokeWidth, bool fill = true)
    {
        _graphics = graphics;
        _color = color;
        _pen = new XPen(color, strokeWidth);
        _fill = fill;
    }

    /// <summary>
    /// Fills only the area each glyph gains from a horizontal dilation by
    /// <paramref name="halfWidth"/> per side (see <see cref="HorizontalEmboldenOverlay"/>).
    /// The caller paints the original glyph separately. No stroke is painted,
    /// so viewers cannot widen the emphasis to a one-device-pixel hairline.
    /// </summary>
    internal PdfGlyphOutlineRenderer(XGraphics graphics, XColor color, double halfWidth)
    {
        _graphics = graphics;
        _color = color;
        _fill = true;
        _emboldenHalfWidth = halfWidth;
        _contours = new List<IReadOnlyList<(double X, double Y)>>();
    }

    public void BeginText(in FontRectangle bounds) { }
    public void EndText() { }
    public bool BeginGlyph(in FontRectangle bounds, in GlyphRendererParameters parameters)
    {
        _path = new XGraphicsPath();
        _contours?.Clear();
        _contour = null;
        return true;
    }

    public TextDecorations EnabledDecorations() => TextDecorations.None;
    public void SetDecoration(TextDecorations textDecorations, Vector2 start, Vector2 end, float thickness) { }

    public void EndGlyph()
    {
        if (_contours is not null)
        {
            var quads = HorizontalEmboldenOverlay.Build(_contours, _emboldenHalfWidth);
            if (quads.Count == 0) return;
            // One nonzero path: adjacent trapezoids share edges without seams.
            var overlay = new XGraphicsPath { FillMode = XFillMode.Winding };
            foreach (var quad in quads)
                overlay.AddPolygon(quad.Select(point => new XPoint(point.X, point.Y)).ToArray());
            _graphics.DrawPath(new XSolidBrush(_color), overlay);
            return;
        }
        if (_fill) _graphics.DrawPath(_pen!, new XSolidBrush(_color), _path);
        else _graphics.DrawPath(_pen!, _path);
    }

    public void BeginFigure() => _path.StartFigure();

    public void EndFigure()
    {
        _path.CloseFigure();
        _contour = null;
    }

    public void MoveTo(Vector2 point)
    {
        _position = point;
        if (_contours is null) return;
        _contour = new List<(double X, double Y)> { (point.X, point.Y) };
        _contours.Add(_contour);
    }

    public void LineTo(Vector2 point)
    {
        _path.AddLine(_position.X, _position.Y, point.X, point.Y);
        AddContourPoint(point);
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
        if (_contours is not null)
        {
            for (var step = 1; step <= CurveSegments; step++)
            {
                var t = step / (float)CurveSegments;
                var u = 1 - t;
                AddContourPoint(u * u * u * _position + 3 * u * u * t * control1 + 3 * u * t * t * control2 + t * t * t * point);
            }
        }
        _position = point;
    }

    private void AddContourPoint(Vector2 point)
    {
        if (_contours is null) return;
        if (_contour is null)
        {
            _contour = new List<(double X, double Y)> { (_position.X, _position.Y) };
            _contours.Add(_contour);
        }
        _contour.Add((point.X, point.Y));
    }
}
