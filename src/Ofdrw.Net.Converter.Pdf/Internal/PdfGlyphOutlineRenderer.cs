using System;
using System.Numerics;
using System.Threading;
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
    private readonly bool _fill;
    private readonly bool _geometricStroke;
    private readonly CancellationToken _cancellationToken;
    private PdfGlyphStrokeGeometry? _strokeGeometry;
    private XGraphicsPath _path = new();
    private Vector2 _position;

    internal PdfGlyphOutlineRenderer(XGraphics graphics, XColor color, double strokeWidth, bool fill = true,
        bool geometricStroke = false, CancellationToken cancellationToken = default)
    {
        _graphics = graphics;
        _pen = new XPen(color, strokeWidth);
        _fill = fill;
        _geometricStroke = geometricStroke;
        _cancellationToken = cancellationToken;
    }

    public void BeginText(in FontRectangle bounds) { }
    public void EndText() { }
    public bool BeginGlyph(in FontRectangle bounds, in GlyphRendererParameters parameters)
    {
        _path = new XGraphicsPath { FillMode = _geometricStroke ? XFillMode.Winding : XFillMode.Alternate };
        _strokeGeometry = _geometricStroke ? new PdfGlyphStrokeGeometry(StrokeFlatness()) : null;
        return true;
    }

    private double StrokeFlatness()
    {
        var matrix = _graphics.Transform;
        var scale = Math.Max(Math.Max(Math.Abs(matrix.M11), Math.Abs(matrix.M12)),
            Math.Max(Math.Abs(matrix.M21), Math.Abs(matrix.M22)));
        if (double.IsNaN(scale) || double.IsInfinity(scale))
            throw new NotSupportedException("Glyph stroke transform must remain finite.");
        if (scale == 0) return .001;
        // Largest singular value of the current linear transform. Normalize
        // first to avoid overflow; translation does not amplify curve error.
        var a = matrix.M11 / scale; var b = matrix.M12 / scale;
        var c = matrix.M21 / scale; var d = matrix.M22 / scale;
        var trace = a * a + b * b + c * c + d * d;
        var determinant = a * d - b * c;
        var largest = (trace + Math.Sqrt(Math.Max(0, trace * trace - 4 * determinant * determinant))) * .5;
        var amplification = scale * Math.Sqrt(largest);
        return .001 / Math.Max(1, amplification);
    }

    public TextDecorations EnabledDecorations() => TextDecorations.None;
    public void SetDecoration(TextDecorations textDecorations, Vector2 start, Vector2 end, float thickness) { }

    public void EndGlyph()
    {
        if (_strokeGeometry is not null)
        {
            var brush = new XSolidBrush(_pen.Color);
            // Overlapping nonzero subpaths can amplify edge coverage even in
            // one fill. Admitted rectilinear glyphs use only the union boundary.
            if (_strokeGeometry.TryBuildFillUnion(_pen.Width, _cancellationToken, out var combined))
            {
                _graphics.DrawPath(brush, combined!);
                return;
            }
            _graphics.DrawPath(brush, _path);
            _graphics.DrawPath(brush, _strokeGeometry.Build(_pen.Width));
            return;
        }
        if (_fill) _graphics.DrawPath(_pen, new XSolidBrush(_pen.Color), _path);
        else _graphics.DrawPath(_pen, _path);
    }
    public void BeginFigure() { _path.StartFigure(); _strokeGeometry?.BeginFigure(); }
    public void EndFigure() { _path.CloseFigure(); _strokeGeometry?.EndFigure(); }
    public void MoveTo(Vector2 point) { _position = point; _strokeGeometry?.MoveTo(point); }

    public void LineTo(Vector2 point)
    {
        _path.AddLine(_position.X, _position.Y, point.X, point.Y);
        _strokeGeometry?.LineTo(point);
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
        _strokeGeometry?.CubicBezierTo(_position, control1, control2, point);
        _position = point;
    }
}
