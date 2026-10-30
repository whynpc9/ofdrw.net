using System;
using System.Collections.Generic;
using System.Numerics;
using PdfSharpCore.Drawing;

namespace Ofdrw.Net.Converter.Pdf.Internal;

/// <summary>Closed glyph strokes as filled regions, before the existing graphics transform.</summary>
internal sealed class PdfGlyphStrokeGeometry
{
    private const int MaximumPoints = 65536;
    private readonly double _flatnessPoints;
    private readonly List<List<Point>> _contours = new();
    private List<Point>? _current;
    private int _points;

    internal PdfGlyphStrokeGeometry(double flatnessPoints)
    {
        if (!(flatnessPoints > 0) || double.IsInfinity(flatnessPoints))
            throw new NotSupportedException("Glyph stroke transform exceeds its finite precision budget.");
        _flatnessPoints = flatnessPoints;
    }

    internal void BeginFigure()
    {
        _current = new List<Point>();
        _contours.Add(_current);
    }

    internal void MoveTo(Vector2 point)
    {
        if (_current is null) BeginFigure();
        Add(new Point(point.X, point.Y));
    }

    internal void LineTo(Vector2 point) => Add(new Point(point.X, point.Y));

    internal void CubicBezierTo(Vector2 from, Vector2 control1, Vector2 control2, Vector2 to) =>
        Flatten(new Point(from.X, from.Y), new Point(control1.X, control1.Y),
            new Point(control2.X, control2.Y), new Point(to.X, to.Y), 0);

    internal void EndFigure() => _current = null;

    private void Add(Point point)
    {
        if (_current is null || double.IsNaN(point.X) || double.IsNaN(point.Y) ||
            double.IsInfinity(point.X) || double.IsInfinity(point.Y) || ++_points > MaximumPoints)
            throw new NotSupportedException("Glyph stroke exceeds its finite contour budget.");
        _current.Add(point);
    }

    private void Flatten(Point a, Point b, Point c, Point d, int depth)
    {
        // Distance to the segment also catches collinear curves that double back.
        if (Math.Max(Distance(b, a, d), Distance(c, a, d)) <= _flatnessPoints)
        {
            Add(d);
            return;
        }
        if (depth >= 20)
            throw new NotSupportedException("Glyph stroke exceeds its curve subdivision budget.");
        var ab = (a + b) * .5; var bc = (b + c) * .5; var cd = (c + d) * .5;
        var abc = (ab + bc) * .5; var bcd = (bc + cd) * .5; var middle = (abc + bcd) * .5;
        Flatten(a, ab, abc, middle, depth + 1);
        Flatten(middle, bcd, cd, d, depth + 1);
    }

    internal XGraphicsPath Build(double width)
    {
        var result = new XGraphicsPath { FillMode = XFillMode.Winding };
        Polygons(width, points =>
        {
            var vertices = new XPoint[points.Length];
            for (var i = 0; i < points.Length; i++) vertices[i] = new XPoint(points[i].X, points[i].Y);
            result.AddPolygon(vertices);
        });
        return result;
    }

    // Distance to the actual filled stroke, including its miter/bevel joins.
    // Native DOCX uses this same geometry to constrain optional emphasis.
    internal double Clearance(XPoint center, double width, Action visit)
    {
        var point = new Point(center.X, center.Y);
        var clearance = double.PositiveInfinity;
        Polygons(width, polygon =>
        {
            var winding = 0;
            for (var i = 0; i < polygon.Length; i++)
            {
                visit(); var a = polygon[i]; var b = polygon[(i + 1) % polygon.Length];
                clearance = Math.Min(clearance, Distance(point, a, b));
                var turn = Cross(b - a, point - a);
                if (a.Y <= point.Y && point.Y < b.Y && turn > 0) winding++;
                else if (b.Y <= point.Y && point.Y < a.Y && turn < 0) winding--;
            }
            if (winding != 0) clearance = 0;
        });
        return clearance;
    }

    private void Polygons(double width, Action<Point[]> emit)
    {
        var radius = width * .5;
        // The old XPen uses the PDF default miter join/limit10 and closed contours.
        // All component polygons have the same orientation: nonzero fill unions
        // overlaps without painting the unoccupied counter interiors.
        void Polygon(params Point[] points)
        {
            var area = 0d;
            for (var i = 0; i < points.Length; i++) area += Cross(points[i], points[(i + 1) % points.Length]);
            if (Math.Abs(area) < 1e-14) return;
            if (area < 0) Array.Reverse(points);
            emit(points);
        }

        foreach (var contour in _contours)
        {
            var points = new List<Point>();
            foreach (var point in contour)
                if (points.Count == 0 || (point - points[points.Count - 1]).Length > 1e-9) points.Add(point);
            if (points.Count > 1 && (points[points.Count - 1] - points[0]).Length <= 1e-9)
                points.RemoveAt(points.Count - 1);
            if (points.Count < 2) continue;
            var directions = new Point[points.Count];
            var normals = new Point[points.Count];
            for (var i = 0; i < points.Count; i++)
            {
                var next = (i + 1) % points.Count;
                var delta = points[next] - points[i];
                directions[i] = delta * (1 / delta.Length);
                normals[i] = new Point(-directions[i].Y, directions[i].X);
                var normal = normals[i] * radius;
                Polygon(points[i] - normal, points[i] + normal, points[next] + normal, points[next] - normal);
            }
            for (var i = 0; i < points.Count; i++)
            {
                var previous = (i + points.Count - 1) % points.Count;
                var turn = Cross(directions[previous], directions[i]);
                if (Math.Abs(turn) < 1e-12) continue;
                var side = -Math.Sign(turn);
                var a = points[i] + normals[previous] * (side * radius);
                var b = points[i] + normals[i] * (side * radius);
                var denominator = 1 + Dot(directions[previous], directions[i]);
                if (denominator > 1e-12)
                {
                    var miter = points[i] + (normals[previous] + normals[i]) * (side * radius / denominator);
                    if ((miter - points[i]).Length <= 10 * radius)
                    {
                        Polygon(points[i], a, miter, b);
                        continue;
                    }
                }
                Polygon(points[i], a, b);
            }
        }
    }

    private static double Dot(Point a, Point b) => a.X * b.X + a.Y * b.Y;
    private static double Cross(Point a, Point b) => a.X * b.Y - a.Y * b.X;
    private static double Distance(Point point, Point a, Point b)
    {
        var delta = b - a;
        var length = Dot(delta, delta);
        var fraction = length > 0 ? Math.Max(0, Math.Min(1, Dot(point - a, delta) / length)) : 0;
        return (point - (a + delta * fraction)).Length;
    }

    private readonly struct Point
    {
        internal readonly double X, Y;
        internal Point(double x, double y) { X = x; Y = y; }
        internal double Length => Math.Sqrt(X * X + Y * Y);
        public static Point operator +(Point a, Point b) => new(a.X + b.X, a.Y + b.Y);
        public static Point operator -(Point a, Point b) => new(a.X - b.X, a.Y - b.Y);
        public static Point operator *(Point a, double factor) => new(a.X * factor, a.Y * factor);
    }
}
