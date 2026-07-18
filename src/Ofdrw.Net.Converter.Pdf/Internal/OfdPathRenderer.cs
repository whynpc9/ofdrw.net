using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text.RegularExpressions;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Drawing;

namespace Ofdrw.Net.Converter.Pdf.Internal;

internal static class OfdPathRenderer
{
    private static readonly Regex TokenPattern = new(
        @"[A-Za-z]|[-+]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][-+]?\d+)?",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static bool TryDraw(
        XGraphics graphics,
        OfdPathElement element,
        double pageOriginX,
        double pageOriginY)
    {
        if (string.IsNullOrWhiteSpace(element.AbbreviatedData) ||
            (!element.Stroke && !element.Fill))
        {
            return false;
        }

        try
        {
            var tokens = Tokenize(element.AbbreviatedData);
            var path = new XGraphicsPath();
            var state = new PathState(path, element, pageOriginX, pageOriginY);
            if (!state.Build(tokens))
            {
                return false;
            }

            XPen? pen = null;
            XSolidBrush? brush = null;
            if (element.Stroke)
            {
                var scale = GetLineScale(element.Transform);
                pen = new XPen(
                    ToXColor(element.StrokeColor),
                    MillimetersToPoints(Math.Max(element.LineWidthMillimeters * scale, 0.01d)));
            }

            if (element.Fill)
            {
                brush = new XSolidBrush(ToXColor(element.FillColor ?? element.StrokeColor));
            }

            if (pen is not null && brush is not null)
            {
                graphics.DrawPath(pen, brush, path);
            }
            else if (pen is not null)
            {
                graphics.DrawPath(pen, path);
            }
            else if (brush is not null)
            {
                graphics.DrawPath(brush, path);
            }

            return true;
        }
        catch
        {
            return false;
        }
    }

    private static IReadOnlyList<string> Tokenize(string data)
    {
        var matches = TokenPattern.Matches(data);
        var tokens = new List<string>(matches.Count);
        foreach (Match match in matches)
        {
            tokens.Add(match.Value);
        }

        return tokens;
    }

    private static XColor ToXColor(OfdColor color)
    {
        return XColor.FromArgb(color.Alpha, color.Red, color.Green, color.Blue);
    }

    private static double GetLineScale(double[]? transform)
    {
        if (transform is not { Length: 6 })
        {
            return 1d;
        }

        var xScale = Math.Sqrt((transform[0] * transform[0]) + (transform[1] * transform[1]));
        var yScale = Math.Sqrt((transform[2] * transform[2]) + (transform[3] * transform[3]));
        var average = (xScale + yScale) / 2d;
        return average > 0 ? average : 1d;
    }

    private static double MillimetersToPoints(double millimeters)
    {
        return millimeters * 72d / 25.4d;
    }

    private sealed class PathState
    {
        private readonly XGraphicsPath _path;
        private readonly OfdPathElement _element;
        private readonly double _pageOriginX;
        private readonly double _pageOriginY;
        private PointD _current;
        private PointD _figureStart;
        private bool _hasGeometry;
        private char _command;

        public PathState(
            XGraphicsPath path,
            OfdPathElement element,
            double pageOriginX,
            double pageOriginY)
        {
            _path = path;
            _element = element;
            _pageOriginX = pageOriginX;
            _pageOriginY = pageOriginY;
        }

        public bool Build(IReadOnlyList<string> tokens)
        {
            var index = 0;
            while (index < tokens.Count)
            {
                if (IsCommand(tokens[index]))
                {
                    _command = tokens[index][0];
                    index++;
                }
                else if (_command == '\0')
                {
                    return false;
                }

                var upper = char.ToUpperInvariant(_command);
                var relative = char.IsLower(_command);
                switch (upper)
                {
                    case 'M':
                        {
                            if (!TryReadPoint(tokens, ref index, relative, out var point))
                            {
                                return false;
                            }

                            _path.StartFigure();
                            _current = point;
                            _figureStart = point;
                            _command = relative ? 'l' : 'L';
                            break;
                        }
                    case 'L':
                        {
                            if (!TryReadPoint(tokens, ref index, relative, out var point))
                            {
                                return false;
                            }

                            AddLine(_current, point);
                            _current = point;
                            break;
                        }
                    case 'H':
                        {
                            if (!TryReadNumber(tokens, ref index, out var x))
                            {
                                return false;
                            }

                            var point = new PointD(relative ? _current.X + x : x, _current.Y);
                            AddLine(_current, point);
                            _current = point;
                            break;
                        }
                    case 'V':
                        {
                            if (!TryReadNumber(tokens, ref index, out var y))
                            {
                                return false;
                            }

                            var point = new PointD(_current.X, relative ? _current.Y + y : y);
                            AddLine(_current, point);
                            _current = point;
                            break;
                        }
                    case 'B':
                        {
                            if (!TryReadPoint(tokens, ref index, relative, out var control1) ||
                                !TryReadPoint(tokens, ref index, relative, out var control2) ||
                                !TryReadPoint(tokens, ref index, relative, out var end))
                            {
                                return false;
                            }

                            AddBezier(_current, control1, control2, end);
                            _current = end;
                            break;
                        }
                    case 'Q':
                        {
                            if (!TryReadPoint(tokens, ref index, relative, out var control) ||
                                !TryReadPoint(tokens, ref index, relative, out var end))
                            {
                                return false;
                            }

                            var control1 = new PointD(
                                _current.X + ((control.X - _current.X) * 2d / 3d),
                                _current.Y + ((control.Y - _current.Y) * 2d / 3d));
                            var control2 = new PointD(
                                end.X + ((control.X - end.X) * 2d / 3d),
                                end.Y + ((control.Y - end.Y) * 2d / 3d));
                            AddBezier(_current, control1, control2, end);
                            _current = end;
                            break;
                        }
                    case 'A':
                        {
                            if (!TryReadNumber(tokens, ref index, out var radiusX) ||
                                !TryReadNumber(tokens, ref index, out var radiusY) ||
                                !TryReadNumber(tokens, ref index, out var rotation) ||
                                !TryReadNumber(tokens, ref index, out var largeArc) ||
                                !TryReadNumber(tokens, ref index, out var sweep) ||
                                !TryReadPoint(tokens, ref index, relative, out var end))
                            {
                                return false;
                            }

                            AddArc(_current, end, radiusX, radiusY, rotation, largeArc != 0d, sweep != 0d);
                            _current = end;
                            break;
                        }
                    case 'C':
                    case 'Z':
                        _path.CloseFigure();
                        _current = _figureStart;
                        _hasGeometry = true;
                        _command = '\0';
                        break;
                    default:
                        return false;
                }
            }

            return _hasGeometry;
        }

        private bool TryReadPoint(
            IReadOnlyList<string> tokens,
            ref int index,
            bool relative,
            out PointD point)
        {
            point = default;
            if (!TryReadNumber(tokens, ref index, out var x) ||
                !TryReadNumber(tokens, ref index, out var y))
            {
                return false;
            }

            point = relative
                ? new PointD(_current.X + x, _current.Y + y)
                : new PointD(x, y);
            return true;
        }

        private static bool TryReadNumber(IReadOnlyList<string> tokens, ref int index, out double value)
        {
            value = 0d;
            if (index >= tokens.Count || IsCommand(tokens[index]))
            {
                return false;
            }

            return double.TryParse(
                tokens[index++],
                NumberStyles.Float,
                CultureInfo.InvariantCulture,
                out value);
        }

        private void AddLine(PointD start, PointD end)
        {
            var transformedStart = Transform(start);
            var transformedEnd = Transform(end);
            _path.AddLine(
                transformedStart.X,
                transformedStart.Y,
                transformedEnd.X,
                transformedEnd.Y);
            _hasGeometry = true;
        }

        private void AddBezier(PointD start, PointD control1, PointD control2, PointD end)
        {
            var transformedStart = Transform(start);
            var transformedControl1 = Transform(control1);
            var transformedControl2 = Transform(control2);
            var transformedEnd = Transform(end);
            _path.AddBezier(
                transformedStart.X,
                transformedStart.Y,
                transformedControl1.X,
                transformedControl1.Y,
                transformedControl2.X,
                transformedControl2.Y,
                transformedEnd.X,
                transformedEnd.Y);
            _hasGeometry = true;
        }

        private void AddArc(
            PointD start,
            PointD end,
            double radiusX,
            double radiusY,
            double rotationDegrees,
            bool largeArc,
            bool sweep)
        {
            radiusX = Math.Abs(radiusX);
            radiusY = Math.Abs(radiusY);
            if (radiusX <= double.Epsilon ||
                radiusY <= double.Epsilon ||
                (Math.Abs(start.X - end.X) <= double.Epsilon &&
                 Math.Abs(start.Y - end.Y) <= double.Epsilon))
            {
                AddLine(start, end);
                return;
            }

            var rotation = rotationDegrees * Math.PI / 180d;
            var cos = Math.Cos(rotation);
            var sin = Math.Sin(rotation);
            var dx = (start.X - end.X) / 2d;
            var dy = (start.Y - end.Y) / 2d;
            var x1Prime = (cos * dx) + (sin * dy);
            var y1Prime = (-sin * dx) + (cos * dy);

            var radiiScale =
                ((x1Prime * x1Prime) / (radiusX * radiusX)) +
                ((y1Prime * y1Prime) / (radiusY * radiusY));
            if (radiiScale > 1d)
            {
                var scale = Math.Sqrt(radiiScale);
                radiusX *= scale;
                radiusY *= scale;
            }

            var rx2 = radiusX * radiusX;
            var ry2 = radiusY * radiusY;
            var numerator =
                (rx2 * ry2) -
                (rx2 * y1Prime * y1Prime) -
                (ry2 * x1Prime * x1Prime);
            var denominator =
                (rx2 * y1Prime * y1Prime) +
                (ry2 * x1Prime * x1Prime);
            var sign = largeArc == sweep ? -1d : 1d;
            var factor = denominator <= double.Epsilon
                ? 0d
                : sign * Math.Sqrt(Math.Max(0d, numerator / denominator));
            var centerXPrime = factor * radiusX * y1Prime / radiusY;
            var centerYPrime = factor * -radiusY * x1Prime / radiusX;
            var center = new PointD(
                (cos * centerXPrime) - (sin * centerYPrime) + ((start.X + end.X) / 2d),
                (sin * centerXPrime) + (cos * centerYPrime) + ((start.Y + end.Y) / 2d));

            var startVector = new PointD(
                (x1Prime - centerXPrime) / radiusX,
                (y1Prime - centerYPrime) / radiusY);
            var endVector = new PointD(
                (-x1Prime - centerXPrime) / radiusX,
                (-y1Prime - centerYPrime) / radiusY);
            var startAngle = Math.Atan2(startVector.Y, startVector.X);
            var deltaAngle = VectorAngle(startVector, endVector);
            if (!sweep && deltaAngle > 0d)
            {
                deltaAngle -= 2d * Math.PI;
            }
            else if (sweep && deltaAngle < 0d)
            {
                deltaAngle += 2d * Math.PI;
            }

            var segments = Math.Max(1, (int)Math.Ceiling(Math.Abs(deltaAngle) / (Math.PI / 2d)));
            var segmentAngle = deltaAngle / segments;
            var segmentStart = startAngle;
            var segmentStartPoint = start;
            for (var i = 0; i < segments; i++)
            {
                var segmentEnd = segmentStart + segmentAngle;
                var alpha = 4d / 3d * Math.Tan(segmentAngle / 4d);
                var startDerivative = EllipseDerivative(
                    radiusX, radiusY, rotation, segmentStart);
                var endDerivative = EllipseDerivative(
                    radiusX, radiusY, rotation, segmentEnd);
                var segmentEndPoint = EllipsePoint(
                    center, radiusX, radiusY, rotation, segmentEnd);
                var control1 = new PointD(
                    segmentStartPoint.X + (alpha * startDerivative.X),
                    segmentStartPoint.Y + (alpha * startDerivative.Y));
                var control2 = new PointD(
                    segmentEndPoint.X - (alpha * endDerivative.X),
                    segmentEndPoint.Y - (alpha * endDerivative.Y));
                AddBezier(segmentStartPoint, control1, control2, segmentEndPoint);
                segmentStartPoint = segmentEndPoint;
                segmentStart = segmentEnd;
            }
        }

        private PointD Transform(PointD point)
        {
            var x = point.X;
            var y = point.Y;
            if (_element.Transform is { Length: 6 } matrix)
            {
                var transformedX = (matrix[0] * x) + (matrix[2] * y) + matrix[4];
                var transformedY = (matrix[1] * x) + (matrix[3] * y) + matrix[5];
                x = transformedX;
                y = transformedY;
            }

            return new PointD(
                MillimetersToPoints(_element.XMillimeters + x - _pageOriginX),
                MillimetersToPoints(_element.YMillimeters + y - _pageOriginY));
        }

        private static PointD EllipsePoint(
            PointD center,
            double radiusX,
            double radiusY,
            double rotation,
            double angle)
        {
            var cosRotation = Math.Cos(rotation);
            var sinRotation = Math.Sin(rotation);
            var cosAngle = Math.Cos(angle);
            var sinAngle = Math.Sin(angle);
            return new PointD(
                center.X + (radiusX * cosRotation * cosAngle) - (radiusY * sinRotation * sinAngle),
                center.Y + (radiusX * sinRotation * cosAngle) + (radiusY * cosRotation * sinAngle));
        }

        private static PointD EllipseDerivative(
            double radiusX,
            double radiusY,
            double rotation,
            double angle)
        {
            var cosRotation = Math.Cos(rotation);
            var sinRotation = Math.Sin(rotation);
            var cosAngle = Math.Cos(angle);
            var sinAngle = Math.Sin(angle);
            return new PointD(
                (-radiusX * cosRotation * sinAngle) - (radiusY * sinRotation * cosAngle),
                (-radiusX * sinRotation * sinAngle) + (radiusY * cosRotation * cosAngle));
        }

        private static double VectorAngle(PointD from, PointD to)
        {
            return Math.Atan2(
                (from.X * to.Y) - (from.Y * to.X),
                (from.X * to.X) + (from.Y * to.Y));
        }
    }

    private readonly struct PointD
    {
        public PointD(double x, double y)
        {
            X = x;
            Y = y;
        }

        public double X { get; }

        public double Y { get; }
    }

    private static bool IsCommand(string token)
    {
        return token.Length == 1 && char.IsLetter(token[0]);
    }
}
