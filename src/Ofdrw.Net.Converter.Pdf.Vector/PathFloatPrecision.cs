using System;
using System.Numerics;
using SkiaSharp;
using UglyToad.PdfPig.Core;

namespace Ofdrw.Net.Converter.Pdf.Vector;

// Bounds only added producer conversion error relative to parsed doubles, not renderer ink/topology.
internal sealed class PathFloatPrecision
{
    private readonly Binary _a, _b, _c, _d, _e, _f, _height;
    private readonly Binary _fa, _fb, _fc, _fd, _fe, _fy;
    private readonly bool _finite;
    private bool _failed, _fillClosureFailed;
    internal PathFloatPrecision(TransformationMatrix matrix, double height)
    {
        _finite = Finite(matrix.A) && Finite(matrix.B) && Finite(matrix.C) && Finite(matrix.D) && Finite(matrix.E) && Finite(matrix.F) && Finite(height) &&
            Finite((float)matrix.A) && Finite((float)matrix.B) && Finite((float)matrix.C) && Finite((float)matrix.D) && Finite((float)matrix.E) && Finite((float)(height - matrix.F));
        if (!_finite) return;
        _a = Binary.From(matrix.A); _b = Binary.From(matrix.B); _c = Binary.From(matrix.C); _d = Binary.From(matrix.D);
        _e = Binary.From(matrix.E); _f = Binary.From(matrix.F); _height = Binary.From(height);
        _fa = Binary.From((float)matrix.A); _fb = Binary.From((float)matrix.B); _fc = Binary.From((float)matrix.C); _fd = Binary.From((float)matrix.D);
        _fe = Binary.From((float)matrix.E); _fy = Binary.From((float)(height - matrix.F));
    }
    internal readonly struct Point
    {
        internal readonly Binary X, Y; internal readonly SKPoint Candidate;
        internal Point(double x, double y, SKPoint candidate) : this(Binary.From(x), Binary.From(y), candidate) { }
        internal Point(Binary x, Binary y, SKPoint candidate) { X = x; Y = y; Candidate = candidate; }
    }
    internal void Primitive(params Point[] points)
    {
        if (_failed) return;
        if (!_finite) { _failed = true; return; }
        for (var i = 0; i < points.Length; i++)
        {
            if (!Position(points[i])) { _failed = true; return; }
            for (var j = 0; j < i; j++)
                if (!Vector(points[j], points[i])) { _failed = true; return; }
        }
    }
    internal void ImplicitClosure(Point start, Point end)
    { if (!_finite || !Vector(start, end)) _fillClosureFailed = true; }
    private (Binary X, Binary Y) Source(Point p) => (_a * p.X + _c * p.Y + _e, _height - (_b * p.X + _d * p.Y + _f));
    private (Binary X, Binary Y) Candidate(Point p)
    {
        var x = Binary.From(p.Candidate.X); var y = Binary.From(p.Candidate.Y);
        return (_fa * x + _fc * y + _fe, _fy - (_fb * x + _fd * y));
    }
    private bool Position(Point p)
    {
        var source = Source(p); var candidate = Candidate(p);
        return ((candidate.X - source.X).Abs * 10000).CompareTo(Binary.One) <= 0 &&
            ((candidate.Y - source.Y).Abs * 10000).CompareTo(Binary.One) <= 0;
    }
    private bool Vector(Point first, Point second)
    {
        var sourceFirst = Source(first); var sourceSecond = Source(second);
        var candidateFirst = Candidate(first); var candidateSecond = Candidate(second);
        var x = sourceSecond.X - sourceFirst.X; var y = sourceSecond.Y - sourceFirst.Y;
        var error = Binary.Max((candidateSecond.X - candidateFirst.X - x).Abs, (candidateSecond.Y - candidateFirst.Y - y).Abs);
        return (error * 100000).CompareTo(Binary.Max(x.Abs, y.Abs)) <= 0;
    }
    internal bool Accept(bool fill, bool stroke, double width, float candidateWidth)
    {
        if (!_finite || _failed || (fill && _fillClosureFailed)) return false;
        var determinant = _a * _d - _b * _c;
        var candidateDeterminant = _fa * _fd - _fb * _fc;
        if (determinant.Sign == 0 || determinant.Sign != candidateDeterminant.Sign) return false;
        // Relative infinity-norm distortion of the entire linear map, including stroke normals.
        var da = _fa - _a; var db = _fb - _b; var dc = _fc - _c; var dd = _fd - _d;
        var row1 = (da * _d - dc * _b).Abs + (dc * _a - da * _c).Abs;
        var row2 = (db * _d - dd * _b).Abs + (dd * _a - db * _c).Abs;
        if ((Binary.Max(row1, row2) * 100000).CompareTo(determinant.Abs) > 0) return false;
        if (stroke)
        {
            if (!(candidateWidth > 0) || !Finite(candidateWidth)) return false;
            var w = Binary.From(width); var fw = Binary.From(candidateWidth);
            var error1 = (fw * _fa - w * _a).Abs + (fw * _fc - w * _c).Abs;
            var error2 = (fw * _fb - w * _b).Abs + (fw * _fd - w * _d).Abs;
            // Fixed miter10: half-width times miter bound5, added page-space error <=1/10000pt.
            if ((Binary.Max(error1, error2) * 50000).CompareTo(Binary.One) > 0) return false;
        }
        return true;
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    // Exact dyadic arithmetic over bounded binary64/binary32 operands. No accumulated page-wide sums.
    internal readonly struct Binary : IComparable<Binary>
    {
        private readonly BigInteger _value; private readonly int _exponent;
        private Binary(BigInteger value, int exponent) { _value = value; _exponent = value.IsZero ? 0 : exponent; }
        internal static Binary One => new(BigInteger.One, 0);
        internal int Sign => _value.Sign;
        internal Binary Abs => new(BigInteger.Abs(_value), _exponent);
        internal static Binary From(double value)
        {
            if (!Finite(value)) throw new ArgumentException("Binary precision operand must be finite.");
            var bits = unchecked((ulong)BitConverter.DoubleToInt64Bits(value));
            var fraction = bits & 0x000ffffffffffffful; var exponent = (int)((bits >> 52) & 0x7ff);
            var integer = new BigInteger(exponent == 0 ? fraction : fraction | 0x0010000000000000ul);
            if ((bits >> 63) != 0) integer = -integer;
            return new Binary(integer, exponent == 0 ? -1074 : exponent - 1023 - 52);
        }
        public int CompareTo(Binary other) => (this - other)._value.Sign;
        internal static Binary Max(Binary left, Binary right) => left.CompareTo(right) >= 0 ? left : right;
        public static Binary operator +(Binary left, Binary right)
        {
            if (left._value.IsZero) return right; if (right._value.IsZero) return left;
            var exponent = Math.Min(left._exponent, right._exponent);
            return new Binary((left._value << (left._exponent - exponent)) + (right._value << (right._exponent - exponent)), exponent);
        }
        public static Binary operator -(Binary left, Binary right) => left + new Binary(-right._value, right._exponent);
        public static Binary operator *(Binary left, Binary right) => new(left._value * right._value, left._exponent + right._exponent);
        public static Binary operator *(Binary left, int right) => new(left._value * right, left._exponent);
    }
}
