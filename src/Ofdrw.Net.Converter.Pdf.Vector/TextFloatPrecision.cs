using System;
using SkiaSharp;
using UglyToad.PdfPig.Core;
using Binary = Ofdrw.Net.Converter.Pdf.Vector.PathFloatPrecision.Binary;

namespace Ofdrw.Net.Converter.Pdf.Vector;

// Bounds added producer error from callback doubles to the actual emitted float operands.
// Parsed metrics/state and downstream decimal serialization/renderer ink remain separate.
internal sealed class TextFloatPrecision
{
    private readonly Binary _height, _a, _b, _c, _d, _e, _f, _size;
    private readonly bool _finite;
    private Binary _prefix, _lastX, _lastY, _lastCandidateX, _lastCandidateY, _lastDoubleX, _lastDoubleY;
    private double _doublePrefix;
    private bool _first = true;

    internal TextFloatPrecision(SKMatrix candidate, float size, double height)
    {
        _finite = size > 0 && Finite(size) && Finite(height) && Finite(candidate.ScaleX) && Finite(candidate.SkewY) &&
            Finite(candidate.SkewX) && Finite(candidate.ScaleY) && Finite(candidate.TransX) && Finite(candidate.TransY);
        if (!_finite) return;
        _height = Binary.From(height); _a = Binary.From(candidate.ScaleX); _b = Binary.From(-candidate.SkewY);
        _c = Binary.From(-candidate.SkewX); _d = Binary.From(candidate.ScaleY); _e = Binary.From(candidate.TransX);
        _f = _height - Binary.From(candidate.TransY); _size = Binary.From(size);
    }

    internal bool Glyph(TransformationMatrix text, TransformationMatrix ctm, double size, PdfRectangle bounds, float advance)
    {
        if (!_finite || !(size > 0) || !Finite(size) || !Finite(advance) || !Finite(text) || !Finite(ctm) ||
            !Finite(bounds.Left) || !Finite(bounds.Right) || !Finite(bounds.Bottom) || !Finite(bounds.Top)) return false;
        // Independent exact composition, rather than certifying rounded Multiply/inverse results.
        var ta = Binary.From(text.A); var tb = Binary.From(text.B); var tc = Binary.From(text.C); var td = Binary.From(text.D);
        var te = Binary.From(text.E); var tf = Binary.From(text.F);
        var ca = Binary.From(ctm.A); var cb = Binary.From(ctm.B); var cc = Binary.From(ctm.C); var cd = Binary.From(ctm.D);
        var a = ta * ca + tb * cc; var b = ta * cb + tb * cd;
        var c = tc * ca + td * cc; var d = tc * cb + td * cd;
        var x = te * ca + tf * cc + Binary.From(ctm.E); var y = te * cb + tf * cd + Binary.From(ctm.F);
        if (!Map(a, b, c, d, _a, _b, _c, _d)) return false;
        if (!_first) { _prefix += Binary.From(advance); _doublePrefix += advance; }
        if (!Finite(_doublePrefix)) return false;
        var cx = _a * _prefix + _e; var cy = _b * _prefix + _f;
        // Also check the existing consumer's sequential binary64 addition order.
        var doublePrefix = Binary.From(_doublePrefix); var dx = _a * doublePrefix + _e; var dy = _b * doublePrefix + _f;
        if (!Position(x, y, cx, cy) || !Position(x, y, dx, dy)) return false;
        if (!_first && (!Vector(x - _lastX, y - _lastY, cx - _lastCandidateX, cy - _lastCandidateY) ||
                       !Vector(x - _lastX, y - _lastY, dx - _lastDoubleX, dy - _lastDoubleY))) return false;

        var s = Binary.From(size); var sa = s * a; var sb = s * b; var sc = s * c; var sd = s * d;
        var fa = _size * _a; var fb = _size * _b; var fc = _size * _c; var fd = _size * _d;
        if (!Map(sa, sb, sc, sd, fa, fb, fc, fd)) return false;
        // Callback normalized glyph bounds plus the em square. Four corners, no outline reconstruction.
        var left = Binary.From(Math.Min(0, bounds.Left)); var right = Binary.From(Math.Max(1, bounds.Right));
        var bottom = Binary.From(Math.Min(0, bounds.Bottom)); var top = Binary.From(Math.Max(1, bounds.Top));
        bool Corner(Binary gx, Binary gy)
        {
            var sx = x + sa * gx + sc * gy; var sy = y + sb * gx + sd * gy;
            return Position(sx, sy, cx + fa * gx + fc * gy, cy + fb * gx + fd * gy) &&
                   Position(sx, sy, dx + fa * gx + fc * gy, dy + fb * gx + fd * gy);
        }
        if (!Corner(left, bottom) || !Corner(left, top) || !Corner(right, bottom) || !Corner(right, top)) return false;
        _lastX = x; _lastY = y; _lastCandidateX = cx; _lastCandidateY = cy; _lastDoubleX = dx; _lastDoubleY = dy;
        _first = false; return true;
    }

    private static bool Position(Binary x, Binary y, Binary cx, Binary cy) =>
        (Binary.Max((cx - x).Abs, (cy - y).Abs) * 10000).CompareTo(Binary.One) <= 0;
    private static bool Vector(Binary x, Binary y, Binary cx, Binary cy) =>
        (Binary.Max((cx - x).Abs, (cy - y).Abs) * 100000).CompareTo(Binary.Max(x.Abs, y.Abs)) <= 0;
    private static bool Map(Binary a, Binary b, Binary c, Binary d, Binary fa, Binary fb, Binary fc, Binary fd)
    {
        var determinant = a * d - b * c; var candidate = fa * fd - fb * fc;
        if (determinant.Sign == 0 || candidate.Sign != determinant.Sign) return false;
        var da = fa - a; var db = fb - b; var dc = fc - c; var dd = fd - d;
        var row1 = (da * d - dc * b).Abs + (dc * a - da * c).Abs;
        var row2 = (db * d - dd * b).Abs + (dd * a - db * c).Abs;
        return (Binary.Max(row1, row2) * 100000).CompareTo(determinant.Abs) <= 0;
    }
    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static bool Finite(TransformationMatrix m) => Finite(m.A) && Finite(m.B) && Finite(m.C) && Finite(m.D) && Finite(m.E) && Finite(m.F);
}
