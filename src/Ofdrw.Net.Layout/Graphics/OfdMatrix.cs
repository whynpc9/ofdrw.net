using System;

namespace Ofdrw.Net.Layout.Graphics;

/// <summary>Immutable OFD affine matrix: x'=A*x+C*y+E, y'=B*x+D*y+F. Translation uses millimeters.</summary>
public sealed class OfdMatrix
{
    /// <summary>Creates a finite affine matrix in OFD six-value order.</summary>
    public OfdMatrix(double a, double b, double c, double d, double e, double f)
    {
        GraphicsValidation.Finite(a, b, c, d, e, f);
        A = a; B = b; C = c; D = d; E = e; F = f;
    }
    /// <summary>Horizontal scale.</summary>
    public double A { get; }
    /// <summary>Vertical shear.</summary>
    public double B { get; }
    /// <summary>Horizontal shear.</summary>
    public double C { get; }
    /// <summary>Vertical scale.</summary>
    public double D { get; }
    /// <summary>Horizontal translation in millimeters.</summary>
    public double E { get; }
    /// <summary>Vertical translation in millimeters.</summary>
    public double F { get; }
    /// <summary>Identity matrix.</summary>
    public static OfdMatrix Identity { get; } = new(1, 0, 0, 1, 0, 0);
    /// <summary>Creates a translation.</summary>
    public static OfdMatrix Translation(double x, double y) => new(1, 0, 0, 1, x, y);
    /// <summary>Creates a scale about the origin.</summary>
    public static OfdMatrix Scale(double x, double y) => new(x, 0, 0, y, 0, 0);
    /// <summary>Creates a clockwise rotation in degrees in the downward-positive page coordinate system.</summary>
    public static OfdMatrix Rotation(double degrees)
    {
        GraphicsValidation.Finite(degrees);
        var radians = degrees % 360 * Math.PI / 180;
        return new(Math.Cos(radians), Math.Sin(radians), -Math.Sin(radians), Math.Cos(radians), 0, 0);
    }
    /// <summary>Returns this * right; right acts first on a point. Does not mutate either matrix.</summary>
    public OfdMatrix Multiply(OfdMatrix right)
    {
        if (right is null) throw new ArgumentNullException(nameof(right));
        return new(A * right.A + C * right.B, B * right.A + D * right.B,
            A * right.C + C * right.D, B * right.C + D * right.D,
            A * right.E + C * right.F + E, B * right.E + D * right.F + F);
    }
    /// <summary>Transforms a point in millimeters; rejects overflow.</summary>
    public (double X, double Y) TransformPoint(double x, double y)
    {
        GraphicsValidation.Finite(x, y);
        var result = (X: A * x + C * y + E, Y: B * x + D * y + F);
        GraphicsValidation.Finite(result.X, result.Y);
        return result;
    }
    internal double[] ToArray() => new[] { A, B, C, D, E, F };
}
