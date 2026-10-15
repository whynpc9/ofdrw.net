using System.Globalization;
using System.Numerics;

namespace Ofdrw.Net.Core.Models;

internal static class OfdNumericFormat
{
    internal static bool Nonsingular(double a, double b, double c, double d)
    {
        var leftA = DecimalParts(a); var leftD = DecimalParts(d);
        var rightB = DecimalParts(b); var rightC = DecimalParts(c);
        var leftScale = leftA.Scale + leftD.Scale;
        var rightScale = rightB.Scale + rightC.Scale;
        var scale = System.Math.Max(leftScale, rightScale);
        // Each input comes from finite-double Plain output; decimal lengths and
        // product scales are bounded, without decimal's 28-digit range limit.
        var left = leftA.Value * leftD.Value * BigInteger.Pow(10, scale - leftScale);
        var right = rightB.Value * rightC.Value * BigInteger.Pow(10, scale - rightScale);
        return left != right;
    }
    private static (BigInteger Value, int Scale) DecimalParts(double value)
    {
        var literal = Plain(value);
        var dot = literal.IndexOf('.');
        return (BigInteger.Parse(literal.Replace(".", ""), CultureInfo.InvariantCulture), dot < 0 ? 0 : literal.Length - dot - 1);
    }
    internal static string Plain(double value)
    {
        // Preserve round-trip digits, but expand an exponent: OFD path letters
        // must never be confused with an E token by decimal-only readers.
        var literal = value.ToString("R", CultureInfo.InvariantCulture);
        var exponentIndex = literal.IndexOf('E');
        if (exponentIndex < 0) return literal;
        var negative = literal[0] == '-';
        var mantissa = literal.Substring(negative ? 1 : 0, exponentIndex - (negative ? 1 : 0));
        var dot = mantissa.IndexOf('.');
        var decimalPosition = (dot < 0 ? mantissa.Length : dot) + int.Parse(literal.Substring(exponentIndex + 1), CultureInfo.InvariantCulture);
        var digits = mantissa.Replace(".", "");
        var expanded = decimalPosition <= 0 ? "0." + new string('0', -decimalPosition) + digits
            : decimalPosition >= digits.Length ? digits + new string('0', decimalPosition - digits.Length)
            : digits.Insert(decimalPosition, ".");
        return negative ? "-" + expanded : expanded;
    }
}
