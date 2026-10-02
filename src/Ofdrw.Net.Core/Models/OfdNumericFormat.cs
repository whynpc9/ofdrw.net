using System.Globalization;

namespace Ofdrw.Net.Core.Models;

internal static class OfdNumericFormat
{
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
