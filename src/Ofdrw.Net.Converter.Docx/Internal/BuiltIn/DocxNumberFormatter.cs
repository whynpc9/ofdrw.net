using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Ofdrw.Net.Converter.Docx.Internal.BuiltIn;

// Pure OOXML list-number formatting; unsupported formats degrade to decimal.
internal static class DocxNumberFormatter
{
    private const string ChineseDigits = "〇一二三四五六七八九";

    internal static string Format(int value, NumberFormatValues? format, out bool degraded)
    {
        degraded = false;
        if (format is null || format == NumberFormatValues.Decimal) return Decimal(value);
        if (format == NumberFormatValues.DecimalZero)
            return value is >= 0 and < 10 ? "0" + Decimal(value) : Decimal(value);
        if (format == NumberFormatValues.LowerRoman) return ToRoman(value)?.ToLowerInvariant() ?? Decimal(value);
        if (format == NumberFormatValues.UpperRoman) return ToRoman(value) ?? Decimal(value);
        if (format == NumberFormatValues.LowerLetter) return ToLetters(value)?.ToLowerInvariant() ?? Decimal(value);
        if (format == NumberFormatValues.UpperLetter) return ToLetters(value) ?? Decimal(value);
        if (format == NumberFormatValues.ChineseCounting || format == NumberFormatValues.ChineseCountingThousand)
        {
            var chinese = ToChineseCounting(value);
            if (chinese is not null) return chinese;
            degraded = true;
            return Decimal(value);
        }
        if (format == NumberFormatValues.IdeographDigital) return ToIdeographDigits(value);
        degraded = true;
        return Decimal(value);
    }

    internal static string? ToRoman(int value)
    {
        if (value is < 1 or > 3999) return null;
        var numerals = new[] { (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"), (90, "XC"),
            (50, "L"), (40, "XL"), (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I") };
        var builder = new StringBuilder();
        foreach (var (amount, text) in numerals)
            while (value >= amount) { builder.Append(text); value -= amount; }
        return builder.ToString();
    }

    // Word style: 1→A … 26→Z, 27→AA, 52→ZZ, 53→AAA (the letter repeats ⌈n/26⌉ times).
    internal static string? ToLetters(int value)
    {
        if (value < 1 || value > 26 * 780) return null;
        var letter = (char)('A' + (value - 1) % 26);
        return new string(letter, (value - 1) / 26 + 1);
    }

    internal static string? ToChineseCounting(int value)
    {
        if (value is < 1 or > 99) return null;
        if (value < 10) return ChineseDigits[value].ToString();
        var tens = value / 10;
        var ones = value % 10;
        return (tens == 1 ? "" : ChineseDigits[tens].ToString()) + "十" + (ones == 0 ? "" : ChineseDigits[ones].ToString());
    }

    internal static string ToIdeographDigits(int value)
    {
        if (value < 0) return Decimal(value);
        var digits = Decimal(value);
        var builder = new StringBuilder(digits.Length);
        foreach (var digit in digits) builder.Append(ChineseDigits[digit - '0']);
        return builder.ToString();
    }

    private static string Decimal(int value) => value.ToString(CultureInfo.InvariantCulture);
}
