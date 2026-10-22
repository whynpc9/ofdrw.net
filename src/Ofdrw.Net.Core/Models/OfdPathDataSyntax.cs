using System;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Admit original path data before a renderer can silently skip or normalize it.
internal static class OfdPathDataSyntax
{
    internal static bool HasSingleValidDrawing(XElement path, CancellationToken token)
    {
        var data = path.Elements(path.Name.Namespace + "AbbreviatedData").ToArray();
        return data.Length == 1 && !data[0].HasElements && IsValidLiteral(data[0].Value, token, requireDrawableSegment: true);
    }

    internal static bool IsValidLiteral(string value, CancellationToken token, bool requireDrawableSegment = false)
    {
        var index = 0;
        var command = '\0';
        var operands = 0;
        var sawMove = false;
        var sawSegment = false;
        var previousNumber = false;
        while (index < value.Length)
        {
            token.ThrowIfCancellationRequested();
            var comma = false;
            while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ','))
            {
                if (value[index] == ',')
                {
                    if (comma || !previousNumber) return false;
                    comma = true;
                }
                index++;
            }
            if (index == value.Length)
            { if (comma) return false; break; }
            var current = value[index];
            if (char.IsLetter(current))
            {
                if (comma) return false;
                if (!Complete(command, operands)) return false;
                if (command == 'M' && operands > 2 ||
                    (command is 'L' or 'B' or 'Q' or 'A') && operands > 0)
                    sawSegment = true;
                command = char.ToUpperInvariant(current);
                if (command is not ('M' or 'L' or 'B' or 'Q' or 'A' or 'C') ||
                    !sawMove && command != 'M') return false;
                if (command == 'M') sawMove = true;
                operands = 0; previousNumber = false; index++;
                continue;
            }
            if (command is '\0' or 'C') return false;
            var start = index;
            if (value[index] is '+' or '-') index++;
            var digits = 0;
            while (index < value.Length && Digit(value[index])) { index++; digits++; }
            if (index < value.Length && value[index] == '.')
            {
                index++;
                while (index < value.Length && Digit(value[index])) { index++; digits++; }
            }
            if (digits == 0) return false;
            if (index < value.Length && value[index] is 'e' or 'E')
            {
                index++;
                if (index < value.Length && value[index] is '+' or '-') index++;
                var exponent = index;
                while (index < value.Length && Digit(value[index])) index++;
                if (exponent == index) return false;
            }
            var literal = value.Substring(start, index - start);
            if (!double.TryParse(literal, NumberStyles.Float,
                    CultureInfo.InvariantCulture, out var number) || double.IsNaN(number) || double.IsInfinity(number))
                return false;
            if (command == 'A')
            {
                var slot = operands % 7;
                if ((slot is 0 or 1) && number <= 0 ||
                    (slot is 3 or 4) && literal is not ("0" or "1")) return false;
            }
            operands++;
            previousNumber = true;
        }
        if (command == 'M' && operands > 2 ||
            (command is 'L' or 'B' or 'Q' or 'A') && operands > 0)
            sawSegment = true;
        return sawMove && (!requireDrawableSegment || sawSegment) && Complete(command, operands);
    }

    private static bool Complete(char command, int operands)
    {
        if (command == '\0') return true;
        var arity = command switch { 'M' or 'L' => 2, 'B' => 6, 'Q' => 4, 'A' => 7, 'C' => 0, _ => -1 };
        return arity == 0 ? operands == 0 : arity > 0 && operands >= arity && operands % arity == 0;
    }

    private static bool Digit(char value) => value is >= '0' and <= '9';
}
