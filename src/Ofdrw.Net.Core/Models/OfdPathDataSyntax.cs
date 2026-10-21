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
        return data.Length == 1 && !data[0].HasElements && IsValid(data[0].Value, token);
    }

    private static bool IsValid(string value, CancellationToken token)
    {
        var index = 0;
        var command = '\0';
        var operands = 0;
        var sawMove = false;
        while (index < value.Length)
        {
            token.ThrowIfCancellationRequested();
            while (index < value.Length && (char.IsWhiteSpace(value[index]) || value[index] == ',')) index++;
            if (index == value.Length) break;
            var current = value[index];
            if (char.IsLetter(current))
            {
                if (!Complete(command, operands)) return false;
                command = char.ToUpperInvariant(current);
                if (command is not ('M' or 'L' or 'B' or 'Q' or 'A' or 'C') ||
                    !sawMove && command != 'M') return false;
                if (command == 'M') sawMove = true;
                operands = 0; index++;
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
        }
        return sawMove && Complete(command, operands);
    }

    private static bool Complete(char command, int operands)
    {
        if (command == '\0') return true;
        var arity = command switch { 'M' or 'L' => 2, 'B' => 6, 'Q' => 4, 'A' => 7, 'C' => 0, _ => -1 };
        return arity == 0 ? operands == 0 : arity > 0 && operands >= arity && operands % arity == 0;
    }

    private static bool Digit(char value) => value is >= '0' and <= '9';
}
