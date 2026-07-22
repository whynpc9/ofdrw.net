using System.Text;

namespace Ofdrw.Net.Converter.Docx.Internal;

internal static class ProcessArguments
{
    internal static string Join(params string[] arguments)
    {
        var result = new StringBuilder();
        foreach (var argument in arguments)
        {
            if (result.Length > 0)
            {
                result.Append(' ');
            }

            AppendQuoted(result, argument);
        }

        return result.ToString();
    }

    private static void AppendQuoted(StringBuilder result, string argument)
    {
        if (argument.Length > 0 && argument.IndexOfAny(new[] { ' ', '\t', '\n', '\v', '"' }) < 0)
        {
            result.Append(argument);
            return;
        }

        result.Append('"');
        var backslashes = 0;
        foreach (var character in argument)
        {
            if (character == '\\')
            {
                backslashes++;
                continue;
            }

            if (character == '"')
            {
                result.Append('\\', backslashes * 2 + 1);
                result.Append('"');
                backslashes = 0;
                continue;
            }

            result.Append('\\', backslashes);
            result.Append(character);
            backslashes = 0;
        }

        result.Append('\\', backslashes * 2);
        result.Append('"');
    }
}
