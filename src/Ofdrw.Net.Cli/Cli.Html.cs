using System.Globalization;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;

internal static partial class Cli
{
    private static async Task<int> ConvertHtmlAsync(string[] args, CancellationToken token)
    {
        string? input = null, output = null, pages = null;
        var options = new OfdHtmlOptions();
        for (var i = 0; i < args.Length; i++)
        {
            token.ThrowIfCancellationRequested(); var option = args[i];
            if (option is "-h" or "--help") { PrintHelp(); return 0; }
            if (!option.StartsWith('-'))
            { if (input is null) input = option; else if (output is null) output = option; else throw new ArgumentException("Too many HTML positional arguments."); continue; }
            var value = ReadValue(args, ref i, option);
            switch (option)
            {
                case "--input": case "-i": input = value; break;
                case "--output": case "-o": output = value; break;
                case "--pages": case "-p": pages = value; break;
                case "--title": options.Title = value; break;
                case "--max-pages": options.MaxPages = checked((int)HtmlLimit(value, option)); break;
                case "--max-output-bytes": options.MaxOutputBytes = HtmlLimit(value, option); break;
                case "--max-page-bytes": options.MaxPageBytes = HtmlLimit(value, option); break;
                default: throw new ArgumentException("Unknown HTML option: " + option);
            }
        }
        if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output)) throw new ArgumentException("Input and output paths are required.");
        var converter = new OfdToHtmlConverter(options);
        var selected = ParseHtmlPages(pages, options.MaxPages, token);
        EnsureDifferentPaths(input, output);
        await using var source = File.OpenRead(input);
        await using var destination = new AtomicOutput(output);
        await converter.ConvertAsync(source, destination.Stream, selected, token).ConfigureAwait(false);
        await destination.CommitAsync(token).ConfigureAwait(false);
        WriteSuccessNotification($"Converted {input} -> {output}"); return 0;
    }
    private static long HtmlLimit(string value, string name)
    { if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var n) || n <= 0) throw new ArgumentException(name + " requires a positive integer."); return n; }
    internal static IReadOnlyList<int>? ParseHtmlPages(string? value, int maximum, CancellationToken token)
    {
        if (value is null) return null;
        if (maximum <= 0 || value.Length == 0 || value.Length > 8192) throw new ArgumentException("Invalid HTML page selection length/limit.");
        var result = new List<int>();
        foreach (var segment in value.Split(','))
        {
            token.ThrowIfCancellationRequested(); var parts = segment.Trim().Split('-');
            if (parts.Length is < 1 or > 2 || !int.TryParse(parts[0], NumberStyles.None, CultureInfo.InvariantCulture, out var first) || first <= 0)
                throw new ArgumentException("Invalid HTML page selection.");
            var last = first;
            if (parts.Length == 2 && (!int.TryParse(parts[1], NumberStyles.None, CultureInfo.InvariantCulture, out last) || last < first))
                throw new ArgumentException("Invalid HTML page range.");
            var count = (long)last - first + 1;
            if (count > maximum - (long)result.Count) throw new ArgumentException("HTML page selection capacity exceeded.");
            for (long page = first; page <= last; page++) { token.ThrowIfCancellationRequested(); result.Add((int)page - 1); }
        }
        return result;
    }
}
