using System.Globalization;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;

internal static partial class Cli
{
    private static async Task<int> PlainTextAsync(string[] args, CancellationToken token)
    {
        var options = new PlainTextOfdOptions();
        var positional = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        string? input = null, output = null;
        for (var index = 0; index < args.Length; index++)
        {
            var option = args[index];
            if (option is "-h" or "--help") { PrintHelp(); return 0; }
            if (!option.StartsWith('-')) { positional.Add(option); continue; }
            option = option switch { "-i" => "--input", "-o" => "--output", _ => option };
            if (!seen.Add(option)) throw new ArgumentException($"Duplicate option: {option}");
            var value = ReadValue(args, ref index, option);
            switch (option)
            {
                case "--input": input = value; break;
                case "--output": output = value; break;
                case "--font-family": options.Layout.DefaultFontFamily = value; break;
                case "--font-size-mm": options.Layout.DefaultFontSizeMillimeters = PlainTextNumber(value, option); break;
                case "--page-width": options.Layout.PageWidthMillimeters = PlainTextNumber(value, option); break;
                case "--page-height": options.Layout.PageHeightMillimeters = PlainTextNumber(value, option); break;
                case "--margin-left": options.Layout.MarginLeftMillimeters = PlainTextNumber(value, option, true); break;
                case "--margin-right": options.Layout.MarginRightMillimeters = PlainTextNumber(value, option, true); break;
                case "--margin-top": options.Layout.MarginTopMillimeters = PlainTextNumber(value, option, true); break;
                case "--margin-bottom": options.Layout.MarginBottomMillimeters = PlainTextNumber(value, option, true); break;
                case "--max-input-bytes": options.MaxInputBytes = ImageBytes(value, option); break;
                case "--max-characters": options.Layout.MaxCharacters = ImageCount(value, option); break;
                case "--max-pages": options.Layout.MaxPageCount = ImageCount(value, option); break;
                case "--max-text-elements": options.Layout.MaxTextElements = ImageCount(value, option); break;
                default: throw new ArgumentException($"Unknown option for text-to-ofd: {option}");
            }
        }
        if (positional.Count != 0)
        {
            if (positional.Count != 2 || input is not null || output is not null)
                throw new ArgumentException("Use two positional paths or --input and --output, without mixing them.");
            input = positional[0]; output = positional[1];
        }
        if (string.IsNullOrWhiteSpace(input) || string.IsNullOrWhiteSpace(output) || input == "-" || output == "-")
            throw new ArgumentException("Input and output file paths are required; stdin/stdout are unsupported.");
        if (!string.Equals(Path.GetExtension(output), ".ofd", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("text-to-ofd output must end in .ofd.");
        if (seen.Contains("--page-width") != seen.Contains("--page-height"))
            throw new ArgumentException("--page-width and --page-height must be provided together.");
        var converter = new PlainTextOfdConverter(options);
        EnsureDifferentPaths(input, output);
        await using var source = File.OpenRead(input);
        var package = await converter.RenderUtf8Async(source, token).ConfigureAwait(false);
        await using var destination = new AtomicOutput(output);
        await new OfdPackageWriter().WriteAsync(package, destination.Stream, token).ConfigureAwait(false);
        await destination.CommitAsync(token).ConfigureAwait(false);
        WriteSuccessNotification($"Converted {input} -> {output}");
        return 0;
    }

    private static double PlainTextNumber(string value, string option, bool allowZero = false)
    {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ||
            !double.IsFinite(number) || (allowZero ? number < 0 : number <= 0))
            throw new ArgumentException($"{option} requires a finite {(allowZero ? "nonnegative" : "positive")} number.");
        return number;
    }

}
