using Ofdrw.Net.Converter.Pdf.Converters;

return await Cli.RunAsync(args);

internal static class Cli
{
    public static async Task<int> RunAsync(string[] args)
    {
        if (args.Length == 0 || IsHelp(args[0]))
        {
            PrintHelp();
            return args.Length == 0 ? 1 : 0;
        }

        var command = args[0].Trim().ToLowerInvariant();
        if (command is not ("convert" or "pdf-to-ofd" or "ofd-to-pdf"))
        {
            Console.Error.WriteLine($"Unknown command: {args[0]}");
            PrintHelp();
            return 1;
        }

        try
        {
            var options = ParseOptions(args.Skip(1).ToArray());
            if (IsHelpRequested(options))
            {
                PrintHelp();
                return 0;
            }

            if (string.IsNullOrWhiteSpace(options.InputPath) || string.IsNullOrWhiteSpace(options.OutputPath))
            {
                throw new ArgumentException("Input and output paths are required.");
            }

            var mode = ResolveMode(command, options.InputPath, options.OutputPath);
            var pages = ParsePages(options.Pages);
            await ConvertAsync(mode, options.InputPath, options.OutputPath, pages).ConfigureAwait(false);
            Console.WriteLine($"Converted {options.InputPath} -> {options.OutputPath}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static async Task ConvertAsync(ConversionMode mode, string inputPath, string outputPath, IReadOnlyList<int>? pages)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input file does not exist.", inputPath);
        }

        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }

        await using var input = File.OpenRead(inputPath);
        await using var output = File.Create(outputPath);

        if (mode == ConversionMode.PdfToOfd)
        {
            var converter = new PdfToOfdConverter();
            await converter.ConvertAsync(input, output, pages).ConfigureAwait(false);
            return;
        }

        var ofdToPdf = new OfdToPdfConverter();
        await ofdToPdf.ConvertAsync(input, output, pages).ConfigureAwait(false);
    }

    private static ConversionMode ResolveMode(string command, string inputPath, string outputPath)
    {
        return command switch
        {
            "pdf-to-ofd" => ConversionMode.PdfToOfd,
            "ofd-to-pdf" => ConversionMode.OfdToPdf,
            _ => InferMode(inputPath, outputPath)
        };
    }

    private static ConversionMode InferMode(string inputPath, string outputPath)
    {
        var inputExtension = Path.GetExtension(inputPath).ToLowerInvariant();
        var outputExtension = Path.GetExtension(outputPath).ToLowerInvariant();

        if (inputExtension == ".pdf" && outputExtension == ".ofd")
        {
            return ConversionMode.PdfToOfd;
        }

        if (inputExtension == ".ofd" && outputExtension == ".pdf")
        {
            return ConversionMode.OfdToPdf;
        }

        throw new ArgumentException("Unable to infer conversion direction. Use pdf-to-ofd or ofd-to-pdf.");
    }

    private static CliOptions ParseOptions(string[] args)
    {
        var options = new CliOptions();
        var positionals = new List<string>();

        for (var i = 0; i < args.Length; i++)
        {
            var arg = args[i];
            switch (arg)
            {
                case "-h":
                case "--help":
                    options.Help = true;
                    break;
                case "-i":
                case "--input":
                    options.InputPath = ReadValue(args, ref i, arg);
                    break;
                case "-o":
                case "--output":
                    options.OutputPath = ReadValue(args, ref i, arg);
                    break;
                case "-p":
                case "--pages":
                    options.Pages = ReadValue(args, ref i, arg);
                    break;
                default:
                    if (arg.StartsWith("-", StringComparison.Ordinal))
                    {
                        throw new ArgumentException($"Unknown option: {arg}");
                    }

                    positionals.Add(arg);
                    break;
            }
        }

        if (options.InputPath is null && positionals.Count > 0)
        {
            options.InputPath = positionals[0];
        }

        if (options.OutputPath is null && positionals.Count > 1)
        {
            options.OutputPath = positionals[1];
        }

        if (positionals.Count > 2)
        {
            throw new ArgumentException("Too many positional arguments.");
        }

        return options;
    }

    private static string ReadValue(string[] args, ref int index, string option)
    {
        if (index + 1 >= args.Length || args[index + 1].StartsWith("-", StringComparison.Ordinal))
        {
            throw new ArgumentException($"Missing value for {option}.");
        }

        index++;
        return args[index];
    }

    private static IReadOnlyList<int>? ParsePages(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var pages = new List<int>();
        foreach (var part in value.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var range = part.Split('-', StringSplitOptions.TrimEntries);
            if (range.Length == 1)
            {
                pages.Add(ParseOneBasedPage(range[0]));
                continue;
            }

            if (range.Length != 2)
            {
                throw new ArgumentException($"Invalid page range: {part}");
            }

            var start = ParseOneBasedPage(range[0]);
            var end = ParseOneBasedPage(range[1]);
            if (end < start)
            {
                throw new ArgumentException($"Invalid descending page range: {part}");
            }

            for (var page = start; page <= end; page++)
            {
                pages.Add(page);
            }
        }

        return pages;
    }

    private static int ParseOneBasedPage(string value)
    {
        if (!int.TryParse(value, out var page) || page <= 0)
        {
            throw new ArgumentException($"Invalid page number: {value}");
        }

        return page - 1;
    }

    private static bool IsHelp(string arg)
    {
        return arg is "-h" or "--help" or "help";
    }

    private static bool IsHelpRequested(CliOptions options)
    {
        return options.Help;
    }

    private static void PrintHelp()
    {
        Console.WriteLine("""
        Ofdrw.Net CLI

        Usage:
          ofdrw convert <input> <output> [--pages 1,3-5]
          ofdrw pdf-to-ofd --input <input.pdf> --output <output.ofd> [--pages 1,3-5]
          ofdrw ofd-to-pdf --input <input.ofd> --output <output.pdf> [--pages 1,3-5]

        Commands:
          convert     Infer conversion direction from .pdf/.ofd extensions.
          pdf-to-ofd  Convert PDF to OFD.
          ofd-to-pdf  Convert OFD to PDF.

        Options:
          -i, --input   Input file path.
          -o, --output  Output file path.
          -p, --pages   1-based page list or ranges, for example 1,3-5.
          -h, --help    Show help.
        """);
    }

    private sealed class CliOptions
    {
        public string? InputPath { get; set; }

        public string? OutputPath { get; set; }

        public string? Pages { get; set; }

        public bool Help { get; set; }
    }

    private enum ConversionMode
    {
        PdfToOfd,
        OfdToPdf
    }
}
