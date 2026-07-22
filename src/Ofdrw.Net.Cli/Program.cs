using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Signatures.Verification;

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
        if (command is not ("convert" or "docx-to-pdf" or "docx-to-ofd" or
            "pdf-to-ofd" or "ofd-to-pdf" or "ofd-to-svg" or
            "extract-text" or "merge" or "reorder" or "verify-signatures"))
        {
            Console.Error.WriteLine($"Unknown command: {args[0]}");
            PrintHelp();
            return 1;
        }

        try
        {
            if (command == "merge")
            {
                return await MergeAsync(args.Skip(1).ToArray()).ConfigureAwait(false);
            }

            var options = ParseOptions(args.Skip(1).ToArray());
            if (IsHelpRequested(options))
            {
                PrintHelp();
                return 0;
            }

            if (command == "extract-text")
            {
                if (string.IsNullOrWhiteSpace(options.InputPath))
                {
                    throw new ArgumentException("An input OFD path is required.");
                }

                await ExtractTextAsync(
                    options.InputPath,
                    options.OutputPath,
                    options.IncludeTemplates).ConfigureAwait(false);
                return 0;
            }

            if (command == "verify-signatures")
            {
                if (string.IsNullOrWhiteSpace(options.InputPath))
                {
                    throw new ArgumentException("An input OFD path is required.");
                }

                return await VerifySignaturesAsync(options.InputPath)
                    .ConfigureAwait(false);
            }

            if (string.IsNullOrWhiteSpace(options.InputPath) || string.IsNullOrWhiteSpace(options.OutputPath))
            {
                throw new ArgumentException("Input and output paths are required.");
            }

            if (command == "reorder")
            {
                var pageOrder = ParsePages(options.Pages)
                    ?? throw new ArgumentException("--pages is required for reorder.");
                await ReorderAsync(
                    options.InputPath,
                    options.OutputPath,
                    pageOrder).ConfigureAwait(false);
                Console.WriteLine($"Reordered {options.InputPath} -> {options.OutputPath}");
                return 0;
            }

            if (command == "ofd-to-svg")
            {
                var svgPages = ParsePages(options.Pages);
                if (svgPages is { Count: > 1 })
                {
                    throw new ArgumentException("OFD to SVG accepts at most one page.");
                }

                await ConvertToSvgAsync(
                    options.InputPath,
                    options.OutputPath,
                    svgPages is { Count: 1 } ? svgPages[0] : 0).ConfigureAwait(false);
                Console.WriteLine($"Converted {options.InputPath} -> {options.OutputPath}");
                return 0;
            }

            var mode = ResolveMode(command, options.InputPath, options.OutputPath);
            var pages = ParsePages(options.Pages);
            await ConvertAsync(
                mode,
                options.InputPath,
                options.OutputPath,
                pages,
                options.DocxEngine,
                options.LibreOfficePath,
                options.FontDirectories).ConfigureAwait(false);
            Console.WriteLine($"Converted {options.InputPath} -> {options.OutputPath}");
            return 0;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }

    private static async Task ConvertAsync(
        ConversionMode mode,
        string inputPath,
        string outputPath,
        IReadOnlyList<int>? pages,
        DocxConversionEngine docxEngine,
        string? libreOfficePath,
        IReadOnlyList<string> fontDirectories)
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

        if (mode == ConversionMode.DocxToPdf)
        {
            if (pages is not null)
            {
                throw new ArgumentException("DOCX to PDF conversion does not support --pages.");
            }

            var converter = new DocxToPdfConverter(CreateDocxOptions(
                docxEngine,
                libreOfficePath,
                fontDirectories));
            await converter.ConvertAsync(input, output).ConfigureAwait(false);
            return;
        }

        if (mode == ConversionMode.DocxToOfd)
        {
            var converter = new DocxToOfdConverter(CreateDocxOptions(
                docxEngine,
                libreOfficePath,
                fontDirectories));
            await converter.ConvertAsync(input, output, pages).ConfigureAwait(false);
            return;
        }

        if (mode == ConversionMode.PdfToOfd)
        {
            var converter = new PdfToOfdConverter();
            await converter.ConvertAsync(input, output, pages).ConfigureAwait(false);
            return;
        }

        var ofdToPdf = new OfdToPdfConverter();
        await ofdToPdf.ConvertAsync(input, output, pages).ConfigureAwait(false);
    }

    private static DocxConversionOptions CreateDocxOptions(
        DocxConversionEngine engine,
        string? libreOfficePath,
        IEnumerable<string> fontDirectories)
    {
        var options = new DocxConversionOptions
        {
            Engine = engine,
            LibreOfficePath = libreOfficePath
        };
        foreach (var directory in fontDirectories)
        {
            options.FontDirectories.Add(directory);
        }

        return options;
    }

    private static async Task ExtractTextAsync(
        string inputPath,
        string? outputPath,
        bool includeTemplates)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input file does not exist.", inputPath);
        }

        await using var input = File.OpenRead(inputPath);
        var package = await new OfdReader().ReadAsync(input).ConfigureAwait(false);
        var text = new OfdTextExtractor().Extract(package, includeTemplates);
        if (string.IsNullOrWhiteSpace(outputPath))
        {
            Console.WriteLine(text);
            return;
        }

        EnsureOutputDirectory(outputPath);
        await File.WriteAllTextAsync(outputPath, text).ConfigureAwait(false);
        Console.WriteLine($"Extracted text {inputPath} -> {outputPath}");
    }

    private static async Task ConvertToSvgAsync(
        string inputPath,
        string outputPath,
        int pageIndex)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input file does not exist.", inputPath);
        }

        EnsureOutputDirectory(outputPath);
        await using var input = File.OpenRead(inputPath);
        await using var output = File.Create(outputPath);
        await new OfdToSvgConverter()
            .ConvertAsync(input, output, pageIndex)
            .ConfigureAwait(false);
    }

    private static async Task<int> VerifySignaturesAsync(string inputPath)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input file does not exist.", inputPath);
        }

        await using var input = File.OpenRead(inputPath);
        var report = await new OfdSignatureVerifier()
            .VerifyAsync(input)
            .ConfigureAwait(false);
        foreach (var issue in report.Issues)
        {
            Console.WriteLine($"Issue: {issue}");
        }

        foreach (var signature in report.Signatures)
        {
            var matchingReferences = signature.References.Count(reference =>
                reference.EntryExists &&
                reference.DigestAlgorithmSupported &&
                reference.DigestMatches);
            Console.WriteLine(
                $"Signature {signature.Id}: references " +
                $"{matchingReferences}/{signature.References.Count}, " +
                $"signed-value {signature.CryptographicStatus}, " +
                $"method {signature.SignatureMethod}");
            if (!string.IsNullOrWhiteSpace(signature.CryptographicMessage))
            {
                Console.WriteLine($"  {signature.CryptographicMessage}");
            }
        }

        if (report.FullyValid)
        {
            Console.WriteLine("Signature verification: fully valid.");
            return 0;
        }

        if (report.ReferenceIntegrityValid &&
            report.Signatures.All(signature =>
                signature.CryptographicStatus ==
                OfdCryptographicVerificationStatus.Unsupported))
        {
            Console.WriteLine(
                "Signature verification: reference integrity valid; " +
                "signed-value algorithm requires a registered verifier.");
            return 2;
        }

        Console.WriteLine("Signature verification: invalid or incomplete.");
        return 3;
    }

    private static async Task ReorderAsync(
        string inputPath,
        string outputPath,
        IReadOnlyList<int> pageOrder)
    {
        if (!File.Exists(inputPath))
        {
            throw new FileNotFoundException("Input file does not exist.", inputPath);
        }

        await using var input = File.OpenRead(inputPath);
        var package = await new OfdReader().ReadAsync(input).ConfigureAwait(false);
        OfdDocumentEditor.ReorderPages(package, pageOrder);
        EnsureOutputDirectory(outputPath);
        await using var output = File.Create(outputPath);
        await new OfdPackageWriter().WriteAsync(package, output).ConfigureAwait(false);
    }

    private static async Task<int> MergeAsync(string[] args)
    {
        if (args.Length == 0 || args.Any(IsHelp))
        {
            PrintHelp();
            return 0;
        }

        var skipUnsupported = args.Contains("--skip-unsupported", StringComparer.Ordinal);
        var paths = args.Where(arg => !string.Equals(
            arg,
            "--skip-unsupported",
            StringComparison.Ordinal)).ToArray();
        if (paths.Length < 3)
        {
            throw new ArgumentException(
                "Merge requires an output path followed by at least two input OFD files.");
        }

        var outputPath = paths[0];
        var sources = new List<Ofdrw.Net.Core.Models.OfdDocumentPackage>();
        foreach (var inputPath in paths.Skip(1))
        {
            if (!File.Exists(inputPath))
            {
                throw new FileNotFoundException("Input file does not exist.", inputPath);
            }

            await using var input = File.OpenRead(inputPath);
            sources.Add(await new OfdReader().ReadAsync(input).ConfigureAwait(false));
        }

        var merged = OfdDocumentMerger.Merge(
            sources,
            new OfdDocumentMergeOptions
            {
                SkipUnsupportedRawElements = skipUnsupported
            });
        EnsureOutputDirectory(outputPath);
        await using var output = File.Create(outputPath);
        await new OfdPackageWriter().WriteAsync(merged, output).ConfigureAwait(false);
        Console.WriteLine($"Merged {sources.Count} OFD files -> {outputPath}");
        return 0;
    }

    private static void EnsureOutputDirectory(string outputPath)
    {
        var outputDirectory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
        if (!string.IsNullOrWhiteSpace(outputDirectory))
        {
            Directory.CreateDirectory(outputDirectory);
        }
    }

    private static ConversionMode ResolveMode(string command, string inputPath, string outputPath)
    {
        return command switch
        {
            "docx-to-pdf" => ConversionMode.DocxToPdf,
            "docx-to-ofd" => ConversionMode.DocxToOfd,
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

        if (inputExtension == ".docx" && outputExtension == ".pdf")
        {
            return ConversionMode.DocxToPdf;
        }

        if (inputExtension == ".docx" && outputExtension == ".ofd")
        {
            return ConversionMode.DocxToOfd;
        }

        throw new ArgumentException(
            "Unable to infer conversion direction. Use docx-to-pdf, docx-to-ofd, pdf-to-ofd, or ofd-to-pdf.");
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
                case "--include-templates":
                    options.IncludeTemplates = true;
                    break;
                case "--libreoffice":
                    options.LibreOfficePath = ReadValue(args, ref i, arg);
                    break;
                case "--docx-engine":
                    options.DocxEngine = ParseDocxEngine(ReadValue(args, ref i, arg));
                    break;
                case "--font-directory":
                    options.FontDirectories.Add(ReadValue(args, ref i, arg));
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

    private static DocxConversionEngine ParseDocxEngine(string value)
    {
        return value.Trim().ToLowerInvariant() switch
        {
            "auto" => DocxConversionEngine.Auto,
            "word" or "microsoft-word" => DocxConversionEngine.MicrosoftWord,
            "libreoffice" => DocxConversionEngine.LibreOffice,
            _ => throw new ArgumentException(
                "Invalid --docx-engine value. Use auto, word, or libreoffice.")
        };
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
          ofdrw docx-to-pdf --input <input.docx> --output <output.pdf> [--docx-engine auto|word|libreoffice]
          ofdrw docx-to-ofd --input <input.docx> --output <output.ofd> [--pages 1,3-5] [--docx-engine auto|word|libreoffice]
          ofdrw pdf-to-ofd --input <input.pdf> --output <output.ofd> [--pages 1,3-5]
          ofdrw ofd-to-pdf --input <input.ofd> --output <output.pdf> [--pages 1,3-5]
          ofdrw ofd-to-svg --input <input.ofd> --output <output.svg> [--pages 1]
          ofdrw verify-signatures --input <input.ofd>
          ofdrw extract-text <input.ofd> [output.txt] [--include-templates]
          ofdrw reorder <input.ofd> <output.ofd> --pages 3,1,2
          ofdrw merge <output.ofd> <input1.ofd> <input2.ofd> [...] [--skip-unsupported]

        Commands:
          convert     Infer conversion direction from .docx/.pdf/.ofd extensions.
          docx-to-pdf Convert DOCX to PDF using Word on macOS when available, or LibreOffice.
          docx-to-ofd Convert DOCX to OFD through the PDF rendering pipeline.
          pdf-to-ofd  Convert PDF to OFD.
          ofd-to-pdf  Convert OFD to PDF.
          ofd-to-svg  Convert one OFD page to self-contained SVG.
          verify-signatures Verify protected-entry digests and registered signed-value algorithms.
          extract-text Extract page text, optionally including template text.
          reorder      Reorder every page using a complete 1-based page list.
          merge        Merge OFD pages into a self-contained output document.

        Options:
          -i, --input   Input file path.
          -o, --output  Output file path.
          -p, --pages   1-based page list or ranges, for example 1,3-5.
          --include-templates Include template text during extraction.
          --skip-unsupported  Drop unsupported raw objects during merge.
          --docx-engine      DOCX renderer: auto, word (macOS), or libreoffice.
          --libreoffice       Path to the LibreOffice soffice executable for DOCX conversion.
          --font-directory    Additional font directory for DOCX rendering; may be repeated.
          -h, --help    Show help.
        """);
    }

    private sealed class CliOptions
    {
        public string? InputPath { get; set; }

        public string? OutputPath { get; set; }

        public string? Pages { get; set; }

        public bool IncludeTemplates { get; set; }

        public string? LibreOfficePath { get; set; }

        public DocxConversionEngine DocxEngine { get; set; } = DocxConversionEngine.Auto;

        public List<string> FontDirectories { get; } = new();

        public bool Help { get; set; }
    }

    private enum ConversionMode
    {
        DocxToPdf,
        DocxToOfd,
        PdfToOfd,
        OfdToPdf
    }
}
