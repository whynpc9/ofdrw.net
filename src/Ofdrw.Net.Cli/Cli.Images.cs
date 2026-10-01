using System.Globalization;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;

internal static partial class Cli
{
    // A small command-local parser keeps multi-image positional syntax out of existing commands.
    private static async Task<int> ConvertImagesAsync(string command, string[] args, CancellationToken token)
    {
        if (args.Any(IsHelp)) { PrintHelp(); return 0; }
        var export = command == "ofd-to-image";
        var inputs = new List<string>();
        int? lastPositionalIndex = null;
        string? destination = null;
        var exportOptions = new OfdToImageOptions();
        var importOptions = new ImageToOfdOptions();
        double? pageWidth = null, pageHeight = null;
        var page = 0;
        for (var index = 0; index < args.Length; index++)
        {
            var option = args[index];
            if (!option.StartsWith('-')) { lastPositionalIndex = inputs.Count; inputs.Add(option); continue; }
            var value = ReadValue(args, ref index, option);
            switch (option)
            {
                case "--input": case "-i": inputs.Add(value); break;
                case "--output": case "-o": destination = value; break;
                case "--ppm": exportOptions.PixelsPerMillimeter = importOptions.PixelsPerMillimeter = ImageNumber(value, option); break;
                case "--max-pixels": exportOptions.MaxPixels = importOptions.MaxPixelsPerImage = ImageBytes(value, option); break;
                case "--max-working-bytes": exportOptions.MaxRasterWorkingBytes = importOptions.MaxRasterWorkingBytes = ImageBytes(value, option); break;
                case "--max-input-bytes": exportOptions.PackageLoadOptions.MaxInputBytes = importOptions.MaxInputBytesPerImage = ImageBytes(value, option); break;
                case "--max-output-bytes": exportOptions.MaxOutputBytes = importOptions.MaxOutputBytes = ImageBytes(value, option); break;
                case "--pages" or "-p" when export:
                    // A single numeric page avoids allocating enormous ranges just to reject multiple pages.
                    page = ParseOneBasedPage(value); break;
                case "--format" when export:
                    exportOptions.Format = value.ToLowerInvariant() switch
                    { "png" => OfdImageFormat.Png, "jpeg" or "jpg" => OfdImageFormat.Jpeg,
                        _ => throw new ArgumentException("--format must be png or jpeg.") }; break;
                case "--jpeg-quality" when export: exportOptions.JpegQuality = ImageCount(value, option); break;
                case "--max-pdf-bytes" when export: exportOptions.MaxIntermediatePdfBytes = ImageBytes(value, option); break;
                case "--page-width" when !export: pageWidth = ImageNumber(value, option); break;
                case "--page-height" when !export: pageHeight = ImageNumber(value, option); break;
                case "--max-total-input-bytes" when !export: importOptions.MaxTotalInputBytes = ImageBytes(value, option); break;
                case "--max-pages" when !export: importOptions.MaxPageCount = ImageCount(value, option); break;
                case "--max-entries" when !export: importOptions.MaxEntryCount = ImageCount(value, option); break;
                default: throw new ArgumentException($"Unknown option for {command}: {option}");
            }
        }
        if (destination is null && lastPositionalIndex.HasValue && inputs.Count >= 2)
        {
            var candidate = inputs[lastPositionalIndex.Value];
            if (!export && !string.Equals(Path.GetExtension(candidate), ".ofd", StringComparison.OrdinalIgnoreCase))
                throw new ArgumentException("image-to-ofd requires --output; an implicit positional output must end in .ofd.");
            destination = candidate;
            inputs.RemoveAt(lastPositionalIndex.Value);
        }
        if (inputs.Count == 0 || string.IsNullOrWhiteSpace(destination) || inputs.Any(string.IsNullOrWhiteSpace))
            throw new ArgumentException("Input and output paths are required.");
        if (export && inputs.Count != 1) throw new ArgumentException("ofd-to-image accepts one input OFD and one output image.");
        if (pageWidth.HasValue != pageHeight.HasValue) throw new ArgumentException("--page-width and --page-height must be provided together.");
        if (pageWidth.HasValue) importOptions.PageSize = new OfdPageSize { WidthMillimeters = pageWidth.Value, HeightMillimeters = pageHeight!.Value };
        // Validate options before creating even the staging output.
        var exporter = export ? new OfdToImageConverter(exportOptions) : null;
        var importer = export ? null : new ImageToOfdConverter(importOptions);
        if (!export && (inputs.Count > importOptions.MaxPageCount || 5L + inputs.Count * 2L > importOptions.MaxEntryCount))
            throw new ArgumentException("Input image count exceeds the page/entry budget.");
        foreach (var input in inputs) EnsureDifferentPaths(input, destination);
        var streams = new List<Stream>();
        try
        {
            foreach (var input in inputs) { token.ThrowIfCancellationRequested(); streams.Add(File.OpenRead(input)); }
            await using var output = new AtomicOutput(destination);
            if (export) await exporter!.ConvertAsync(streams[0], output.Stream, page, token).ConfigureAwait(false);
            else await importer!.ConvertAsync(streams, output.Stream, token).ConfigureAwait(false);
            await output.CommitAsync(token).ConfigureAwait(false);
        }
        finally { foreach (var stream in streams) stream.Dispose(); }
        Console.WriteLine($"Converted {inputs.Count} input(s) -> {destination}");
        return 0;
    }

    private static double ImageNumber(string value, string option)
    {
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) ||
            !double.IsFinite(number) || number <= 0) throw new ArgumentException($"{option} requires a finite positive number.");
        return number;
    }
    private static long ImageBytes(string value, string option)
    {
        if (!long.TryParse(value, NumberStyles.None, CultureInfo.InvariantCulture, out var number) || number <= 0)
            throw new ArgumentException($"{option} requires a positive integer.");
        return number;
    }
    private static int ImageCount(string value, string option)
    {
        var number = ImageBytes(value, option);
        if (number > int.MaxValue) throw new ArgumentException($"{option} exceeds the integer limit.");
        return (int)number;
    }
}
