using System.Globalization;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

internal static partial class Cli
{
    private static async Task<int> DocumentToolsAsync(string command, string[] args, CancellationToken token)
    {
        if (args.Length == 0 || args.Any(IsHelp))
        {
            PrintHelp();
            Console.WriteLine("Watermark options: --pages 1,2 --text TEXT | --image FILE --layer ID --layer-type Background|Body|Foreground --x MM --y MM --width MM --height MM --font NAME --font-size MM --alpha 0..255");
            return 0;
        }
        if (command == "mix")
        {
            if (args.Length < 3 || args.Length % 2 != 1) throw new ArgumentException("Mix requires output followed by one or more input.ofd / 1-based page pairs.");
            if ((args.Length - 1) / 2 > 1_000) throw new ArgumentException("Mix source count exceeds 1000.");
            var sources = new List<OfdMixSource>();
            long bytes = 0;
            for (var i = 1; i < args.Length; i += 2)
            {
                var package = await ReadToolInputAsync(args[i], token);
                bytes = checked(bytes + package.PreservedEntries.Values.Sum(data => (long)data.Length));
                if (bytes > 512L * 1024 * 1024) throw new ArgumentException("Mix cumulative expanded input budget exceeded.");
                if (!int.TryParse(args[i + 1], NumberStyles.None, CultureInfo.InvariantCulture, out var number) || number <= 0)
                    throw new ArgumentException("Mix page numbers must be positive 1-based integers.");
                sources.Add(new OfdMixSource(package, number - 1));
            }
            var mixed = OfdDocumentMixer.Mix(sources, cancellationToken: token);
            await WriteToolOutputAsync(mixed, args[0], token);
            return 0;
        }
        if (args.Length < 2) throw new ArgumentException("Input and output OFD paths are required.");
        if (command == "clean-signatures")
        {
            if (args.Length != 2) throw new ArgumentException("Clean-signatures takes only input and output paths.");
            await using var output = new AtomicOutput(args[1]);
            OfdPackageWriteResult result;
            // Close input before publishing to support atomic in-place cleanup.
            await using (var input = File.OpenRead(args[0]))
                result = await OfdSignatureCleaner.CleanAsync(input, output.Stream, cancellationToken: token);
            foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine($"Warning: {diagnostic}");
            await output.CommitAsync(token);
            return 0;
        }
        var options = new Dictionary<string, string>(StringComparer.Ordinal);
        var allowed = command == "split" ? new[] { "--pages" } : new[] { "--pages", "--text", "--image", "--layer", "--layer-type", "--x", "--y", "--width", "--height", "--font", "--font-size", "--alpha" };
        for (var i = 2; i < args.Length; i += 2)
        {
            if (!allowed.Contains(args[i]) || i + 1 >= args.Length || !options.TryAdd(args[i], args[i + 1])) throw new ArgumentException("Unknown, duplicate or incomplete document-tool option.");
        }
        var source = await ReadToolInputAsync(args[0], token);
        var pages = ParsePages(options.GetValueOrDefault("--pages"));
        if (command == "split")
        {
            if (pages is null) throw new ArgumentException("Split requires --pages.");
            await WriteToolOutputAsync(OfdDocumentSplitter.Split(source, pages, token), args[1], token);
            return 0;
        }
        if (options.ContainsKey("--text") == options.ContainsKey("--image")) throw new ArgumentException("Watermark requires exactly one of --text or --image.");
        double Number(string key, double fallback) => options.TryGetValue(key, out var value)
            ? double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture) : fallback;
        var placement = new OfdWatermarkOptions
        {
            LayerId = options.GetValueOrDefault("--layer"), LayerType = options.GetValueOrDefault("--layer-type", "Foreground"),
            XMillimeters = Number("--x", 20), YMillimeters = Number("--y", 30),
            WidthMillimeters = Number("--width", 100), HeightMillimeters = Number("--height", 20)
        };
        pages ??= Enumerable.Range(0, source.Pages.Count).ToArray();
        var alpha = options.TryGetValue("--alpha", out var alphaText) ? int.Parse(alphaText, CultureInfo.InvariantCulture) : 160;
        if (alpha < 0 || alpha > 255) throw new ArgumentOutOfRangeException("--alpha");
        if (options.TryGetValue("--text", out var text))
            OfdWatermark.AddText(source, pages, text, placement, options.GetValueOrDefault("--font", "SimSun"), Number("--font-size", 8), new OfdColor(160, 80, 80, alpha), token);
        else
        {
            var path = options["--image"];
            if (new FileInfo(path).Length > placement.MaxImageBytes) throw new ArgumentException("Image encoded byte budget exceeded.");
            var mediaType = Path.GetExtension(path).ToLowerInvariant() switch { ".png" => "image/png", ".jpg" or ".jpeg" => "image/jpeg", _ => throw new ArgumentException("Image must be PNG/JPEG.") };
            OfdWatermark.AddImage(source, pages, await File.ReadAllBytesAsync(path, token), mediaType, placement, alpha, token);
        }
        await WriteToolOutputAsync(source, args[1], token);
        return 0;
    }

    private static async Task<OfdDocumentPackage> ReadToolInputAsync(string path, CancellationToken token)
    {
        await using var input = File.OpenRead(path);
        var package = await new OfdReader().ReadAsync(input, token);
        using var root = new MemoryStream(package.PreservedEntries["OFD.xml"], false);
        if (XDocument.Load(root).Root!.Elements().Count(node => node.Name.LocalName == "DocBody") != 1)
            throw new NotSupportedException("Page editing requires a single DocBody; clean-signatures supports all documents.");
        return package;
    }

    private static async Task WriteToolOutputAsync(OfdDocumentPackage package, string path, CancellationToken token)
    {
        await using var output = new AtomicOutput(path);
        var result = await new OfdPackageWriter().WriteWithResultAsync(package, output.Stream, token);
        foreach (var diagnostic in result.Diagnostics) Console.Error.WriteLine($"Warning: {diagnostic}");
        await output.CommitAsync(token);
    }
}
