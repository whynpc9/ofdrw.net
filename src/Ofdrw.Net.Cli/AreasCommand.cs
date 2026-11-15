using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Ofdrw.Net.Layout.Editing;
using SixLabors.ImageSharp;

internal static partial class Cli
{
    private const string AreasSchema = "ofdrw-cli-areas/1";

    private static async Task<int> AreasAsync(string[] args, CancellationToken token)
    {
        if (args.Length == 0) throw new ArgumentException("areas requires list or fill.");
        var subcommand = args[0];
        if (subcommand is not ("list" or "fill"))
            throw new ArgumentException($"Unknown areas subcommand: {subcommand}");

        var positionals = new List<string>();
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        var list = subcommand == "list";
        var json = false;
        var allowed = list
            ? new[] { "--document", "--json" }
            : new[] { "--document", "--name", "--text", "--image", "--font", "--font-size" };
        for (var index = 1; index < args.Length; index++)
        {
            var arg = args[index];
            if (!arg.StartsWith("-", StringComparison.Ordinal)) { positionals.Add(arg); continue; }
            if (Array.IndexOf(allowed, arg) < 0) throw new ArgumentException($"Unknown areas {subcommand} option: {arg}");
            if (arg == "--json")
            {
                if (json) throw new ArgumentException($"Duplicate option: {arg}");
                json = true;
                continue;
            }
            if (values.ContainsKey(arg)) throw new ArgumentException($"Duplicate option: {arg}");
            if (index + 1 >= args.Length) throw new ArgumentException($"{arg} requires a value.");
            values[arg] = args[++index];
        }

        if (positionals.Count != (list ? 1 : 2))
            throw new ArgumentException(list
                ? "areas list takes one input OFD path."
                : "areas fill takes an input and an output OFD path.");
        var input = positionals[0];
        var documentIndex = values.TryGetValue("--document", out var document) ? ParseOneBasedPage(document) : 0;
        if (list)
        {
            var package = await ReadToolInputAsync(input, token, documentIndex, allowMultiple: true).ConfigureAwait(false);
            ListAreas(new OfdNamedAreaEditor().GetAreas(package), json);
            return 0;
        }

        await FillAreaAsync(input, positionals[1], values, documentIndex, token).ConfigureAwait(false);
        return 0;
    }

    private static void ListAreas(IReadOnlyList<OfdNamedArea> areas, bool json)
    {
        if (!json)
        {
            foreach (var area in areas)
            {
                Console.WriteLine(string.Join("\t", TerminalText.Escape(area.Name),
                    (area.PageIndex + 1).ToString(CultureInfo.InvariantCulture),
                    FormatMillimeters(area.XMillimeters), FormatMillimeters(area.YMillimeters),
                    FormatMillimeters(area.WidthMillimeters), FormatMillimeters(area.HeightMillimeters),
                    area.IsFilled ? "filled" : "blank"));
            }
            return;
        }

        var items = new JsonArray();
        foreach (var area in areas)
        {
            items.Add(new JsonObject
            {
                ["name"] = area.Name,
                ["page"] = area.PageIndex + 1,
                ["xMillimeters"] = area.XMillimeters,
                ["yMillimeters"] = area.YMillimeters,
                ["widthMillimeters"] = area.WidthMillimeters,
                ["heightMillimeters"] = area.HeightMillimeters,
                ["filled"] = area.IsFilled
            });
        }
        var document = new JsonObject { ["schema"] = AreasSchema, ["areas"] = items };
        Console.WriteLine(document.ToJsonString(new JsonSerializerOptions { WriteIndented = true }));
    }

    private static string FormatMillimeters(double value) => value.ToString("0.###", CultureInfo.InvariantCulture);

    private static async Task FillAreaAsync(string input, string output, Dictionary<string, string> values,
        int documentIndex, CancellationToken token)
    {
        var name = values.GetValueOrDefault("--name");
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("--name is required.");
        var text = values.GetValueOrDefault("--text");
        var imagePath = values.GetValueOrDefault("--image");
        if ((text is null) == (imagePath is null))
            throw new ArgumentException("areas fill requires exactly one of --text or --image.");
        var font = values.GetValueOrDefault("--font");
        var fontSize = values.GetValueOrDefault("--font-size");
        if (imagePath is not null && (font is not null || fontSize is not null))
            throw new ArgumentException("--font and --font-size apply only to --text.");

        var options = new NamedAreaTextOptions();
        if (font is not null) options.FontFamily = font;
        if (fontSize is not null)
        {
            if (!double.TryParse(fontSize, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) ||
                !double.IsFinite(size) || size <= 0)
                throw new ArgumentException("Invalid --font-size value.");
            options.FontSizeMillimeters = size;
        }

        var package = await ReadToolInputAsync(input, token, documentIndex).ConfigureAwait(false);
        var editor = new OfdNamedAreaEditor();
        if (imagePath is null)
        {
            editor.FillText(package, name, text!, options, token);
        }
        else
        {
            var mediaType = Path.GetExtension(imagePath).ToLowerInvariant() switch
            {
                ".png" => "image/png",
                ".jpg" or ".jpeg" => "image/jpeg",
                _ => throw new ArgumentException("Image must be PNG/JPEG.")
            };
            ConversionPathIdentity.RequireDifferentPaths(imagePath, output);
            await using var image = File.OpenRead(imagePath);
            var length = image.Length;
            if (length > 16 * 1024 * 1024) throw new ArgumentException("Image encoded byte budget exceeded.");
            var bytes = new byte[(int)length];
            await image.ReadExactlyAsync(bytes, token).ConfigureAwait(false);
            // Restrict the format before invoking an image decoder's identification method.
            if (Image.DetectFormat(bytes)?.DefaultMimeType != mediaType)
                throw new InvalidDataException("Image must be PNG/JPEG.");
            var info = Image.Identify(bytes) ?? throw new InvalidDataException("Image must be PNG/JPEG.");
            editor.FillImage(package, name, bytes, mediaType, info.Width, info.Height, token);
        }

        await WriteToolOutputAsync(package, output, token).ConfigureAwait(false);
        WriteSuccessNotification($"Filled area {TerminalText.Escape(name)} -> {output}");
    }
}
