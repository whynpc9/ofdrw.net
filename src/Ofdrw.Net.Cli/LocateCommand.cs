using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using Ofdrw.Net.Converter.Pdf.Geometry;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Readers;

internal static partial class Cli
{
    private const string LocateSchema = "ofdrw-cli-locate/1";

    // ofdrw locate <input.ofd> --find TEXT [--trusted-font SHA256:BYTECOUNT ...] [--pages 1,3-5] [--document N] [--json]
    // Read-only preview for replace-text. Exit: 0 complete with hits, 2 complete without hits, 3 incomplete.
    private static async Task<int> LocateAsync(string[] args, CancellationToken token)
    {
        string? input = null, keyword = null, pages = null, document = null;
        var json = false;
        var trusted = new List<OfdTrustedEmbeddedFontIdentity>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        for (var index = 0; index < args.Length; index++)
        {
            var arg = args[index];
            if (arg == "--json")
            {
                if (json) throw new ArgumentException("Duplicate locate option: --json");
                json = true;
                continue;
            }
            if (!arg.StartsWith("-", StringComparison.Ordinal))
            {
                if (input is not null) throw new ArgumentException("locate takes exactly one input OFD.");
                input = arg;
                continue;
            }
            if (arg is not ("--find" or "--trusted-font" or "--pages" or "--document"))
                throw new ArgumentException("Unknown locate option: " + arg);
            if (arg != "--trusted-font" && !seen.Add(arg)) throw new ArgumentException("Duplicate locate option: " + arg);
            if (index + 1 >= args.Length) throw new ArgumentException(arg + " requires a value.");
            var value = args[++index];
            switch (arg)
            {
                case "--find": keyword = value; break;
                case "--pages": pages = value; break;
                case "--document": document = value; break;
                default: trusted.Add(TextReplacementArguments.ParseTrustedFont(value, trusted.Count)); break;
            }
        }

        if (string.IsNullOrWhiteSpace(input)) throw new ArgumentException("An input OFD path is required.");
        if (string.IsNullOrEmpty(keyword)) throw new ArgumentException("A nonempty --find value is required.");
        if (!File.Exists(input)) throw new FileNotFoundException("Input file does not exist.", input);
        var selection = pages is null ? null : TextReplacementArguments.ParsePages(pages);
        var documentIndex = document is null ? 0 : ParseOneBasedPage(document);

        Ofdrw.Net.Core.Models.OfdDocumentPackage package;
        await using (var stream = File.OpenRead(input))
            package = await new OfdReader().ReadAsync(stream, new OfdPackageLoadOptions { DocumentIndex = documentIndex }, token).ConfigureAwait(false);
        // An empty identity list trusts nothing; the locator then reports what it can.
        var locator = new OfdKeywordLocator(new OfdTrustedEmbeddedFontGeometryProvider(trusted));
        var result = selection is null
            ? locator.Find(package, keyword, cancellationToken: token)
            : locator.Find(package, selection, keyword, cancellationToken: token);

        if (json) WriteLocateJson(result);
        else
        {
            foreach (var hit in result.Hits)
            {
                var box = hit.BoundsMillimeters is { } b
                    ? string.Join(",", new[] { b.Left, b.Top, b.Right, b.Bottom }.Select(v => v.ToString("0.000", CultureInfo.InvariantCulture)))
                    : "-";
                Console.WriteLine($"page {hit.PageIndex + 1} \"{hit.MatchedText}\" box {box} replaceable={(hit.CanReplace ? "true" : "false")}");
            }
            foreach (var diagnostic in result.Diagnostics)
                Console.WriteLine($"diagnostic {diagnostic.Code} page {(diagnostic.PageIndex is int page ? (page + 1).ToString(CultureInfo.InvariantCulture) : "-")} {diagnostic.SourcePath ?? "-"}");
            Console.WriteLine($"Complete: {(result.IsComplete ? "true" : "false")}");
        }

        if (!result.IsComplete) return 3;
        return result.Hits.Count > 0 ? 0 : 2;
    }

    private static void WriteLocateJson(OfdKeywordSearchResult result)
    {
        var hits = new JsonArray();
        foreach (var hit in result.Hits)
            hits.Add(new JsonObject
            {
                ["page"] = hit.PageIndex + 1,
                ["text"] = hit.MatchedText,
                ["bounds"] = hit.BoundsMillimeters is { } b
                    ? new JsonObject { ["left"] = b.Left, ["top"] = b.Top, ["right"] = b.Right, ["bottom"] = b.Bottom }
                    : null,
                ["canReplace"] = hit.CanReplace
            });
        var diagnostics = new JsonArray();
        foreach (var diagnostic in result.Diagnostics)
            diagnostics.Add(new JsonObject
            {
                ["code"] = diagnostic.Code,
                ["page"] = diagnostic.PageIndex is int page ? page + 1 : null,
                ["sourcePath"] = diagnostic.SourcePath,
                ["makesScopeIncomplete"] = diagnostic.MakesScopeIncomplete
            });
        var document = new JsonObject
        {
            ["schema"] = LocateSchema, ["complete"] = result.IsComplete, ["hits"] = hits, ["diagnostics"] = diagnostics
        };
        Console.Out.Write(document.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) + Environment.NewLine);
    }
}
