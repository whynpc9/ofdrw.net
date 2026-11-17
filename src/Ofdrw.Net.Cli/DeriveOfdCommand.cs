using System.Text.Json;
using System.Text.Json.Nodes;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

internal static partial class Cli
{
    private const string DeriveOfdSchema = "ofdrw-cli-derive-ofd/1";

    /// <summary>Exit code for a content-changing command that refused an input declaring DocType="OFD-A".</summary>
    internal const int ArchivalOriginalExitCode = 5;

    // ofdrw derive-ofd <input.ofd> <output.ofd> [--json]
    private static async Task<int> DeriveOfdAsync(string[] args, CancellationToken token)
    {
        var json = false;
        var paths = new List<string>();
        foreach (var arg in args)
        {
            if (arg is "-h" or "--help") { PrintHelp(); return 0; }
            if (arg == "--json")
            {
                if (json) throw new ArgumentException("Duplicate derive-ofd option: --json");
                json = true;
            }
            else if (arg.StartsWith("-", StringComparison.Ordinal)) throw new ArgumentException($"Unknown derive-ofd option: {arg}");
            else paths.Add(arg);
        }
        if (paths.Count != 2) throw new ArgumentException("derive-ofd takes an input OFD-A path and a new output OFD path.");
        if (!File.Exists(paths[0])) throw new FileNotFoundException("Input file does not exist.", paths[0]);
        var result = await OfdOrdinaryDerivation.CreateAsync(paths[0], paths[1], cancellationToken: token).ConfigureAwait(false);
        if (json)
        {
            var document = new JsonObject
            {
                ["schema"] = DeriveOfdSchema,
                ["operation"] = result.Operation,
                ["source"] = new JsonObject
                {
                    ["path"] = result.SourcePath, ["bytes"] = result.SourceLength,
                    ["sha256"] = result.SourceSha256, ["docType"] = result.SourceDocType
                },
                ["output"] = new JsonObject
                {
                    ["path"] = result.OutputPath, ["bytes"] = result.OutputLength,
                    ["sha256"] = result.OutputSha256, ["docType"] = result.OutputDocType
                },
                ["changedEntry"] = result.ChangedEntry,
                ["inheritsArchivalConclusion"] = result.InheritsArchivalConclusion
            };
            try { Console.Out.Write(document.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) + Environment.NewLine); } catch { }
            return 0;
        }
        WriteSuccessNotification(string.Join(Environment.NewLine,
            $"Derived ordinary OFD {TerminalText.Escape(result.OutputPath)}",
            $"operation\t{result.Operation}",
            $"source\t{TerminalText.Escape(result.SourcePath)}\t{result.SourceLength} bytes\tsha256 {result.SourceSha256}\tDocType {TerminalText.Escape(result.SourceDocType)}",
            $"output\t{TerminalText.Escape(result.OutputPath)}\t{result.OutputLength} bytes\tsha256 {result.OutputSha256}\tDocType {result.OutputDocType}",
            $"changed\t{result.ChangedEntry} (root DocType only)",
            "The derivative does not inherit any archival conclusion or signature validity of the source."));
        return 0;
    }

    // Model-level attachment removal has no operation context of its own; the CLI verb is the user entry point.
    private static void RejectArchivalForAttachmentRemoval(OfdDocumentPackage package)
    {
        if (OfdDocTypes.IsArchival(package.Options.DocType)) throw new OfdArchivalOriginalException("attachments remove");
    }
}
