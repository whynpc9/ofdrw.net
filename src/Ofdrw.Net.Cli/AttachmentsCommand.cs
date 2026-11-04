using System.Text.Json;
using System.Text.Json.Nodes;
using Ofdrw.Net.Core.Models;

internal static partial class Cli
{
    private const string AttachmentsSchema = "ofdrw-cli-attachments/1";

    // ofdrw attachments list <input.ofd> [--document N] [--json]
    // ofdrw attachments extract <input.ofd> --id ID --output <file> [--document N]
    // ofdrw attachments remove <input.ofd> <output.ofd> --id ID [--physical] [--document N]
    private static async Task<int> AttachmentsAsync(string[] args, CancellationToken token)
    {
        if (args.Length == 0) throw new ArgumentException("attachments requires list, extract or remove.");
        var subcommand = args[0];
        if (subcommand is not ("list" or "extract" or "remove"))
            throw new ArgumentException($"Unknown attachments subcommand: {subcommand}");

        var positionals = new List<string>();
        var values = new Dictionary<string, string>(StringComparer.Ordinal);
        var flags = new HashSet<string>(StringComparer.Ordinal);
        var allowed = subcommand switch
        {
            "list" => new[] { "--document", "--json" },
            "extract" => new[] { "--document", "--id", "--output" },
            _ => new[] { "--document", "--id", "--physical" }
        };
        for (var index = 1; index < args.Length; index++)
        {
            var arg = args[index];
            if (!arg.StartsWith("-", StringComparison.Ordinal)) { positionals.Add(arg); continue; }
            if (Array.IndexOf(allowed, arg) < 0) throw new ArgumentException($"Unknown attachments {subcommand} option: {arg}");
            if (arg is "--json" or "--physical")
            {
                if (!flags.Add(arg)) throw new ArgumentException($"Duplicate option: {arg}");
                continue;
            }
            if (values.ContainsKey(arg)) throw new ArgumentException($"Duplicate option: {arg}");
            if (index + 1 >= args.Length) throw new ArgumentException($"{arg} requires a value.");
            values[arg] = args[++index];
        }

        var expectedPositionals = subcommand == "remove" ? 2 : 1;
        if (positionals.Count != expectedPositionals)
            throw new ArgumentException(subcommand == "remove"
                ? "attachments remove takes an input and an output OFD path."
                : $"attachments {subcommand} takes one input OFD path.");
        var input = positionals[0];
        if (!File.Exists(input)) throw new FileNotFoundException("Input file does not exist.", input);
        var documentIndex = values.TryGetValue("--document", out var document) ? ParseOneBasedPage(document) : 0;

        if (subcommand == "list")
        {
            // Read-only: multi-DocBody inputs are fine once --document selects one body.
            var package = await ReadToolInputAsync(input, token, documentIndex, allowMultiple: true).ConfigureAwait(false);
            ListAttachments(package, flags.Contains("--json"));
            return 0;
        }

        var id = values.TryGetValue("--id", out var requested) && requested.Length > 0
            ? requested : throw new ArgumentException("--id is required.");
        if (subcommand == "extract")
        {
            // Never derive the destination from the stored attachment name.
            var outputPath = values.TryGetValue("--output", out var output) && output.Length > 0
                ? output : throw new ArgumentException("--output is required for attachments extract.");
            var package = await ReadToolInputAsync(input, token, documentIndex, allowMultiple: true).ConfigureAwait(false);
            var attachment = SelectAttachment(package, id);
            if (attachment.IsExternal) throw new InvalidOperationException("External attachments have no embedded payload.");
            await using var file = new AtomicOutput(outputPath);
            await file.Stream.WriteAsync(attachment.Data, token).ConfigureAwait(false);
            await file.CommitAsync(token).ConfigureAwait(false);
            WriteSuccessNotification($"Extracted attachment {id} ({attachment.Data.Length} bytes) -> {outputPath}");
            return 0;
        }

        var target = positionals[1];
        var mode = flags.Contains("--physical") ? OfdAttachmentRemovalMode.Physical : OfdAttachmentRemovalMode.Logical;
        // Writing keeps the page-editing restriction to a single DocBody.
        var source = await ReadToolInputAsync(input, token, documentIndex).ConfigureAwait(false);
        source.RemoveAttachment(SelectAttachment(source, id), mode);
        await WriteToolOutputAsync(source, target, token).ConfigureAwait(false);
        WriteSuccessNotification($"Removed attachment {id} ({mode})");
        return 0;
    }

    private static OfdAttachment SelectAttachment(OfdDocumentPackage package, string id)
    {
        var matches = package.Attachments.Where(attachment => string.Equals(attachment.Id, id, StringComparison.Ordinal)).ToArray();
        return matches.Length switch
        {
            0 => throw new InvalidOperationException($"No attachment has ID '{id}'."),
            1 => matches[0],
            _ => throw new InvalidOperationException($"Attachment ID '{id}' is ambiguous.")
        };
    }

    private static void ListAttachments(OfdDocumentPackage package, bool json)
    {
        if (!json)
        {
            foreach (var attachment in package.Attachments)
                Console.WriteLine(string.Join("\t", attachment.Id ?? string.Empty, attachment.Name, attachment.MediaType,
                    attachment.IsExternal ? "" : attachment.Data.Length.ToString(System.Globalization.CultureInfo.InvariantCulture),
                    attachment.IsExternal ? "external" : "internal"));
            return;
        }

        var items = new JsonArray();
        foreach (var attachment in package.Attachments)
            items.Add(new JsonObject
            {
                ["id"] = attachment.Id,
                ["name"] = attachment.Name,
                ["mediaType"] = attachment.MediaType,
                ["bytes"] = attachment.IsExternal ? null : attachment.Data.Length,
                ["external"] = attachment.IsExternal,
                ["externalPath"] = attachment.ExternalPath
            });
        var document = new JsonObject { ["schema"] = AttachmentsSchema, ["attachments"] = items };
        Console.Out.Write(document.ToJsonString(new JsonSerializerOptions { WriteIndented = true }) + Environment.NewLine);
    }
}
