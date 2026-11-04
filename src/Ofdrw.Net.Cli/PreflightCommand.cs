using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using Ofdrw.Net.Packaging.Preflight;

internal static partial class Cli
{
    private const string PreflightSchema = "ofdrw-cli-preflight/1";

    private static readonly JsonSerializerOptions PreflightJson = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        Converters = { new JsonStringEnumConverter() }
    };

    // ofdrw preflight <input.ofd> [--json] [--output <report.json>] [--allow-multiple-documents] [--allow-encryption]
    // Exit: 0 complete and Pass/NotApplicable, 2 Fail, 3 Unknown or incomplete, 1 usage/IO, 130 cancelled.
    internal static Task<int> PreflightAsync(string[] args, CancellationToken token) =>
        PreflightAsync(args, new ArchivalPreflightOptions(), token);

    internal static async Task<int> PreflightAsync(string[] args, ArchivalPreflightOptions options, CancellationToken token)
    {
        string? input = null, outputPath = null;
        bool json = false, outputSeen = false;
        for (var index = 0; index < args.Length; index++)
        {
            var arg = args[index];
            switch (arg)
            {
                case "--json":
                    if (json) throw new ArgumentException("Duplicate preflight option: --json");
                    json = true;
                    break;
                case "--output":
                case "-o":
                    if (outputSeen) throw new ArgumentException("Duplicate preflight option: --output");
                    if (index + 1 >= args.Length || string.IsNullOrWhiteSpace(args[index + 1]))
                        throw new ArgumentException("--output requires a report path.");
                    outputSeen = true;
                    outputPath = args[++index];
                    break;
                case "--allow-multiple-documents":
                    options.RequireSingleDocument = false;
                    break;
                case "--allow-encryption":
                    options.RequireNoEncryption = false;
                    break;
                default:
                    if (arg.StartsWith("-", StringComparison.Ordinal))
                        throw new ArgumentException($"Unknown preflight option: {arg}");
                    if (input is not null) throw new ArgumentException("Preflight takes exactly one input OFD.");
                    input = arg;
                    break;
            }
        }

        if (string.IsNullOrWhiteSpace(input)) throw new ArgumentException("An input OFD path is required.");
        if (outputPath is not null && !json) throw new ArgumentException("--output requires --json.");
        if (!File.Exists(input)) throw new FileNotFoundException("Input file does not exist.", input);

        long length;
        string? sha256 = null;
        ArchivalPreflightReport report;
        await using (var stream = File.OpenRead(input))
        {
            length = stream.Length;
            // Bounded and cancellable: oversized input is never hashed and the checker rejects it by length.
            if (length <= options.LoadOptions.MaxInputBytes)
            {
                sha256 = Convert.ToHexString(await SHA256.HashDataAsync(stream, token).ConfigureAwait(false)).ToLowerInvariant();
                stream.Position = 0;
            }

            report = await new ArchivalPreflightChecker().CheckAsync(stream, options, token).ConfigureAwait(false);
        }

        var ruleIds = report.RuleResults.Select(result => result.RuleId).Distinct(StringComparer.Ordinal)
            .OrderBy(id => id, StringComparer.Ordinal).ToArray();
        if (json)
        {
            var outcomes = new JsonObject();
            foreach (var id in ruleIds) outcomes[id] = report.GetRuleOutcome(id).ToString();
            var document = new JsonObject
            {
                ["schema"] = PreflightSchema,
                // File name only: the report must not reveal directory structure.
                ["input"] = new JsonObject { ["fileName"] = Path.GetFileName(input), ["bytes"] = length, ["sha256"] = sha256 },
                ["ruleOutcomes"] = outcomes,
                ["report"] = JsonSerializer.SerializeToNode(report, PreflightJson)
            };
            var text = document.ToJsonString(PreflightJson) + Environment.NewLine;
            if (outputPath is null) Console.Out.Write(text);
            else
            {
                await using var output = new AtomicOutput(outputPath);
                var bytes = new UTF8Encoding(false).GetBytes(text);
                await output.Stream.WriteAsync(bytes, token).ConfigureAwait(false);
                await output.CommitAsync(token).ConfigureAwait(false);
            }
        }
        else
        {
            Console.WriteLine($"Preflight {report.SummaryOutcome} (run {report.RunState}, profile {report.ProfileVersion})");
            foreach (var id in ruleIds) Console.WriteLine($"{id} {report.GetRuleOutcome(id)}");
            foreach (var result in report.RuleResults.Where(result => result.Outcome != PreflightOutcome.Pass))
                Console.WriteLine($"  {result.RuleId} {result.Outcome} {result.ReasonCode} {result.Evidence.Entry} {result.Evidence.Selector}");
            Console.WriteLine($"Coverage gaps: {report.CoverageGaps.Count}");
        }

        if (report.RunState == PreflightRunState.Incomplete || report.SummaryOutcome == PreflightOutcome.Unknown) return 3;
        return report.SummaryOutcome == PreflightOutcome.Fail ? 2 : 0;
    }
}
