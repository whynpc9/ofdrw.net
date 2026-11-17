using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using Ofdrw.Net.Archival;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

// Package consumer for the opt-in Ofdrw.Net.Archival read-only assessment.
// Usage: <repository root> <output directory>. Uses the open NotoSans fixture; no host font.
var root = args[0];
var output = Directory.CreateDirectory(args[1]).FullName;
var font = File.ReadAllBytes(Path.Combine(root, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));

async Task<byte[]> Layout(bool embed)
{
    var document = new FlowDocument();
    document.Options.DefaultFontFamily = "Noto Sans";
    document.Blocks.Add(new Paragraph("Archival package consumer: read-only assessment."));
    var package = document.Render();
    if (embed) foreach (var resource in package.Fonts) resource.Data = font;
    using var stream = new MemoryStream();
    await new OfdPackageWriter().WriteAsync(package, stream);
    return stream.ToArray();
}

async Task<ArchivalAssessmentReport> Assess(byte[] bytes, ArchivalAssessmentOptions options)
{
    var hash = SHA256.HashData(bytes);
    using var input = new MemoryStream(bytes, false);
    var report = await new ArchivalAssessmentService().AssessAsync(input, options);
    Check(hash.SequenceEqual(SHA256.HashData(input.ToArray())), "input bytes changed");
    return report;
}

static void Check(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException("Archival package E2E: " + message);
}

var embedded = await Layout(true);
var report = await Assess(embedded, new ArchivalAssessmentOptions());
Check(report.ExecutionState == ArchivalExecutionState.Complete, "run not complete");
Check(report.InputSupport.Support == ArchivalInputSupport.Supported, "input not in profile");
foreach (var clause in new[] { "6.2.6(b)(1)", "6.2.6(c)(1)", "6.2.6(c)(2)" })
    Check(report.GetClause(clause).Outcome == PreflightOutcome.Pass, clause + " is " + report.GetClause(clause).Outcome);
Check(report.Document.StandardOutcome != PreflightOutcome.Pass, "overall outcome must never be Pass");
Check(report.Document.Gaps.Contains("NoFullStandardAssessment"), "overall gap missing");
Check(report.Coverage!.FoundationalRules.LocalPreflight.ProfileVersion == ArchivalPreflightReport.KnownCustomTagsProfileVersion,
    "archival entry did not use the known-custom-tags APF profile");

var nameOnly = await Layout(false);
var strict = await Assess(nameOnly, new ArchivalAssessmentOptions { FontExemptionPolicy = ArchivalFontExemptionPolicy.NoExemptionsAccepted });
var failure = strict.GetClause("6.2.6(b)(1)").Observations.Single();
Check(failure.Outcome == PreflightOutcome.Fail && failure.ReasonCode == "FontPayloadNotEmbedded", "no-exemption policy did not locate the font");
var lenient = await Assess(nameOnly, new ArchivalAssessmentOptions());
Check(lenient.GetClause("6.2.6(b)(1)").Outcome == PreflightOutcome.Unknown, "default policy must leave the exemption Unknown");

var json = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, Converters = { new JsonStringEnumConverter() } };
File.WriteAllText(Path.Combine(output, "archival-layout-report.json"), JsonSerializer.Serialize(report, json));
File.WriteAllText(Path.Combine(output, "archival-receipt.json"), JsonSerializer.Serialize(new
{
    archivalAssembly = typeof(ArchivalAssessmentService).Assembly.Location,
    executionState = report.ExecutionState.ToString(),
    standardOutcome = report.Document.StandardOutcome.ToString(),
    fontClause = report.GetClause("6.2.6(b)(1)").Outcome.ToString(),
    strictReason = failure.ReasonCode,
    lenientReason = lenient.GetClause("6.2.6(b)(1)").Observations.Single().ReasonCode
}, json));
Console.WriteLine($"[Archival E2E] {report.ExecutionState}; 6.2.6(b)(1)={report.GetClause("6.2.6(b)(1)").Outcome}; strict={failure.ReasonCode}");
