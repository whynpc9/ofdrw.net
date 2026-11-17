using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Encodings.Web;
using System.Text.Json;
using System.Text.Json.Serialization;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Medical;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Preflight;

// Package consumer for the opt-in Ofdrw.Net.Medical (GB/T 48666 OFD-H) profile and read-only assessment.
// Usage: <repository root> <output directory>. Uses the open NotoSans fixture; no host font.
var root = args[0];
var output = Directory.CreateDirectory(args[1]).FullName;
var font = File.ReadAllBytes(Path.Combine(root, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));

static void Check(bool condition, string message)
{
    if (!condition) throw new InvalidOperationException("Medical package E2E: " + message);
}

async Task<byte[]> Write(bool medical)
{
    var document = new FlowDocument();
    document.Options.DefaultFontFamily = "Noto Sans";
    document.Blocks.Add(new Paragraph("Medical package consumer: OFD-H profile and read-only assessment."));
    var package = document.Render();
    foreach (var resource in package.Fonts) resource.Data = font;
    if (medical)
        MedicalDocumentProfile.Apply(package.Options, new MedicalDocumentMetadata
        {
            DocumentIdentifier = "E2E-0001", PatientId = "P-E2E", ResponsibleParty = "E2E", Keywords = "e2e",
            CreationTime = new DateTimeOffset(2026, 11, 17, 0, 0, 0, TimeSpan.Zero), FileType = "Test", SecretLevel = "Internal"
        });
    using var stream = new MemoryStream();
    await new OfdPackageWriter().WriteAsync(package, stream);
    return stream.ToArray();
}

async Task<MedicalAssessmentReport> Assess(byte[] bytes)
{
    var hash = SHA256.HashData(bytes);
    using var input = new MemoryStream(bytes, false);
    var report = await new MedicalAssessmentService().AssessAsync(input);
    Check(hash.SequenceEqual(SHA256.HashData(input.ToArray())), "input bytes changed");
    Check(report.Document.StandardOutcome != PreflightOutcome.Pass, "overall outcome must never be Pass");
    return report;
}

var medical = await Write(true);
File.WriteAllBytes(Path.Combine(output, "medical-layout.ofd"), medical);
using (var zip = new ZipArchive(new MemoryStream(medical, false)))
    Check(zip.Entries[0].FullName == "OFD.xml", "linearized order did not start with OFD.xml");
var report = await Assess(medical);
Check(report.ExecutionState == MedicalExecutionState.Complete, "run not complete");
Check(report.Declaration.Status == MedicalDeclarationStatus.Medical, "declaration not OFD-H");
foreach (var clause in new[] { "7.1(a)", "6.1.2(a)", "6.1.3(a)", "6.1.3(b)", "6.2(c)", "7.2(b)", "7.3(a)", "7.3(b)" })
    Check(report.GetClause(clause).Outcome == PreflightOutcome.Pass, clause + " is " + report.GetClause(clause).Outcome);
Check(report.Archival?.ExecutionState == Ofdrw.Net.Archival.ArchivalExecutionState.Complete, "embedded archival run not complete");

var ordinary = await Assess(await Write(false));
Check(ordinary.GetClause("7.1(a)").Outcome == PreflightOutcome.Fail, "ordinary OFD must fail 7.1(a)");
Check(ordinary.Document.StandardOutcome == PreflightOutcome.Fail, "ordinary OFD must fail overall");

var json = new JsonSerializerOptions { WriteIndented = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping, Converters = { new JsonStringEnumConverter() } };
File.WriteAllText(Path.Combine(output, "medical-receipt.json"), JsonSerializer.Serialize(new
{
    medicalAssembly = typeof(MedicalAssessmentService).Assembly.Location,
    executionState = report.ExecutionState.ToString(),
    declaration = report.Declaration.Status.ToString(),
    standardOutcome = report.Document.StandardOutcome.ToString(),
    passedClauses = report.Clauses.Where(x => x.Outcome == PreflightOutcome.Pass).Select(x => x.Clause).ToArray(),
    ordinaryFailedClauses = ordinary.Clauses.Where(x => x.Outcome == PreflightOutcome.Fail).Select(x => x.Clause).ToArray()
}, json));
Console.WriteLine($"[Medical E2E] {report.ExecutionState}; declaration={report.Declaration.Status}; overall={report.Document.StandardOutcome}");
