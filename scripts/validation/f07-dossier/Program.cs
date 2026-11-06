using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Ofdrw.Net.Packaging.Preflight;

var output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var sourcePath = Path.Combine(output, "source.ofd");
var resultPath = Path.Combine(output, "copy.ofd");
using (var file = File.Create(sourcePath))
using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
{
    Write("OFD.xml", "<OFD xmlns='http://www.ofdspec.org' DocType='OFD-A'><DocBody><DocRoot>Doc/Document.xml</DocRoot></DocBody></OFD>");
    Write("Doc/Document.xml", "<Document xmlns='http://www.ofdspec.org'><Permissions><Edit>false</Edit></Permissions><Pages><Page ID='1' BaseLoc='Page.xml'/></Pages></Document>");
    Write("Doc/Page.xml", "<Page xmlns='http://www.ofdspec.org'><Content><Layer><PathObject ID='2' Boundary='0 0 10 10'/></Layer></Content></Page>");
    void Write(string name, string content)
    { using var stream = zip.CreateEntry(name).Open(); stream.Write(Encoding.UTF8.GetBytes(content)); }
}
var copy = await new ArchivalCopyConverter().ConvertAsync(sourcePath, resultPath, new());
using var source = File.OpenRead(sourcePath);
using var result = File.OpenRead(resultPath);
using var audit = File.OpenRead(copy.AuditPath);
var profile = new ArchivalDossierProfile("f07-unsigned-copy/1", null,
    "R17 unsigned Permissions removal; evidence collection only", Array.Empty<KeyValuePair<string, string>>());
string Hash(Stream stream) { var digest = Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant(); stream.Position = 0; return digest; }
var evidence = new ArchivalEvidenceInput("copy-audit", "6.2.2(a)", ArchivalEvidenceKind.SourceTrace,
    audit, Hash(audit), "ArchivalCopyConverter audit; not independent approval", Hash(source), Hash(result));
var report = await new ArchivalDossierService().CheckAsync(source, result, profile, [evidence], new());
if (report.Clauses.Count != 105 || report.EvidenceIntegrity != PreflightOutcome.Pass ||
    report.StandardOutcome != PreflightOutcome.Unknown || report.SourceCoverage.StandardOutcome != PreflightOutcome.Fail ||
    !report.Evidence[0].SubjectMatches || report.MissingExternalEvidence.Count != 3 ||
    !report.ExternalAssessmentGaps.Contains("ArchivalAuthorityUnspecified") ||
    report.ExternalAssessmentOutcome != PreflightOutcome.Unknown) throw new Exception("Dossier contract failed.");
var options = new JsonSerializerOptions { WriteIndented = true };
options.Converters.Add(new JsonStringEnumConverter());
File.WriteAllText(Path.Combine(output, "dossier.json"), JsonSerializer.Serialize(report, options));
File.WriteAllText(Path.Combine(output, "environment.json"), JsonSerializer.Serialize(new {
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
    assembly = typeof(ArchivalDossierService).Assembly.Location,
    version = typeof(ArchivalDossierService).Assembly.GetName().Version?.ToString()
}, options));
Console.WriteLine($"105 obligations, evidence integrity {report.EvidenceIntegrity}, standard {report.StandardOutcome}");
