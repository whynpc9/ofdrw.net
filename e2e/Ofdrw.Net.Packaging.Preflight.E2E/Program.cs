using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using Ofdrw.Net.Packaging.Preflight;

var output = Path.GetFullPath(args.Length == 1 ? args[0] : "preflight-output");
Directory.CreateDirectory(output);
var serializer = new JsonSerializerOptions { WriteIndented = true };
serializer.Converters.Add(new JsonStringEnumConverter());
const string ns = "http://www.ofdspec.org";
string Xml(string root, string content, string attributes = "") => $"<{root} xmlns='{ns}' {attributes}>{content}</{root}>";
var entries = new Dictionary<string, string>
{
    ["OFD.xml"] = Xml("OFD", "<DocBody><DocRoot>Doc_0/Document.xml</DocRoot><Signatures>Doc_0/Signs/Signatures.xml</Signatures></DocBody>"),
    ["Doc_0/Document.xml"] = Xml("Document", "<CommonData><PublicRes>PublicRes.xml</PublicRes></CommonData><Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages><Attachments>Attachs/Attachments.xml</Attachments>"),
    ["Doc_0/Pages/P0/Content.xml"] = Xml("Page", "<Content><Layer><TextObject Font='1'><TextCode>归档 Archive</TextCode></TextObject></Layer></Content>"),
    ["Doc_0/PublicRes.xml"] = Xml("Res", "<Fonts><Font ID='1'><FontFile>font.bin</FontFile></Font></Fonts>", "BaseLoc='Res'"),
    ["Doc_0/Res/font.bin"] = "fake", // Nonempty sentinel, deliberately not a font.
    ["Doc_0/Attachs/Attachments.xml"] = Xml("Attachments", "<Attachment ID='A1'><FileLoc>a.txt</FileLoc></Attachment>"),
    ["Doc_0/Attachs/a.txt"] = "",
    ["Doc_0/Signs/Signatures.xml"] = Xml("Signatures", "<Signature ID='S1' BaseLoc='Sign_0/Signature.xml'/>"),
    ["Doc_0/Signs/Sign_0/Signature.xml"] = Xml("Signature", "<SignedInfo><References><Reference FileRef='/Doc_0/Document.xml'/></References></SignedInfo><SignedValue>SignedValue.dat</SignedValue>"),
    ["Doc_0/Signs/Sign_0/SignedValue.dat"] = "invalid signature"
};
var manifest = new List<object>();
await Check("presence-only", PreflightOutcome.Pass, new());
entries.Remove("Doc_0/Res/font.bin");
await Check("missing-font", PreflightOutcome.Fail, new());
entries["Encryptions.xml"] = "not parsed or decrypted";
await Check("encrypted", PreflightOutcome.Fail, new());
entries = new Dictionary<string, string>
{
    ["OFD.xml"] = Xml("OFD", "<DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody>"),
    ["Doc_0/Document.xml"] = Xml("Document", "<Pages/>")
};
await Check("all-inapplicable", PreflightOutcome.NotApplicable, new() { RequireNoEncryption = false, RequireSingleDocument = false });
await CheckArchiveCopy();
await File.WriteAllTextAsync(Path.Combine(output, "manifest.json"), JsonSerializer.Serialize(new
{
    fixtures = manifest,
    assembly = typeof(ArchivalPreflightChecker).Assembly.Location,
    assemblySha256 = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(typeof(ArchivalPreflightChecker).Assembly.Location))),
    runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
    boundary = "Synthetic ZIP/XML presence fixtures only; invalid font and signature sentinels are not validity or visual evidence."
}, serializer));
Console.WriteLine($"Public package API: {manifest.Count} cases passed; reports at {output}");

async Task Check(string name, PreflightOutcome expected, ArchivalPreflightOptions options)
{
    var file = Path.Combine(output, name + ".ofd");
    using (var bytes = new MemoryStream())
    {
        using (var zip = new ZipArchive(bytes, ZipArchiveMode.Create, true))
            foreach (var pair in entries.OrderBy(x => x.Key, StringComparer.Ordinal))
            {
                var entry = zip.CreateEntry(pair.Key);
                entry.LastWriteTime = new DateTimeOffset(2020, 1, 1, 0, 0, 0, TimeSpan.Zero);
                using var stream = entry.Open(); var payload = Encoding.UTF8.GetBytes(pair.Value); stream.Write(payload);
            }
        await File.WriteAllBytesAsync(file, bytes.ToArray());
    }
    var before = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(file)));
    using var input = File.OpenRead(file);
    var report = await new ArchivalPreflightChecker().CheckAsync(input, options);
    if (report.SummaryOutcome != expected || report.RunState != PreflightRunState.Complete ||
        !input.CanRead || input.Position != 0 || report.RuleResults.Select(x => x.RuleId).Distinct().Count() != 6)
        throw new InvalidOperationException($"Unexpected public API result for {name}: {report.SummaryOutcome} / {report.RunState}");
    if (report.Certification != "NotAssessed" || report.RuleResults.Any(x => x.Clause != "" || x.StandardStatus != "Unverified"))
        throw new InvalidOperationException("Preflight must not imply standard certification.");
    var after = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(file)));
    if (before != after) throw new InvalidOperationException("Input changed.");
    await File.WriteAllTextAsync(Path.Combine(output, name + ".json"), JsonSerializer.Serialize(report, serializer));
    manifest.Add(new { name, expected = expected.ToString(), beforeSha256 = before, afterSha256 = after, bytes = input.Length });
}

async Task CheckArchiveCopy()
{
    var source = Path.Combine(output, "archival-subset-source.ofd");
    var target = Path.Combine(output, "archival-subset-copy.ofd");
    var items = new Dictionary<string, string>
    {
        ["OFD.xml"] = Xml("OFD", "<DocBody><DocRoot>Doc_0/Document.xml</DocRoot></DocBody>"),
        ["Doc_0/Document.xml"] = Xml("Document", "<Pages><Page ID='1' BaseLoc='Pages/P0/Content.xml'/></Pages><Permissions><Edit>false</Edit></Permissions><VPreferences><HideToolbar>true</HideToolbar></VPreferences>"),
        ["Doc_0/Pages/P0/Content.xml"] = Xml("Page", "<Content><Layer><PathObject/></Layer></Content>")
    };
    using (var file = File.Create(source))
    using (var zip = new ZipArchive(file, ZipArchiveMode.Create))
        foreach (var item in items.OrderBy(x => x.Key, StringComparer.Ordinal))
        {
            using var stream = zip.CreateEntry(item.Key).Open();
            var bytes = Encoding.UTF8.GetBytes(item.Value); stream.Write(bytes);
        }
    var sourceSha = Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(source)));
    using (var input = File.OpenRead(source))
    {
        var report = await new ArchivalRuleService().CheckAsync(input, new());
        if (report.RuleResults.Count != 2 || report.RuleResults.Any(x => x.Outcome != PreflightOutcome.Fail) ||
            report.RuleResults[0].Rule.Clause != "6.2.2(a)" || report.RuleResults[1].Rule.Clause != "6.2.2(b)" ||
            report.Certification != "NotAssessed")
            throw new InvalidOperationException("Archival subset preflight did not report the two verified clauses.");
    }
    var copy = await new ArchivalCopyConverter().ConvertAsync(source, target, new());
    if (copy.Postflight.SummaryOutcome != PreflightOutcome.Pass || copy.RemovedDeclarations.Count != 2 ||
        copy.SignaturesInvalidated || sourceSha != Convert.ToHexString(SHA256.HashData(await File.ReadAllBytesAsync(source))))
        throw new InvalidOperationException("Archival subset copy did not preserve input and pass postflight.");
    await File.WriteAllTextAsync(Path.Combine(output, "archival-subset-result.json"), JsonSerializer.Serialize(new
    {
        sourceSha256 = sourceSha, copy.OutputSha256, copy.AuditPath, copy.RemovedDeclarations,
        copy.SignaturesInvalidated, result = copy.Postflight.SummaryOutcome.ToString(), copy.Postflight.Certification
    }, serializer));
}
