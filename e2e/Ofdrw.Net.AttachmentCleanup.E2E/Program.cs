using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

var package = new OfdDocumentPackage();
package.Attachments.Add(new OfdAttachment { Name = "r18-public-api.bin", Data = [0, 1, 2, 255] });
var writer = new OfdPackageWriter();
using var input = new MemoryStream();
await writer.WriteAsync(package, input);

var logical = await new OfdReader().ReadAsync(new MemoryStream(input.ToArray()));
if (!logical.RemoveAttachment(logical.Attachments.Single())) throw new Exception("Logical removal failed.");
using var logicalOutput = new MemoryStream();
await writer.WriteAsync(logical, logicalOutput);
if (!Names(logicalOutput).Any(name => name.EndsWith("r18-public-api.bin", StringComparison.Ordinal)))
    throw new Exception("Logical removal unexpectedly erased the historical payload.");

var physical = await new OfdReader().ReadAsync(new MemoryStream(input.ToArray()));
if (!physical.RemoveAttachment(physical.Attachments.Single(), OfdAttachmentRemovalMode.Physical))
    throw new Exception("Physical removal request failed.");
using var output = new MemoryStream();
var result = await writer.WriteWithResultAsync(physical, output);
if (result.SignaturesInvalidated || result.RemovedEntries.Count != 2 ||
    result.RemovedEntries.Any(name => Names(output).Contains(name)))
    throw new Exception("Physical removal report does not match ZIP entries.");
if ((await new OfdReader().ReadAsync(new MemoryStream(output.ToArray()))).Attachments.Count != 0)
    throw new Exception("Removed attachment reappeared after package readback.");
var evidenceDirectory = Environment.GetEnvironmentVariable("R18_E2E_OUTPUT_DIR");
if (!string.IsNullOrWhiteSpace(evidenceDirectory))
{
    Directory.CreateDirectory(evidenceDirectory);
    var files = new Dictionary<string, byte[]>
    {
        ["source.ofd"] = input.ToArray(), ["logical.ofd"] = logicalOutput.ToArray(), ["physical.ofd"] = output.ToArray()
    };
    foreach (var file in files) File.WriteAllBytes(Path.Combine(evidenceDirectory, file.Key), file.Value);
    var receipt = new
    {
        sourceSha = Environment.GetEnvironmentVariable("R18_SOURCE_SHA"),
        sourceInput = "deterministic in-memory OfdDocumentPackage with one four-byte attachment",
        files = files.ToDictionary(file => file.Key, file => new { bytes = file.Value.Length,
            sha256 = Convert.ToHexString(SHA256.HashData(file.Value)) }),
        removedEntries = result.RemovedEntries,
        signaturesInvalidated = result.SignaturesInvalidated,
        result = "logical retained; physical descriptor and payload absent; readback empty"
    };
    File.WriteAllText(Path.Combine(evidenceDirectory, "receipt.json"), JsonSerializer.Serialize(receipt, new JsonSerializerOptions { WriteIndented = true }));
}
Console.WriteLine("R18 exact-package consumer: logical payload retained; physical descriptor and payload removed; readback empty.");

static HashSet<string> Names(MemoryStream stream)
{
    using var zip = new ZipArchive(new MemoryStream(stream.ToArray()), ZipArchiveMode.Read);
    return zip.Entries.Select(entry => entry.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
}
