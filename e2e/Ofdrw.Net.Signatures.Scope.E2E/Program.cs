using System.IO.Compression;
using System.Text;
using System.Text.Json;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Signatures.SesInterop;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

if (args.Length != 1) throw new ArgumentException("Expected an output directory.");
var output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var package = new OfdDocumentPackage();
package.Options.Namespace = OfdConstants.StandardNamespace;
package.Pages.Add(new OfdPage { WidthMillimeters = 210, HeightMillimeters = 297 });
package.PreservedEntries["Extras/optional.bin"] = [1, 2, 3];
using var unsigned = new MemoryStream();
await new OfdPackageWriter().WriteAsync(package, unsigned);
unsigned.Position = 0;
using (var zip = new ZipArchive(unsigned, ZipArchiveMode.Update, leaveOpen: true))
{
    var original = zip.GetEntry("OFD.xml")!;
    string xml;
    using (var reader = new StreamReader(original.Open(), Encoding.UTF8)) xml = await reader.ReadToEndAsync();
    original.Delete();
    using var writer = new StreamWriter(zip.CreateEntry("OFD.xml").Open(), new UTF8Encoding(false));
    await writer.WriteAsync(xml.Replace("Version=\"1.0\"", "Version=\"2.0\""));
}
unsigned.Position = 0;
var identity = SesTestIdentity.Generate();
var provider = new SesTestSignatureProvider(identity, SesVersion.V4);
var service = new OfdSignatureService();
using var first = new MemoryStream();
await service.SignAsync(unsigned, first, provider);
first.Position = 0;
var oldArchive = await new OfdPackageLoader().LoadAsync(first);
File.WriteAllBytes(Path.Combine(output, "r13-first.ofd"), first.ToArray());
var oldXml = oldArchive.GetBytes("Doc_0/Signs/Sign_0/Signature.xml");
var oldValue = oldArchive.GetBytes("Doc_0/Signs/Sign_0/SignedValue.dat");
var options = new OfdSignatureOptions
{
    RelativeSignatureId = "1",
    CheckMethod = OfdDigestAlgorithms.Sha256Oid,
    ProtectFile = path => path != "/Extras/optional.bin"
};
options.Parameters.Add(new OfdSignatureParameter("purpose", "package-consumer", "text"));
first.Position = 0;
using var signed = new MemoryStream();
var selection = await service.SignWithResultAsync(first, signed, provider, options);
signed.Position = 0;
var archive = await new OfdPackageLoader().LoadAsync(signed);
if (!oldXml.SequenceEqual(archive.GetBytes("Doc_0/Signs/Sign_0/Signature.xml")) ||
    !oldValue.SequenceEqual(archive.GetBytes("Doc_0/Signs/Sign_0/SignedValue.dat")))
    throw new InvalidDataException("Appending changed prior signature bytes.");
if (!selection.UnprotectedFiles.Contains("/Extras/optional.bin") ||
    !selection.ProtectedFiles.Contains("/Doc_0/Document.xml"))
    throw new InvalidDataException("Package consumer received the wrong scope report.");
signed.Position = 0;
var report = await new OfdSignatureVerifier(new[]
    { new SesTestSignedValueVerifier(identity.CertificateDer) }).VerifyAsync(signed);
if (!report.ReferenceIntegrityValid || !report.FullyValid ||
    !report.Signatures[1].RelativeIntegrityValid)
    throw new InvalidDataException("Installed package did not verify the scope and relationship.");
File.WriteAllBytes(Path.Combine(output, "r13-signed.ofd"), signed.ToArray());
File.WriteAllText(Path.Combine(output, "r13-scope.json"), JsonSerializer.Serialize(new
{
    protectedFiles = selection.ProtectedFiles,
    unprotectedFiles = selection.UnprotectedFiles,
    requiredFileCoverageValid = selection.RequiredFileCoverageValid,
    referenceIntegrityValid = report.ReferenceIntegrityValid,
    fullyValid = report.FullyValid,
    relativeIntegrityValid = report.Signatures[1].RelativeIntegrityValid,
    oldSignatureBytesUnchanged = true,
    signatureCount = report.Signatures.Count
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("Installed R13 signature packages: scope, Relative, Parameters, and append verified.");
