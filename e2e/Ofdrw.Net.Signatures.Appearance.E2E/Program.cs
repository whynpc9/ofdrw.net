using System;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Signatures.SesInterop;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

if (args.Length != 1) throw new ArgumentException("Expected an output directory.");
var output = Path.GetFullPath(args[0]);
Directory.CreateDirectory(output);
var source = new OfdDocumentPackage();
for (var index = 0; index < 3; index++) source.Pages.Add(new OfdPage
{ Index = index, WidthMillimeters = 100, HeightMillimeters = 80 });
using var unsigned = new MemoryStream();
await new OfdPackageWriter().WriteAsync(source, unsigned);
unsigned.Position = 0;
var options = new OfdSignatureOptions();
options.StampAppearances.Add(new OfdNormalStampPlacement(1, 10, 10, 20, 20));
var riding = new OfdRidingStampPlacement(OfdStampEdge.Right, 1, 3, 40, 20);
riding.SliceWeights.Add(1); riding.SliceWeights.Add(2); riding.SliceWeights.Add(1);
options.StampAppearances.Add(riding);
var identity = SesTestIdentity.Generate();
using var signed = new MemoryStream();
var coverage = await new OfdSignatureService().SignWithResultAsync(unsigned, signed,
    new SesTestSignatureProvider(identity, SesVersion.V4), options);
signed.Position = 0;
var verification = await new OfdSignatureVerifier(new[]
    { new SesTestSignedValueVerifier(identity.CertificateDer) }).VerifyAsync(signed);
if (!verification.FullyValid || !verification.ReferenceIntegrityValid ||
    !coverage.ProtectedFiles.Any(name => name.EndsWith("/Seal.esl", StringComparison.Ordinal)))
    throw new InvalidDataException("Installed package signature/appearance coverage failed.");
signed.Position = 0;
var archive = await new OfdPackageLoader().LoadAsync(signed);
var xmlPath = archive.EntryNames.Single(name => name.EndsWith("/Signature.xml", StringComparison.Ordinal));
var xml = XDocument.Parse(archive.ReadUtf8Text(xmlPath));
var stamps = xml.Descendants().Where(node => node.Name.LocalName == "StampAnnot").ToArray();
if (stamps.Length != 4 || stamps[0].Attribute("Clip") is not null ||
    stamps.Skip(1).Any(stamp => stamp.Attribute("Clip") is null) ||
    stamps.Select(stamp => stamp.Attribute("PageRef")?.Value).Distinct().Count() != 3)
    throw new InvalidDataException("Installed package did not write the expected stamp placements.");
var ofdPath = Path.Combine(output, "r11-stamped.ofd");
File.WriteAllBytes(ofdPath, signed.ToArray());
File.WriteAllText(Path.Combine(output, "r11-package-consumer.json"), JsonSerializer.Serialize(new
{
    pageCount = 3, stampCount = stamps.Length, referenceIntegrityValid = verification.ReferenceIntegrityValid,
    fullyValid = verification.FullyValid, protectedFiles = coverage.ProtectedFiles,
    ofdBytes = signed.Length, ofdSha256 = Convert.ToHexString(SHA256.HashData(signed.ToArray()))
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("Installed R11 signature packages produced and verified four signed stamp placements.");
