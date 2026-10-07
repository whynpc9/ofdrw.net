using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Docx;
using Ofdrw.Net.Converter.Docx.Converters;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Signatures.SesInterop;
using Ofdrw.Net.Signatures.Signing;
using Ofdrw.Net.Signatures.Verification;

if (args.Length != 3) throw new ArgumentException("Expected repository root, output directory and licensed CI font directory.");
var root = Path.GetFullPath(args[0]); var output = Path.GetFullPath(args[1]); var fonts = Path.GetFullPath(args[2]);
Directory.CreateDirectory(output);
string PathFor(string name) => Path.Combine(output, name);
var fontPath = Path.Combine(fonts, "Ofdrw-CI-NotoSansCJKsc-Regular.ttf");
var fontBytes = File.ReadAllBytes(fontPath);
File.Copy(Path.Combine(fonts, "Ofdrw-CI-Noto-OFL.txt"), PathFor("Noto-OFL.txt"), true);
var sourceDocx = PathFor("licensed-layout.docx");
File.Copy(Path.Combine(root, "e2e/Ofdrw.Net.Converter.Docx.E2E/testdata/generated-layout.docx"), sourceDocx, true);
// Keep the established deterministic content, styles, tables and pagination, with a redistributable face.
using (var zip = ZipFile.Open(sourceDocx, ZipArchiveMode.Update))
{
    foreach (var path in new[] { "word/document.xml", "word/styles.xml" })
    {
        var entry = zip.GetEntry(path)!; XDocument xml;
        using (var stream = entry.Open()) xml = XDocument.Load(stream);
        var ns = xml.Root!.Name.Namespace;
        foreach (var run in xml.Descendants(ns + "r"))
        {
            var properties = run.Element(ns + "rPr");
            if (properties is null) { properties = new XElement(ns + "rPr"); run.AddFirst(properties); }
            if (properties.Element(ns + "rFonts") is null) properties.AddFirst(new XElement(ns + "rFonts"));
        }
        foreach (var face in xml.Descendants(ns + "rFonts"))
        {
            face.RemoveAttributes();
            foreach (var name in new[] { "ascii", "hAnsi", "eastAsia", "cs" }) face.SetAttributeValue(ns + name, "Noto Sans CJK SC");
        }
        entry.Delete(); using var target = zip.CreateEntry(path).Open(); xml.Save(target);
    }
}
var identity = SesTestIdentity.Generate(); var certificate = identity.CertificateDer;
File.WriteAllBytes(PathFor("test-certificate.der"), certificate); // Public certificate only; never persist the private key.
var samples = new List<object>();
foreach (var mode in new[] { "native", "default" })
{
    var baseline = "baseline-" + mode;
    var options = new DocxConversionOptions { FontDirectories = { fonts } };
    if (mode == "native") options.OfdMode = DocxToOfdMode.Native;
    using (var input = File.OpenRead(sourceDocx))
    using (var target = File.Create(PathFor(baseline + ".ofd"))) await new DocxToOfdConverter(options).ConvertAsync(input, target);
    // Embed the exact licensed face used for measurement, making PDF/Preview independent of local viewers' fonts.
    byte[] embedded;
    using (var input = File.OpenRead(PathFor(baseline + ".ofd")))
    {
        var package = await new OfdReader().ReadAsync(input);
        foreach (var font in package.Fonts) { font.Data = fontBytes; font.FileName = "Ofdrw-CI-NotoSansCJKsc-Regular.ttf"; }
        using var target = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, target);
        embedded = target.ToArray();
    }
    File.WriteAllBytes(PathFor(baseline + ".ofd"), embedded);
    var baselineText = await Text(baseline); await Pdf(baseline);
    foreach (var version in new[] { SesVersion.V1, SesVersion.V4 })
    {
        var name = mode + "-" + version.ToString().ToLowerInvariant();
        using (var input = File.OpenRead(PathFor(baseline + ".ofd")))
        using (var target = File.Create(PathFor(name + ".ofd")))
            await new OfdSignatureService().SignAsync(input, target, new SesTestSignatureProvider(identity, version));
        using var ordinaryInput = File.OpenRead(PathFor(name + ".ofd"));
        var ordinary = await new OfdSignatureVerifier().VerifyAsync(ordinaryInput);
        if (!ordinary.ReferenceIntegrityValid || ordinary.FullyValid || ordinary.Signatures.Single().CryptographicStatus != OfdCryptographicVerificationStatus.Unsupported)
            throw new Exception("Default verification must remain reference-only/Unsupported.");
        using var explicitInput = File.OpenRead(PathFor(name + ".ofd"));
        var explicitReport = await new OfdSignatureVerifier(new[] { new SesTestSignedValueVerifier(certificate) }).VerifyAsync(explicitInput);
        if (!explicitReport.FullyValid) throw new Exception("Explicit pinned test verification failed.");
        using var wrongInput = File.OpenRead(PathFor(name + ".ofd"));
        var wrong = await new OfdSignatureVerifier(new[] { new SesTestSignedValueVerifier(SesTestIdentity.Generate().CertificateDer) }).VerifyAsync(wrongInput);
        if (wrong.FullyValid || !wrong.ReferenceIntegrityValid) throw new Exception("Wrong pin must fail independently of references.");
        if (await Text(name) != baselineText) throw new Exception("Signing altered source text.");
        VerifyOriginalEntries(baseline, name);
        await Pdf(name);
        using var zip = ZipFile.OpenRead(PathFor(name + ".ofd"));
        var signedEntry = zip.Entries.Single(e => e.FullName.EndsWith("SignedValue.dat"));
        using var valueStream = signedEntry.Open(); using var value = new MemoryStream(); valueStream.CopyTo(value);
        File.WriteAllBytes(PathFor(name + "-SignedValue.dat"), value.ToArray());
        var parsed = SesSignedValueReader.Parse(value.ToArray());
        samples.Add(new { name, mode, version = (int)version, pages = await Pages(name),
            referenceIntegrityValid = ordinary.ReferenceIntegrityValid, defaultFullyValid = ordinary.FullyValid,
            defaultCryptographicStatus = ordinary.Signatures.Single().CryptographicStatus.ToString(),
            explicitTestFullyValid = explicitReport.FullyValid, wrongPinFullyValid = wrong.FullyValid,
            textEqual = true, originalEntriesPreserved = true, parsed.PropertyInformation,
            signedValueBytes = value.Length, ofdBytes = new FileInfo(PathFor(name + ".ofd")).Length,
            baselineBytes = new FileInfo(PathFor(baseline + ".ofd")).Length });
    }
}
File.WriteAllText(PathFor("functional.json"), JsonSerializer.Serialize(new {
    environment = new { runtime = System.Runtime.InteropServices.RuntimeInformation.FrameworkDescription,
        os = System.Runtime.InteropServices.RuntimeInformation.OSDescription,
        architecture = System.Runtime.InteropServices.RuntimeInformation.ProcessArchitecture.ToString() },
    certificateSha256 = Convert.ToHexString(SHA256.HashData(certificate)),
    fontSha256 = Convert.ToHexString(SHA256.HashData(fontBytes)),
    capabilities = new { coreBuiltInSesSm2 = OfdCryptographicCapabilities.SupportsBuiltInSesSm2Verification,
        explicitTestSelfSigning = SesInteropCapabilities.SupportsExplicitTestSelfSigning }, samples,
    limitations = new[] { "Dev/interop self-signing only; no reader recognition or legal validity", "No qualified TSA, certificate-chain/revocation or signature appearance validation", "Fresh cryptographic randomness; deterministic layout/content only", "Preview acceptance is recorded separately" }
}, new JsonSerializerOptions { WriteIndented = true }));
Console.WriteLine("Four native/default V1/V4 package-consumption samples passed; default verification remains Unsupported.");

async Task<string> Text(string name)
{
    using var input = File.OpenRead(PathFor(name + ".ofd"));
    var text = new OfdTextExtractor().Extract(await new OfdReader().ReadAsync(input), includeTemplates: true);
    File.WriteAllText(PathFor(name + ".txt"), text); return text;
}
async Task<int> Pages(string name)
{
    using var input = File.OpenRead(PathFor(name + ".ofd")); return (await new OfdReader().ReadAsync(input)).Pages.Count;
}
async Task Pdf(string name)
{
    using var input = File.OpenRead(PathFor(name + ".ofd")); using var target = File.Create(PathFor(name + ".pdf"));
    await new OfdToPdfConverter().ConvertAsync(input, target);
}
void VerifyOriginalEntries(string baseline, string signed)
{
    using var before = ZipFile.OpenRead(PathFor(baseline + ".ofd")); using var after = ZipFile.OpenRead(PathFor(signed + ".ofd"));
    foreach (var entry in before.Entries.Where(e => e.FullName != "OFD.xml"))
    {
        using var source = entry.Open(); using var target = after.GetEntry(entry.FullName)!.Open();
        using var left = new MemoryStream(); using var right = new MemoryStream(); source.CopyTo(left); target.CopyTo(right);
        if (!left.ToArray().SequenceEqual(right.ToArray())) throw new Exception("Signing changed original entry: " + entry.FullName);
    }
}
