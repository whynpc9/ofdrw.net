using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using System.IO.Compression;
using System.Xml.Linq;
using System.Text;
using Ofdrw.Net.Signatures.Crypto;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Cli.Tests;

public sealed class DocumentToolsCliTests : IDisposable
{
    private readonly string _directory = Path.Combine(Path.GetTempPath(), "ofd-document-tools-" + Guid.NewGuid().ToString("N"));
    public DocumentToolsCliTests() => Directory.CreateDirectory(_directory);
    private string PathFor(string name) => Path.Combine(_directory, name);
    [Fact]
    public async Task Commands_UseOneBasedPagesAndAtomicInPlaceOutput()
    {
        var input = PathFor("source.ofd"); var output = PathFor("output.ofd");
        var source = new OfdDocumentPackage();
        for (var i = 0; i < 2; i++)
        {
            var page = new OfdPage { Index = i, WidthMillimeters = 100 + i, HeightMillimeters = 150 };
            page.Elements.Add(new OfdTextElement { Text = "PAGE" + (i + 1) }); source.Pages.Add(page);
        }
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(source, stream);
        Assert.Equal(0, await global::Cli.RunAsync(["watermark", input, input, "--pages", "2", "--text", "DRAFT"]));
        await using (var stream = File.OpenRead(input))
        {
            var read = await new OfdReader().ReadAsync(stream);
            Assert.Single(read.Pages[0].Elements); Assert.Equal(2, read.Pages[1].Elements.Count);
        }
        Assert.Equal(0, await global::Cli.RunAsync(["split", input, output, "--pages", "2"]));
        await using (var stream = File.OpenRead(output)) Assert.Equal(101, Assert.Single((await new OfdReader().ReadAsync(stream)).Pages).WidthMillimeters);
        Assert.Equal(0, await global::Cli.RunAsync(["mix", output, input, "2", input, "1"]));
        await using (var stream = File.OpenRead(output))
        {
            var page = Assert.Single((await new OfdReader().ReadAsync(stream)).Pages);
            Assert.Equal(101, page.WidthMillimeters); Assert.Equal(new[] { "PAGE2", "DRAFT", "PAGE1" }, page.Elements.OfType<OfdTextElement>().Select(text => text.Text));
        }
        Assert.Equal(0, await global::Cli.RunAsync(["clean-signatures", output, output]));
        Assert.Equal(0, await global::Cli.RunAsync(["verify-signatures", output]));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("watermark", "--pages", "0")]
    [InlineData("split", "--pages", "0")]
    [InlineData("watermark", "--unknown", "value")]
    public async Task BadCommand_PreservesExistingOutput(string command, string option, string value)
    {
        var input = PathFor("source.ofd"); var output = PathFor("output.ofd");
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(new OfdDocumentPackage(), stream);
        await File.WriteAllTextAsync(output, "unchanged");
        Assert.Equal(1, await global::Cli.RunAsync([command, input, output, option, value]));
        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Theory]
    [InlineData("split")]
    [InlineData("watermark")]
    public async Task ExtendedPageContainer_PreservesExistingCliOutput(string command)
    {
        var input = PathFor("source.ofd"); var output = PathFor("existing.ofd"); var source = new OfdDocumentPackage(); source.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 100 });
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(source, stream);
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Update))
        {
            var entry = zip.Entries.Single(item => item.FullName.EndsWith("/Content.xml", StringComparison.Ordinal)); XDocument xml; using (var stream = entry.Open()) xml = XDocument.Load(stream);
            xml.Root!.SetAttributeValue(XNamespace.Get("urn:vendor") + "Payload", "keep.bin"); var path = entry.FullName; entry.Delete(); using var target = zip.CreateEntry(path).Open(); target.Write(Encoding.UTF8.GetBytes(xml.ToString()));
        }
        await File.WriteAllTextAsync(output, "unchanged"); var args = command == "split" ? new[] { command, input, output, "--pages", "1" } : new[] { command, input, output, "--pages", "1", "--text", "DRAFT" };
        Assert.Equal(1, await global::Cli.RunAsync(args)); Assert.Equal("unchanged", await File.ReadAllTextAsync(output)); Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    [Fact]
    public async Task Verify_MissingDeclaredListMustNotReportNoSignatures()
    {
        var input = PathFor("broken.ofd");
        await using (var stream = File.Create(input)) await new OfdPackageWriter().WriteAsync(new OfdDocumentPackage(), stream);
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Update))
        {
            var rootEntry = zip.GetEntry("OFD.xml")!;
            XDocument root; using (var stream = rootEntry.Open()) root = XDocument.Load(stream);
            root.Root!.Elements().Single().Add(new XElement(root.Root.Name.Namespace + "Signatures", "missing.xml")); rootEntry.Delete();
            using var output = zip.CreateEntry("OFD.xml").Open(); var bytes = Encoding.UTF8.GetBytes(root.ToString()); output.Write(bytes);
        }
        Assert.Equal(3, await global::Cli.RunAsync(["verify-signatures", input]));
    }

    [Theory]
    [InlineData("missing", 3)]
    [InlineData("malformed", 3)]
    [InlineData("empty", 3)]
    [InlineData("blank", 3)]
    [InlineData("complete", 2)]
    public async Task Audit001_CliIncompleteReport(string kind, int expectedExitCode)
    {
        var input = PathFor("declarations.ofd");
        XNamespace ns = "http://www.ofdspec.org/2016";
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Create))
        {
            var body = new XElement(ns + "DocBody", new XElement(ns + "Signatures", "Signs/Signatures.xml"));
            if (kind != "complete") body.Add(new XElement(ns + "Signatures", kind == "blank" ? "   " : "broken.xml"));
            WriteVerificationEntry(zip, "OFD.xml", new XElement(ns + "OFD", body).ToString());
            WriteVerificationEntry(zip, "content.txt", "protected content");
            WriteVerificationEntry(zip, "Signs/Signatures.xml", new XElement(ns + "Signatures", new XElement(ns + "Signature",
                new XAttribute("ID", "1"), new XAttribute("BaseLoc", "Signature.xml"))).ToString());
            var checkValue = Convert.ToBase64String(OfdDigestAlgorithms.Compute(
                OfdDigestAlgorithms.Sha256Oid, Encoding.UTF8.GetBytes("protected content")));
            var references = new XElement(ns + "References",
                new XAttribute("CheckMethod", OfdDigestAlgorithms.Sha256Oid),
                new XElement(ns + "Reference", new XAttribute("FileRef", "/content.txt"),
                    new XElement(ns + "CheckValue", checkValue)));
            var signature = new XElement(ns + "Signature",
                new XElement(ns + "SignedInfo",
                    new XElement(ns + "SignatureMethod", "urn:audit001:unregistered"), references),
                new XElement(ns + "SignedValue", "SignedValue.dat"));
            WriteVerificationEntry(zip, "Signs/Signature.xml", signature.ToString());
            WriteVerificationEntry(zip, "Signs/SignedValue.dat", "unregistered algorithm fixture");
            if (kind == "malformed") WriteVerificationEntry(zip, "broken.xml", "<Signatures>");
            if (kind == "empty") WriteVerificationEntry(zip, "broken.xml", new XElement(ns + "Signatures").ToString());
        }

        await using (var stream = File.OpenRead(input))
        {
            var report = await new OfdSignatureVerifier().VerifyAsync(stream);
            Assert.True(Assert.Single(report.Signatures).ReferenceIntegrityValid);
            Assert.Equal(OfdCryptographicVerificationStatus.Unsupported, report.Signatures[0].CryptographicStatus);
        }
        Assert.Equal(expectedExitCode, await global::Cli.RunAsync(["verify-signatures", input]));
    }

    private static void WriteVerificationEntry(ZipArchive zip, string path, string text)
    {
        using var stream = zip.CreateEntry(path).Open();
        stream.Write(Encoding.UTF8.GetBytes(text));
    }

    [Fact]
    public async Task Clean_RejectsOrdinaryZipWithoutReplacingExistingOutput()
    {
        var input = PathFor("ordinary.zip"); var output = PathFor("existing.ofd");
        using (var zip = ZipFile.Open(input, ZipArchiveMode.Create))
        { using var file = zip.CreateEntry("ordinary.txt").Open(); file.Write(new byte[] { 1 }); }
        await File.WriteAllTextAsync(output, "unchanged");
        Assert.Equal(1, await global::Cli.RunAsync(["clean-signatures", input, output]));
        Assert.Equal("unchanged", await File.ReadAllTextAsync(output));
        Assert.Empty(Directory.GetFiles(_directory, ".ofdrw-*.tmp"));
    }

    public void Dispose() => Directory.Delete(_directory, true);
}
