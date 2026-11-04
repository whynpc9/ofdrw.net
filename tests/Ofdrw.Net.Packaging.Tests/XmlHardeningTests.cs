using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Validation;
using Ofdrw.Net.Signatures.Verification;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class XmlHardeningTests
{
    private const string Dtd = "<!DOCTYPE ofd:OFD [ <!ENTITY a \"aaaaaaaaaa\"> ]>";

    [Fact]
    public async Task StructureChecker_RejectsDtd()
    {
        var archive = await Load(await WithDtdInOfdXml());

        Assert.Contains(OfdPackageStructureChecker.Check(archive), issue => issue.Code == "invalid_ofd_xml" && issue.IsError);
    }

    [Fact]
    public async Task SignatureVerifier_RejectsDtd()
    {
        using var input = new MemoryStream(await WithDtdInOfdXml());

        var report = await new OfdSignatureVerifier().VerifyAsync(input);

        Assert.Contains(report.Issues, issue => issue.StartsWith("OFD.xml could not be parsed", StringComparison.Ordinal));
    }

    [Fact]
    public async Task StructureChecker_ParsesRepeatedDocRootOnce()
    {
        // A missing document makes each validation visible as one issue.
        var valid = await Valid();
        byte[] WithBodies(int count) => Rewrite(valid, "OFD.xml", xml =>
        {
            var document = XDocument.Parse(xml);
            var body = document.Root!.Elements().First(element => element.Name.LocalName == "DocBody");
            body.Elements().First(element => element.Name.LocalName == "DocRoot").Value = "Doc_9/Document.xml";
            for (var index = 1; index < count; index++) document.Root.Add(new XElement(body));
            return document.ToString(SaveOptions.DisableFormatting);
        });
        var once = OfdPackageStructureChecker.Check(await Load(WithBodies(1)));
        var repeated = OfdPackageStructureChecker.Check(await Load(WithBodies(50)));

        Assert.Equal("missing_document_xml", Assert.Single(once).Code);
        Assert.Equal(once.Select(issue => issue.Code), repeated.Select(issue => issue.Code));
    }

    [Fact]
    public void IsVisible_UsesHardenedParser()
    {
        var text = new OfdTextElement { Text = "A", SourceXml = Dtd.Replace("ofd:OFD", "TextObject") + "<TextObject Visible=\"false\"/>" };

        Assert.Throws<XmlException>(() => OfdGraphicXmlContract.IsVisible(text));
        Assert.False(OfdGraphicXmlContract.IsVisible(new OfdTextElement { Text = "A", SourceXml = "<TextObject Visible=\"false\"/>" }));
    }

    private static async Task<byte[]> Valid()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 80 });
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        return output.ToArray();
    }

    private static async Task<byte[]> WithDtdInOfdXml() => Rewrite(await Valid(), "OFD.xml", xml =>
    {
        var start = xml.IndexOf("<ofd:OFD", StringComparison.Ordinal);
        Assert.True(start >= 0, "The writer's root prefix changed; update the DOCTYPE name.");
        return xml.Substring(0, start) + Dtd + xml.Substring(start);
    });

    private static Task<OfdPackageArchive> Load(byte[] ofd) => new OfdPackageLoader().LoadAsync(new MemoryStream(ofd));

    private static byte[] Rewrite(byte[] ofd, string name, Func<string, string> change)
    {
        var entries = new Dictionary<string, byte[]>(StringComparer.Ordinal);
        using (var zip = new ZipArchive(new MemoryStream(ofd)))
            foreach (var entry in zip.Entries)
            {
                using var stream = entry.Open();
                using var copy = new MemoryStream();
                stream.CopyTo(copy);
                entries[entry.FullName] = copy.ToArray();
            }
        entries[name] = Encoding.UTF8.GetBytes(change(Encoding.UTF8.GetString(entries[name]).TrimStart('﻿')));
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var (path, data) in entries)
            {
                using var stream = zip.CreateEntry(path).Open();
                stream.Write(data, 0, data.Length);
            }
        return output.ToArray();
    }
}
