using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class ReaderInputHygieneTests
{
    [Theory]
    [InlineData("NaN")]
    [InlineData("Infinity")]
    [InlineData("-Infinity")]
    [InlineData("1e999")]
    public async Task NonFiniteTextSize_FallsBackToDefault(string size)
    {
        var mutated = Mutate(await Valid(), "/Content.xml", xml => Regex.Replace(xml, "Size=\"[^\"]*\"", $"Size=\"{size}\""));

        var text = await ReadText(mutated);

        Assert.Equal(4d, text.FontSizeMillimeters);
    }

    [Theory]
    [InlineData("NaN")]
    [InlineData("1e999")]
    [InlineData("-5")]
    public async Task InvalidWeight_FallsBackToDefault(string weight)
    {
        var mutated = Mutate(await Valid(), "/Content.xml", xml => Regex.Replace(xml, "<ofd:TextObject ", $"<ofd:TextObject Weight=\"{weight}\" "));

        var text = await ReadText(mutated);

        Assert.Equal(OfdTextElement.DefaultWeight, text.Weight);
    }

    [Fact]
    public async Task DuplicateTemplatePage_IsInvalidData()
    {
        var template = "<ofd:TemplatePage ID=\"9\" BaseLoc=\"Tpls/Tpl_0/Content.xml\"/>";
        var mutated = Mutate(await Valid(), "Doc_0/Document.xml", xml => xml.Replace("</ofd:CommonData>", template + template + "</ofd:CommonData>"));

        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(new MemoryStream(mutated)));
    }

    [Fact]
    public async Task ListDocuments_MalformedOfdXml_IsInvalidData()
    {
        var mutated = Mutate(await Valid(), "OFD.xml", _ => "<ofd:OFD");

        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ReadAsync(new MemoryStream(mutated)));
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdReader().ListDocumentsAsync(new MemoryStream(mutated)));
    }

    private static async Task<OfdTextElement> ReadText(byte[] ofd) =>
        Assert.Single(Assert.Single((await new OfdReader().ReadAsync(new MemoryStream(ofd))).Pages).Elements.OfType<OfdTextElement>());

    private static async Task<byte[]> Valid()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage
        {
            Index = 0, WidthMillimeters = 210, HeightMillimeters = 297,
            Elements = { new OfdTextElement { Text = "Hygiene", XMillimeters = 20, YMillimeters = 20, WidthMillimeters = 60, HeightMillimeters = 10, FontSizeMillimeters = 5 } }
        });
        using var output = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, output);
        return output.ToArray();
    }

    private static byte[] Mutate(byte[] ofd, string suffix, Func<string, string> change)
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
        var target = Assert.Single(entries.Keys, name => name.EndsWith(suffix, StringComparison.Ordinal) && name.Contains("Page") == suffix.Contains("Content"));
        var before = Encoding.UTF8.GetString(entries[target]);
        var after = change(before);
        Assert.NotEqual(before, after);
        entries[target] = Encoding.UTF8.GetBytes(after);
        using var output = new MemoryStream();
        using (var zip = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true))
            foreach (var (name, data) in entries)
            {
                using var stream = zip.CreateEntry(name).Open();
                stream.Write(data, 0, data.Length);
            }
        return output.ToArray();
    }
}
