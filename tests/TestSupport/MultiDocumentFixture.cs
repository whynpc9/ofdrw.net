using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.TestSupport;

/// <summary>Small deterministic, redistributable two-DocBody package built from the public writer.</summary>
internal static class MultiDocumentFixture
{
    internal static async Task<Dictionary<string, byte[]>> CreateEntriesAsync()
    {
        var merged = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        XDocument? root = null;
        for (var index = 0; index < 2; index++)
        {
            var package = new OfdDocumentPackage();
            package.Options.DocumentId = "Doc_" + index;
            package.Options.Metadata.Title = index == 0 ? "First" : "Second";
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Arial" });
            for (var pageIndex = 0; pageIndex < 2; pageIndex++)
            {
                var page = new OfdPage { Index = pageIndex, WidthMillimeters = 100, HeightMillimeters = 80 };
                page.Elements.Add(new OfdTextElement { Text = (index == 0 ? "FIRST DOC0" : "SECOND DOC1") + (pageIndex == 0 ? "" : " PAGE2"),
                    FontName = "Arial", FontResourceId = "10", XMillimeters = 10, YMillimeters = 10,
                    WidthMillimeters = 70, HeightMillimeters = 12 });
                package.Pages.Add(page);
            }
            using var stream = new MemoryStream();
            await new OfdPackageWriter().WriteAsync(package, stream);
            var entries = Unzip(stream);
            var ofd = XDocument.Parse(Encoding.UTF8.GetString(entries["OFD.xml"]));
            if (root is null) root = ofd;
            else root.Root!.Add(new XElement(ofd.Root!.Elements().Single()));
            foreach (var entry in entries.Where(entry => entry.Key != "OFD.xml")) merged.Add(entry.Key, entry.Value);
        }
        merged["OFD.xml"] = Encoding.UTF8.GetBytes(root!.ToString());
        return merged;
    }

    internal static async Task<byte[]> CreateZipBytesAsync()
    { using var stream = Zip(await CreateEntriesAsync()); return stream.ToArray(); }

    internal static MemoryStream Zip(IDictionary<string, byte[]> entries)
    {
        var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var file = zip.CreateEntry(entry.Key).Open(); file.Write(entry.Value); }
        stream.Position = 0; return stream;
    }

    internal static Dictionary<string, byte[]> Unzip(Stream stream)
    {
        stream.Position = 0;
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, true);
        var entries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        { using var file = entry.Open(); using var copy = new MemoryStream(); file.CopyTo(copy); entries.Add(entry.FullName, copy.ToArray()); }
        return entries;
    }
}
