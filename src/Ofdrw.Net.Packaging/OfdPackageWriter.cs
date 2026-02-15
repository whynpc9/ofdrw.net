using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

public sealed class OfdPackageWriter
{
    public async Task WriteAsync(OfdDocumentPackage package, Stream destination, CancellationToken cancellationToken = default)
    {
        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        if (destination is null)
        {
            throw new ArgumentNullException(nameof(destination));
        }

        var entries = BuildEntries(package);
        using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);

        foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
        {
            var zipEntry = zip.CreateEntry(entry.Key, package.Options.EnableDeflateCompression ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
            using var stream = zipEntry.Open();
            await stream.WriteAsync(entry.Value, 0, entry.Value.Length, cancellationToken).ConfigureAwait(false);
        }
    }

    private Dictionary<string, byte[]> BuildEntries(OfdDocumentPackage package)
    {
        var entries = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
        var ns = XNamespace.Get(package.Options.Namespace);
        var docId = package.Options.DocumentId;

        var ofdXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "OFD",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XAttribute("DocType", package.Options.DocType),
                new XElement(ns + "DocBody",
                    new XElement(ns + "DocInfo",
                        BuildDocInfo(ns, package.Options.Metadata)
                    ),
                    new XElement(ns + "DocRoot", $"{docId}/Document.xml")
                )));
        entries[OfdConstants.OfdRootFile] = ToUtf8Bytes(ofdXml);

        var orderedPages = package.Pages.OrderBy(x => x.Index).ToList();
        if (orderedPages.Count == 0)
        {
            orderedPages.Add(new OfdPage
            {
                Index = 0,
                WidthMillimeters = package.Options.DefaultPageWidthMillimeters,
                HeightMillimeters = package.Options.DefaultPageHeightMillimeters
            });
        }

        var docXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Document",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "CommonData",
                    new XElement(ns + "PageArea",
                        new XElement(ns + "PhysicalBox", BuildBox(0, 0, orderedPages[0].WidthMillimeters, orderedPages[0].HeightMillimeters))
                    )
                ),
                new XElement(ns + "Pages",
                    orderedPages.Select((page, i) =>
                        new XElement(ns + "Page",
                            new XAttribute("ID", i + 1),
                            new XAttribute("BaseLoc", $"Pages/Page_{i}/Content.xml"))
                    )
                ),
                orderedPages.Any(x => x.Elements.OfType<OfdImageElement>().Any())
                    ? new XElement(ns + "DocumentRes", "DocumentRes.xml")
                    : null,
                package.Attachments.Count > 0
                    ? new XElement(ns + "Attachments", "Attachs/Attachments.xml")
                    : null));

        entries[$"{docId}/Document.xml"] = ToUtf8Bytes(docXml);
        entries[$"{docId}/DocumentRes.xml"] = ToUtf8Bytes(new XDocument(new XDeclaration("1.0", "UTF-8", null), new XElement(ns + "Res", new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName))));

        BuildPublicResources(package, entries, ns, docId);
        BuildPages(package, orderedPages, entries, ns, docId);
        BuildAttachments(package, entries, ns, docId);
        BuildCustomTags(package, entries, ns, docId);

        return entries;
    }

    private static void BuildPublicResources(OfdDocumentPackage package, IDictionary<string, byte[]> entries, XNamespace ns, string docId)
    {
        var fonts = package.Pages.SelectMany(p => p.Elements)
            .OfType<OfdTextElement>()
            .Select(t => t.FontName)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var fontElements = fonts.Select((font, index) =>
            new XElement(ns + "Font",
                new XAttribute("ID", index + 1),
                new XAttribute("FontName", font)));

        var publicResXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Res",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                fontElements));

        entries[$"{docId}/PublicRes.xml"] = ToUtf8Bytes(publicResXml);
    }

    private static void BuildPages(OfdDocumentPackage package, List<OfdPage> pages, IDictionary<string, byte[]> entries, XNamespace ns, string docId)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var page = pages[i];
            var pageFolder = $"{docId}/Pages/Page_{i}";
            var resources = new List<(string id, string fileName, string mediaType)>();
            var layer = new XElement(ns + "Layer", new XAttribute("ID", 1), new XAttribute("Type", "Body"));
            var elementId = 1;

            foreach (var element in page.Elements)
            {
                if (element is OfdTextElement text)
                {
                    var width = text.WidthMillimeters <= 0 ? Math.Max(text.Text.Length * text.FontSizeMillimeters * 0.5, text.FontSizeMillimeters) : text.WidthMillimeters;
                    var height = text.HeightMillimeters <= 0 ? text.FontSizeMillimeters * 1.5 : text.HeightMillimeters;
                    var textObject = new XElement(ns + "TextObject",
                        new XAttribute("ID", elementId++),
                        new XAttribute("Boundary", BuildBox(text.XMillimeters, text.YMillimeters, width, height)),
                        new XAttribute("Font", text.FontName),
                        new XAttribute("Size", ToInvariant(text.FontSizeMillimeters)),
                        text.Text);
                    layer.Add(textObject);
                }
                else if (element is OfdImageElement image)
                {
                    var id = string.IsNullOrWhiteSpace(image.ResourceId) ? $"IMG_{i}_{resources.Count + 1}" : image.ResourceId;
                    var fileName = string.IsNullOrWhiteSpace(image.FileName) ? $"Image_{resources.Count + 1}{GetImageExtension(image.MediaType)}" : image.FileName;
                    resources.Add((id, fileName, image.MediaType));

                    var imageObject = new XElement(ns + "ImageObject",
                        new XAttribute("ID", elementId++),
                        new XAttribute("Boundary", BuildBox(image.XMillimeters, image.YMillimeters, image.WidthMillimeters, image.HeightMillimeters)),
                        new XAttribute("ResourceID", id));
                    layer.Add(imageObject);

                    entries[$"{pageFolder}/Res/{fileName}"] = image.Data;
                }
            }

            var content = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement(ns + "Page",
                    new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                    new XAttribute("ID", i + 1),
                    new XElement(ns + "Area", new XElement(ns + "PhysicalBox", BuildBox(0, 0, page.WidthMillimeters, page.HeightMillimeters))),
                    new XElement(ns + "Content", layer)));
            entries[$"{pageFolder}/Content.xml"] = ToUtf8Bytes(content);

            var pageRes = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement(ns + "Res",
                    new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                    resources.Select(r =>
                        new XElement(ns + "MultiMedia",
                            new XAttribute("ID", r.id),
                            new XAttribute("Type", "Image"),
                            new XAttribute("Format", r.mediaType),
                            new XAttribute("MediaFile", $"Res/{r.fileName}")))));
            entries[$"{pageFolder}/PageRes.xml"] = ToUtf8Bytes(pageRes);
        }
    }

    private static void BuildAttachments(OfdDocumentPackage package, IDictionary<string, byte[]> entries, XNamespace ns, string docId)
    {
        if (package.Attachments.Count == 0)
        {
            return;
        }

        var attachments = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Attachments",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                package.Attachments.Select((attachment, index) =>
                {
                    var fileName = $"Attach_{index + 1}_{SanitizeFileName(attachment.Name)}";
                    if (!attachment.IsExternal)
                    {
                        entries[$"{docId}/Attachs/{fileName}"] = attachment.Data;
                    }

                    return new XElement(ns + "Attachment",
                        new XAttribute("ID", index + 1),
                        new XAttribute("Name", attachment.Name),
                        new XAttribute("MediaType", attachment.MediaType),
                        new XAttribute("External", attachment.IsExternal),
                        new XAttribute("FileLoc", attachment.IsExternal ? attachment.ExternalPath ?? string.Empty : fileName));
                })));

        entries[$"{docId}/Attachs/Attachments.xml"] = ToUtf8Bytes(attachments);
    }

    private static void BuildCustomTags(OfdDocumentPackage package, IDictionary<string, byte[]> entries, XNamespace ns, string docId)
    {
        if (package.CustomTags.Count == 0)
        {
            return;
        }

        var list = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "CustomTags",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "CustomTag",
                    new XAttribute("ID", "EMR"),
                    new XAttribute("FileLoc", "CustomTag_EMR.xml"))));

        entries[$"{docId}/Tags/CustomTags.xml"] = ToUtf8Bytes(list);

        var detail = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "EMRTags",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                package.CustomTags.Select(x => new XElement(ns + "Tag", new XAttribute("Key", x.Key), new XAttribute("Value", x.Value)))));

        entries[$"{docId}/Tags/CustomTag_EMR.xml"] = ToUtf8Bytes(detail);
    }

    private static IEnumerable<XElement> BuildDocInfo(XNamespace ns, OfdMetadata metadata)
    {
        if (!string.IsNullOrWhiteSpace(metadata.Title))
        {
            yield return new XElement(ns + "Title", metadata.Title);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Author))
        {
            yield return new XElement(ns + "Author", metadata.Author);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Subject))
        {
            yield return new XElement(ns + "Subject", metadata.Subject);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Keywords))
        {
            yield return new XElement(ns + "Keywords", metadata.Keywords);
        }

        if (!string.IsNullOrWhiteSpace(metadata.Creator))
        {
            yield return new XElement(ns + "Creator", metadata.Creator);
        }

        if (metadata.CreationDate.HasValue)
        {
            yield return new XElement(ns + "CreationDate", metadata.CreationDate.Value.ToString("O", CultureInfo.InvariantCulture));
        }

        if (metadata.ModificationDate.HasValue)
        {
            yield return new XElement(ns + "ModDate", metadata.ModificationDate.Value.ToString("O", CultureInfo.InvariantCulture));
        }
    }

    private static string BuildBox(double x, double y, double w, double h)
    {
        return string.Format(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} {3:0.###}", x, y, w, h);
    }

    private static string ToInvariant(double value)
    {
        return value.ToString("0.###", CultureInfo.InvariantCulture);
    }

    private static string SanitizeFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Attachment.bin";
        }

        var invalid = Path.GetInvalidFileNameChars();
        var chars = value.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        return new string(chars);
    }

    private static string GetImageExtension(string mediaType)
    {
        return mediaType switch
        {
            "image/jpeg" => ".jpg",
            "image/bmp" => ".bmp",
            "image/tiff" => ".tiff",
            _ => ".png"
        };
    }

    private static byte[] ToUtf8Bytes(XDocument xml)
    {
        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, new UTF8Encoding(false), 1024, leaveOpen: true))
        {
            xml.Save(writer);
        }

        return ms.ToArray();
    }
}
