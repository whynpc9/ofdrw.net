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
                new XAttribute("Version", "1.0"),
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

        var fontIds = BuildFontIds(package);
        var imageResources = BuildImageResources(orderedPages);
        var maxUnitId = Math.Max(orderedPages.Count, 1) + orderedPages.Sum(x => x.Elements.Count) + fontIds.Count + imageResources.Count + 1;

        var docXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Document",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "CommonData",
                    new XElement(ns + "MaxUnitID", maxUnitId),
                    new XElement(ns + "PageArea",
                        new XElement(ns + "PhysicalBox", BuildBox(0, 0, orderedPages[0].WidthMillimeters, orderedPages[0].HeightMillimeters))
                    ),
                    imageResources.Count > 0
                        ? new XElement(ns + "DocumentRes", "DocumentRes.xml")
                        : null,
                    new XElement(ns + "PublicRes", "PublicRes.xml")
                ),
                new XElement(ns + "Pages",
                    orderedPages.Select((page, i) =>
                        new XElement(ns + "Page",
                            new XAttribute("ID", i + 1),
                            new XAttribute("BaseLoc", $"Pages/Page_{i}/Content.xml"))
                    )
                ),
                package.Attachments.Count > 0
                    ? new XElement(ns + "Attachments", "Attachs/Attachments.xml")
                    : null));

        entries[$"{docId}/Document.xml"] = ToUtf8Bytes(docXml);
        if (imageResources.Count > 0)
        {
            BuildDocumentResources(entries, ns, docId, imageResources);
        }

        BuildPublicResources(entries, ns, docId, fontIds);
        BuildPages(orderedPages, entries, ns, docId, fontIds, imageResources);
        BuildAttachments(package, entries, ns, docId);
        BuildCustomTags(package, entries, ns, docId);

        return entries;
    }

    private static Dictionary<string, int> BuildFontIds(OfdDocumentPackage package)
    {
        return package.Pages.SelectMany(p => p.Elements)
            .OfType<OfdTextElement>()
            .Select(t => t.FontName)
            .Where(x => !string.IsNullOrWhiteSpace(x))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Select((font, index) => new { font, id = index + 1 })
            .ToDictionary(x => x.font, x => x.id, StringComparer.OrdinalIgnoreCase);
    }

    private sealed class ImageResource
    {
        public OfdImageElement Image { get; set; } = new();

        public string Id { get; set; } = string.Empty;

        public string FileName { get; set; } = string.Empty;

        public string Format { get; set; } = "PNG";
    }

    private static Dictionary<OfdImageElement, ImageResource> BuildImageResources(IReadOnlyList<OfdPage> pages)
    {
        var resources = new Dictionary<OfdImageElement, ImageResource>();
        var nextId = 1;

        foreach (var image in pages.SelectMany(x => x.Elements).OfType<OfdImageElement>())
        {
            var id = string.IsNullOrWhiteSpace(image.ResourceId) ? nextId.ToString(CultureInfo.InvariantCulture) : image.ResourceId;
            var format = ToOfdImageFormat(image.MediaType, image.FileName);
            var extension = GetImageExtension(format);
            var fileName = string.IsNullOrWhiteSpace(image.FileName) ? $"Image_{id}{extension}" : image.FileName;
            resources[image] = new ImageResource
            {
                Image = image,
                Id = id,
                FileName = fileName,
                Format = format
            };
            nextId++;
        }

        return resources;
    }

    private static void BuildDocumentResources(IDictionary<string, byte[]> entries, XNamespace ns, string docId, IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources)
    {
        var mediaElements = imageResources.Values.Select(r =>
            new XElement(ns + "MultiMedia",
                new XAttribute("ID", r.Id),
                new XAttribute("Type", "Image"),
                new XAttribute("Format", r.Format),
                new XElement(ns + "MediaFile", r.FileName)));

        var documentResXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Res",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XAttribute("BaseLoc", "Res"),
                new XElement(ns + "MultiMedias", mediaElements)));

        entries[$"{docId}/DocumentRes.xml"] = ToUtf8Bytes(documentResXml);
        foreach (var resource in imageResources.Values)
        {
            entries[$"{docId}/Res/{resource.FileName}"] = resource.Image.Data;
        }
    }

    private static void BuildPublicResources(IDictionary<string, byte[]> entries, XNamespace ns, string docId, IReadOnlyDictionary<string, int> fontIds)
    {
        var fontElements = fontIds.Select(font =>
            new XElement(ns + "Font",
                new XAttribute("ID", font.Value),
                new XAttribute("FontName", font.Key),
                new XAttribute("FamilyName", font.Key)));

        var publicResXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Res",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XAttribute("BaseLoc", "Res"),
                new XElement(ns + "Fonts", fontElements)));

        entries[$"{docId}/PublicRes.xml"] = ToUtf8Bytes(publicResXml);
    }

    private static void BuildPages(
        List<OfdPage> pages,
        IDictionary<string, byte[]> entries,
        XNamespace ns,
        string docId,
        IReadOnlyDictionary<string, int> fontIds,
        IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var page = pages[i];
            var pageFolder = $"{docId}/Pages/Page_{i}";
            var layer = new XElement(ns + "Layer", new XAttribute("ID", 1), new XAttribute("Type", "Body"));
            var elementId = 1;

            foreach (var element in page.Elements)
            {
                if (element is OfdTextElement text)
                {
                    var width = text.WidthMillimeters <= 0 ? Math.Max(text.Text.Length * text.FontSizeMillimeters * 0.5, text.FontSizeMillimeters) : text.WidthMillimeters;
                    var height = text.HeightMillimeters <= 0 ? text.FontSizeMillimeters * 1.5 : text.HeightMillimeters;
                    var fontId = fontIds.TryGetValue(text.FontName, out var resolvedFontId) ? resolvedFontId : 1;
                    var textObject = new XElement(ns + "TextObject",
                        new XAttribute("ID", elementId++),
                        new XAttribute("Boundary", BuildBox(text.XMillimeters, text.YMillimeters, width, height)),
                        new XAttribute("Font", fontId),
                        new XAttribute("Size", ToInvariant(text.FontSizeMillimeters)),
                        new XElement(ns + "TextCode",
                            new XAttribute("X", "0"),
                            new XAttribute("Y", ToInvariant(Math.Max(text.FontSizeMillimeters, 1d))),
                            text.Text));
                    layer.Add(textObject);
                }
                else if (element is OfdImageElement image)
                {
                    var resource = imageResources[image];

                    var imageObject = new XElement(ns + "ImageObject",
                        new XAttribute("ID", elementId++),
                        new XAttribute("Boundary", BuildBox(image.XMillimeters, image.YMillimeters, image.WidthMillimeters, image.HeightMillimeters)),
                        new XAttribute("CTM", BuildMatrix(image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0)),
                        new XAttribute("ResourceID", resource.Id));
                    layer.Add(imageObject);
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

    private static string BuildMatrix(double a, double b, double c, double d, double e, double f)
    {
        return string.Format(CultureInfo.InvariantCulture, "{0:0.###} {1:0.###} {2:0.###} {3:0.###} {4:0.###} {5:0.###}", a, b, c, d, e, f);
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

    private static string GetImageExtension(string format)
    {
        return format switch
        {
            "JPEG" => ".jpg",
            "JPG" => ".jpg",
            "BMP" => ".bmp",
            "TIFF" => ".tiff",
            _ => ".png"
        };
    }

    private static string ToOfdImageFormat(string mediaType, string fileName)
    {
        var normalizedMediaType = mediaType?.Trim().ToLowerInvariant();
        if (normalizedMediaType == "image/jpeg" || normalizedMediaType == "image/jpg")
        {
            return "JPEG";
        }

        if (normalizedMediaType == "image/bmp")
        {
            return "BMP";
        }

        if (normalizedMediaType == "image/tiff")
        {
            return "TIFF";
        }

        var extension = Path.GetExtension(fileName)?.TrimStart('.').ToUpperInvariant();
        return extension switch
        {
            "JPG" => "JPEG",
            "JPEG" => "JPEG",
            "BMP" => "BMP",
            "TIFF" => "TIFF",
            _ => "PNG"
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
