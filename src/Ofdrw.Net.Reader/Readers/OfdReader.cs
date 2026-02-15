using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Interfaces;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Reader.Readers;

public sealed class OfdReader : IOfdReader
{
    private readonly OfdPackageLoader _loader = new();

    public async Task<OfdDocumentPackage> ReadAsync(Stream ofdStream, CancellationToken cancellationToken = default)
    {
        var archive = await _loader.LoadAsync(ofdStream, cancellationToken).ConfigureAwait(false);
        var package = new OfdDocumentPackage();

        var ofdXml = XDocument.Parse(archive.ReadUtf8Text(OfdConstants.OfdRootFile));
        var ofdNs = ofdXml.Root?.Name.Namespace ?? XNamespace.Get(OfdConstants.Namespace);
        var docType = ofdXml.Root?.Attribute("DocType")?.Value ?? OfdConstants.DefaultDocType;

        var docRoot = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .Elements(ofdNs + "DocRoot")
            .Select(x => x.Value)
            .FirstOrDefault() ?? "Doc_0/Document.xml";

        package.Options.DocType = docType;
        package.Options.Namespace = ofdNs.NamespaceName;
        package.Options.DocumentId = docRoot.Split('/')[0];

        var info = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .Elements(ofdNs + "DocInfo")
            .FirstOrDefault();

        if (info is not null)
        {
            package.Options.Metadata.Title = info.Element(ofdNs + "Title")?.Value;
            package.Options.Metadata.Author = info.Element(ofdNs + "Author")?.Value;
            package.Options.Metadata.Subject = info.Element(ofdNs + "Subject")?.Value;
            package.Options.Metadata.Keywords = info.Element(ofdNs + "Keywords")?.Value;
            package.Options.Metadata.Creator = info.Element(ofdNs + "Creator")?.Value;
        }

        var documentXml = XDocument.Parse(archive.ReadUtf8Text(docRoot));
        var docNs = documentXml.Root?.Name.Namespace ?? ofdNs;

        var pages = documentXml.Root?
            .Element(docNs + "Pages")?
            .Elements(docNs + "Page")
            .Select((x, index) => new
            {
                Index = index,
                BaseLoc = x.Attribute("BaseLoc")?.Value
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.BaseLoc))
            .ToList() ?? [];

        foreach (var pageRef in pages)
        {
            var contentPath = Resolve(docRoot, pageRef.BaseLoc!);
            var pageXml = XDocument.Parse(archive.ReadUtf8Text(contentPath));
            var pageNs = pageXml.Root?.Name.Namespace ?? docNs;

            var box = ParseBox(pageXml.Root?
                .Element(pageNs + "Area")?
                .Element(pageNs + "PhysicalBox")?.Value);

            var page = new OfdPage
            {
                Index = pageRef.Index,
                WidthMillimeters = box.w,
                HeightMillimeters = box.h
            };

            var pageDir = GetDirectory(contentPath);
            var pageResPath = $"{pageDir}/PageRes.xml";
            var mediaMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var mediaTypeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (archive.Contains(pageResPath))
            {
                var pageResXml = XDocument.Parse(archive.ReadUtf8Text(pageResPath));
                var pageResNs = pageResXml.Root?.Name.Namespace ?? docNs;
                foreach (var media in pageResXml.Root?.Elements(pageResNs + "MultiMedia") ?? Enumerable.Empty<XElement>())
                {
                    var id = media.Attribute("ID")?.Value;
                    var file = media.Attribute("MediaFile")?.Value;
                    if (!string.IsNullOrWhiteSpace(id) && !string.IsNullOrWhiteSpace(file))
                    {
                        mediaMap[id] = file;
                        mediaTypeMap[id] = media.Attribute("Format")?.Value ?? "image/png";
                    }
                }
            }

            var layer = pageXml.Root?
                .Element(pageNs + "Content")?
                .Element(pageNs + "Layer");

            if (layer is not null)
            {
                foreach (var node in layer.Elements())
                {
                    var localName = node.Name.LocalName;
                    if (string.Equals(localName, "TextObject", StringComparison.OrdinalIgnoreCase))
                    {
                        var boundary = ParseBox(node.Attribute("Boundary")?.Value);
                        var text = new OfdTextElement
                        {
                            XMillimeters = boundary.x,
                            YMillimeters = boundary.y,
                            WidthMillimeters = boundary.w,
                            HeightMillimeters = boundary.h,
                            Text = node.Value,
                            FontName = node.Attribute("Font")?.Value ?? "SimSun",
                            FontSizeMillimeters = ParseDouble(node.Attribute("Size")?.Value, 4d)
                        };
                        page.Elements.Add(text);
                    }

                    if (string.Equals(localName, "ImageObject", StringComparison.OrdinalIgnoreCase))
                    {
                        var boundary = ParseBox(node.Attribute("Boundary")?.Value);
                        var resourceId = node.Attribute("ResourceID")?.Value ?? string.Empty;
                        var image = new OfdImageElement
                        {
                            XMillimeters = boundary.x,
                            YMillimeters = boundary.y,
                            WidthMillimeters = boundary.w,
                            HeightMillimeters = boundary.h,
                            ResourceId = resourceId,
                            MediaType = mediaTypeMap.TryGetValue(resourceId, out var mediaType) ? mediaType : "image/png"
                        };

                        if (mediaMap.TryGetValue(resourceId, out var mediaFile))
                        {
                            var mediaPath = Resolve(pageResPath, mediaFile);
                            if (archive.TryGetBytes(mediaPath, out var bytes))
                            {
                                image.Data = bytes;
                                image.FileName = Path.GetFileName(mediaPath);
                            }
                        }

                        page.Elements.Add(image);
                    }
                }
            }

            package.Pages.Add(page);
        }

        var attachmentsPath = $"{package.Options.DocumentId}/Attachs/Attachments.xml";
        if (archive.Contains(attachmentsPath))
        {
            var attachmentsXml = XDocument.Parse(archive.ReadUtf8Text(attachmentsPath));
            var attachNs = attachmentsXml.Root?.Name.Namespace ?? ofdNs;
            foreach (var node in attachmentsXml.Root?.Elements(attachNs + "Attachment") ?? Enumerable.Empty<XElement>())
            {
                var external = bool.TryParse(node.Attribute("External")?.Value, out var value) && value;
                var fileLoc = node.Attribute("FileLoc")?.Value ?? string.Empty;

                var attachment = new OfdAttachment
                {
                    Name = node.Attribute("Name")?.Value ?? "Attachment",
                    MediaType = node.Attribute("MediaType")?.Value ?? "application/octet-stream",
                    IsExternal = external,
                    ExternalPath = external ? fileLoc : null
                };

                if (!external)
                {
                    var contentPath = Resolve(attachmentsPath, fileLoc);
                    if (archive.TryGetBytes(contentPath, out var bytes))
                    {
                        attachment.Data = bytes;
                    }
                }

                package.Attachments.Add(attachment);
            }
        }

        var customTagPath = $"{package.Options.DocumentId}/Tags/CustomTag_EMR.xml";
        if (archive.Contains(customTagPath))
        {
            var tagsXml = XDocument.Parse(archive.ReadUtf8Text(customTagPath));
            var tagNs = tagsXml.Root?.Name.Namespace ?? ofdNs;
            foreach (var tag in tagsXml.Root?.Elements(tagNs + "Tag") ?? Enumerable.Empty<XElement>())
            {
                var key = tag.Attribute("Key")?.Value;
                var value = tag.Attribute("Value")?.Value;
                if (!string.IsNullOrWhiteSpace(key) && value is not null)
                {
                    package.CustomTags[key] = value;
                }
            }
        }

        return package;
    }

    private static string Resolve(string basePath, string relativePath)
    {
        var root = GetDirectory(basePath);
        if (relativePath.StartsWith("/", StringComparison.Ordinal))
        {
            return relativePath.TrimStart('/');
        }

        var combined = $"{root}/{relativePath}";
        var segments = new Stack<string>();
        foreach (var segment in combined.Split('/'))
        {
            if (string.IsNullOrWhiteSpace(segment) || segment == ".")
            {
                continue;
            }

            if (segment == "..")
            {
                if (segments.Count > 0)
                {
                    segments.Pop();
                }

                continue;
            }

            segments.Push(segment);
        }

        return string.Join("/", segments.Reverse());
    }

    private static string GetDirectory(string path)
    {
        var normalized = path.Replace('\\', '/').TrimStart('/');
        var index = normalized.LastIndexOf('/');
        return index < 0 ? string.Empty : normalized.Substring(0, index);
    }

    private static (double x, double y, double w, double h) ParseBox(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return (0d, 0d, 0d, 0d);
        }

        var parts = value.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length < 4)
        {
            return (0d, 0d, 0d, 0d);
        }

        return (
            ParseDouble(parts[0], 0d),
            ParseDouble(parts[1], 0d),
            ParseDouble(parts[2], 0d),
            ParseDouble(parts[3], 0d));
    }

    private static double ParseDouble(string? value, double fallback)
    {
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : fallback;
    }
}
