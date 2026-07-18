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
        var entries = new Dictionary<string, byte[]>(package.PreservedEntries, StringComparer.OrdinalIgnoreCase);
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
                    new XElement(ns + "DocRoot", $"{docId}/Document.xml"),
                    package.PreservedDocBodyElements.Select(xml =>
                        XElement.Parse(xml, LoadOptions.PreserveWhitespace))
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

        var idAllocator = new OfdIdAllocator();
        idAllocator.AdvancePast(GetPreservedMaxId(package.PreservedEntries));
        var pageIds = orderedPages.ToDictionary(page => page, page => idAllocator.AllocatePreferred(page.Id));
        var elementIds = new Dictionary<OfdElement, string>();
        var layerIds = BuildPageObjectIds(orderedPages, elementIds, idAllocator);
        var imageResources = BuildImageResources(orderedPages, idAllocator);
        var fontIds = BuildFontIds(package, idAllocator);
        var publicResourceLocation = package.PublicResourceLocation ?? "PublicRes.xml";
        var documentResourceLocation = package.DocumentResourceLocation ??
            (imageResources.Count > 0 ? "DocumentRes.xml" : null);
        var newImageResources = imageResources
            .Where(pair => string.IsNullOrWhiteSpace(pair.Key.SourceXml) ||
                string.IsNullOrWhiteSpace(package.DocumentResourceLocation))
            .ToDictionary(pair => pair.Key, pair => pair.Value);

        if (newImageResources.Count > 0 && !string.IsNullOrWhiteSpace(documentResourceLocation))
        {
            BuildDocumentResources(
                entries,
                ns,
                docId,
                documentResourceLocation!,
                newImageResources);
        }

        var fontNamesToAdd = new HashSet<string>(package.Pages
            .SelectMany(page => page.Elements)
            .OfType<OfdTextElement>()
            .GroupBy(text => text.FontName, StringComparer.OrdinalIgnoreCase)
            .Where(group => string.IsNullOrWhiteSpace(package.PublicResourceLocation) ||
                group.All(text => string.IsNullOrWhiteSpace(text.SourceXml)))
            .Select(group => group.Key), StringComparer.OrdinalIgnoreCase);
        foreach (var font in package.Fonts.Where(font => font.Data.Length > 0))
        {
            fontNamesToAdd.Add(font.FontName);
        }

        BuildPublicResources(
            entries,
            ns,
            docId,
            publicResourceLocation,
            fontIds,
            fontNamesToAdd,
            package.Fonts);
        BuildPages(orderedPages, entries, ns, docId, pageIds, layerIds, elementIds, fontIds, imageResources, idAllocator);
        BuildAttachments(package, entries, ns, docId);
        BuildCustomTags(package, entries, ns, docId);

        var docXml = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Document",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "CommonData",
                    new XElement(ns + "MaxUnitID", idAllocator.MaxId),
                    new XElement(ns + "PageArea",
                        new XElement(ns + "PhysicalBox", BuildBox(
                            orderedPages[0].XMillimeters,
                            orderedPages[0].YMillimeters,
                            orderedPages[0].WidthMillimeters,
                            orderedPages[0].HeightMillimeters))
                    ),
                    !string.IsNullOrWhiteSpace(documentResourceLocation)
                        ? new XElement(ns + "DocumentRes", documentResourceLocation)
                        : null,
                    new XElement(ns + "PublicRes", publicResourceLocation),
                    package.PreservedCommonDataElements.Select(xml =>
                        XElement.Parse(xml, LoadOptions.PreserveWhitespace))
                ),
                new XElement(ns + "Pages",
                    orderedPages.Select((page, i) =>
                        new XElement(ns + "Page",
                            new XAttribute("ID", pageIds[page]),
                            new XAttribute("BaseLoc", $"Pages/Page_{i}/Content.xml"))
                    )
                ),
                package.Attachments.Count > 0
                    ? new XElement(ns + "Attachments", "Attachs/Attachments.xml")
                    : null,
                package.CustomTags.Count > 0
                    ? new XElement(ns + "CustomTags", "Tags/CustomTags.xml")
                    : null,
                package.PreservedDocumentElements.Select(xml =>
                    XElement.Parse(xml, LoadOptions.PreserveWhitespace))));

        entries[$"{docId}/Document.xml"] = ToUtf8Bytes(docXml);
        return entries;
    }

    private static Dictionary<string, string> BuildFontIds(OfdDocumentPackage package, OfdIdAllocator idAllocator)
    {
        return package.Pages.SelectMany(p => p.Elements)
            .OfType<OfdTextElement>()
            .Where(t => !string.IsNullOrWhiteSpace(t.FontName))
            .GroupBy(t => t.FontName, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(
                group => group.Key,
                group => idAllocator.AllocatePreferred(
                    group.Select(text => text.FontResourceId)
                        .FirstOrDefault(id => !string.IsNullOrWhiteSpace(id))
                    ?? package.Fonts.FirstOrDefault(font => string.Equals(
                        font.FontName,
                        group.Key,
                        StringComparison.OrdinalIgnoreCase))?.Id),
                StringComparer.OrdinalIgnoreCase);
    }

    private sealed class ImageResource
    {
        public OfdImageElement Image { get; set; } = new();

        public string Id { get; set; } = string.Empty;

        public string FileName { get; set; } = string.Empty;

        public string Format { get; set; } = "PNG";
    }

    private static Dictionary<OfdPage, Dictionary<string, string>> BuildPageObjectIds(
        IReadOnlyList<OfdPage> pages,
        IDictionary<OfdElement, string> elementIds,
        OfdIdAllocator idAllocator)
    {
        var result = new Dictionary<OfdPage, Dictionary<string, string>>();
        foreach (var page in pages)
        {
            var pageLayerIds = new Dictionary<string, string>(StringComparer.Ordinal);
            foreach (var group in page.Elements.GroupBy(GetLayerKey))
            {
                var preferredLayerId = group.First().LayerId;
                pageLayerIds[group.Key] = idAllocator.AllocatePreferred(preferredLayerId);
                foreach (var element in group)
                {
                    elementIds[element] = idAllocator.AllocatePreferred(element.ObjectId);
                }
            }

            result[page] = pageLayerIds;
        }

        return result;
    }

    private static Dictionary<OfdImageElement, ImageResource> BuildImageResources(
        IReadOnlyList<OfdPage> pages,
        OfdIdAllocator idAllocator)
    {
        var resources = new Dictionary<OfdImageElement, ImageResource>();

        foreach (var image in pages.SelectMany(x => x.Elements).OfType<OfdImageElement>())
        {
            var id = idAllocator.AllocatePreferred(image.ResourceId);
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
        }

        return resources;
    }

    private static void BuildDocumentResources(
        IDictionary<string, byte[]> entries,
        XNamespace ns,
        string docId,
        string resourceLocation,
        IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources)
    {
        var resourcePath = ResolveEntryPath($"{docId}/Document.xml", resourceLocation);
        XDocument documentResXml;
        if (entries.TryGetValue(resourcePath, out var existingBytes))
        {
            using var existingStream = new MemoryStream(existingBytes, writable: false);
            documentResXml = XDocument.Load(existingStream, LoadOptions.PreserveWhitespace);
        }
        else
        {
            documentResXml = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement(ns + "Res",
                    new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                    new XAttribute("BaseLoc", "Res")));
        }

        var root = documentResXml.Root ?? throw new InvalidDataException("Document resource XML has no root element.");
        var resourceNs = root.Name.Namespace;
        var mediaContainer = root.Elements()
            .FirstOrDefault(element => element.Name.LocalName == "MultiMedias");
        if (mediaContainer is null)
        {
            mediaContainer = new XElement(resourceNs + "MultiMedias");
            root.Add(mediaContainer);
        }

        var existingIds = new HashSet<string>(root.Descendants()
            .Where(element => element.Name.LocalName == "MultiMedia")
            .Select(element => element.Attribute("ID")?.Value)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!), StringComparer.Ordinal);
        foreach (var resource in imageResources.Values)
        {
            if (!existingIds.Contains(resource.Id))
            {
                mediaContainer.Add(new XElement(resourceNs + "MultiMedia",
                    new XAttribute("ID", resource.Id),
                    new XAttribute("Type", "Image"),
                    new XAttribute("Format", resource.Format),
                    new XElement(resourceNs + "MediaFile", resource.FileName)));
            }
        }

        entries[resourcePath] = ToUtf8Bytes(documentResXml);
        var baseLoc = root.Attribute("BaseLoc")?.Value ?? "Res";
        foreach (var resource in imageResources.Values)
        {
            entries[ResolveEntryPath(resourcePath, $"{baseLoc.TrimEnd('/')}/{resource.FileName}")] =
                resource.Image.Data;
        }
    }

    private static void BuildPublicResources(
        IDictionary<string, byte[]> entries,
        XNamespace ns,
        string docId,
        string resourceLocation,
        IReadOnlyDictionary<string, string> fontIds,
        ISet<string> fontNamesToAdd,
        IReadOnlyList<OfdFontResource> fontResources)
    {
        var resourcePath = ResolveEntryPath($"{docId}/Document.xml", resourceLocation);
        XDocument publicResXml;
        if (entries.TryGetValue(resourcePath, out var existingBytes))
        {
            using var existingStream = new MemoryStream(existingBytes, writable: false);
            publicResXml = XDocument.Load(existingStream, LoadOptions.PreserveWhitespace);
        }
        else
        {
            publicResXml = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                new XElement(ns + "Res",
                    new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                    new XAttribute("BaseLoc", "Res")));
        }

        var root = publicResXml.Root ?? throw new InvalidDataException("Public resource XML has no root element.");
        var resourceNs = root.Name.Namespace;
        var fonts = root.Elements().FirstOrDefault(element => element.Name.LocalName == "Fonts");
        if (fonts is null)
        {
            fonts = new XElement(resourceNs + "Fonts");
            root.AddFirst(fonts);
        }

        var existingIds = new HashSet<string>(root.Descendants()
            .Where(element => element.Name.LocalName == "Font")
            .Select(element => element.Attribute("ID")?.Value)
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .Select(id => id!), StringComparer.Ordinal);
        foreach (var font in fontIds.Where(font => fontNamesToAdd.Contains(font.Key)))
        {
            if (!existingIds.Contains(font.Value))
            {
                var embedded = fontResources.FirstOrDefault(resource => string.Equals(
                    resource.FontName,
                    font.Key,
                    StringComparison.OrdinalIgnoreCase));
                var fileName = embedded?.Data.Length > 0
                    ? embedded.FileName ?? $"Font_{font.Value}.ttf"
                    : null;
                fonts.Add(new XElement(resourceNs + "Font",
                    new XAttribute("ID", font.Value),
                    new XAttribute("FontName", font.Key),
                    new XAttribute("FamilyName", embedded?.FamilyName ?? font.Key),
                    embedded is not null && !string.IsNullOrWhiteSpace(embedded.Charset)
                        ? new XAttribute("Charset", embedded.Charset)
                        : null,
                    embedded?.Bold == true ? new XAttribute("Bold", true) : null,
                    embedded?.Italic == true ? new XAttribute("Italic", true) : null,
                    fileName is not null
                        ? new XElement(resourceNs + "FontFile", fileName)
                        : null));
                if (fileName is not null && embedded is not null)
                {
                    var baseLoc = root.Attribute("BaseLoc")?.Value ?? "Res";
                    entries[ResolveEntryPath(
                        resourcePath,
                        $"{baseLoc.TrimEnd('/')}/{fileName}")] = embedded.Data;
                }
            }
        }

        entries[resourcePath] = ToUtf8Bytes(publicResXml);
    }

    private static void BuildPages(
        List<OfdPage> pages,
        IDictionary<string, byte[]> entries,
        XNamespace ns,
        string docId,
        IReadOnlyDictionary<OfdPage, string> pageIds,
        IReadOnlyDictionary<OfdPage, Dictionary<string, string>> layerIds,
        IReadOnlyDictionary<OfdElement, string> elementIds,
        IReadOnlyDictionary<string, string> fontIds,
        IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources,
        OfdIdAllocator idAllocator)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var page = pages[i];
            var pageFolder = $"{docId}/Pages/Page_{i}";
            var content = new XElement(ns + "Content");

            var layerGroups = page.Elements.GroupBy(GetLayerKey);

            foreach (var layerGroup in layerGroups)
            {
                var firstElement = layerGroup.First();
                var layerId = layerIds[page][layerGroup.Key];
                var layerType = string.IsNullOrWhiteSpace(firstElement.LayerType) ? "Body" : firstElement.LayerType;
                var layer = new XElement(
                    ns + "Layer",
                    new XAttribute("ID", layerId),
                    new XAttribute("Type", layerType));

                foreach (var element in layerGroup)
                {
                    var objectId = elementIds[element];
                    if (element is OfdTextElement text)
                    {
                        var fontId = fontIds.TryGetValue(text.FontName, out var resolvedFontId)
                            ? resolvedFontId
                            : idAllocator.Allocate();
                        XElement textObject;
                        if (!string.IsNullOrWhiteSpace(text.SourceXml))
                        {
                            textObject = XElement.Parse(text.SourceXml!, LoadOptions.PreserveWhitespace);
                            textObject.SetAttributeValue("ID", objectId);
                            textObject.SetAttributeValue("Font", fontId);
                        }
                        else
                        {
                            var width = text.WidthMillimeters <= 0 ? Math.Max(text.Text.Length * text.FontSizeMillimeters * 0.5, text.FontSizeMillimeters) : text.WidthMillimeters;
                            var height = text.HeightMillimeters <= 0 ? text.FontSizeMillimeters * 1.5 : text.HeightMillimeters;
                            textObject = new XElement(ns + "TextObject",
                                new XAttribute("ID", objectId),
                                new XAttribute("Boundary", BuildBox(text.XMillimeters, text.YMillimeters, width, height)),
                                new XAttribute("Font", fontId),
                                new XAttribute("Size", ToInvariant(text.FontSizeMillimeters)),
                                text.Transform is { Length: 6 }
                                    ? new XAttribute("CTM", string.Join(" ", text.Transform.Select(ToInvariant)))
                                    : null,
                                text.FillColor.Red != 0 ||
                                text.FillColor.Green != 0 ||
                                text.FillColor.Blue != 0 ||
                                text.FillColor.Alpha != 255
                                    ? new XElement(ns + "FillColor",
                                        new XAttribute(
                                            "Value",
                                            $"{text.FillColor.Red} {text.FillColor.Green} {text.FillColor.Blue}"),
                                        text.FillColor.Alpha != 255
                                            ? new XAttribute("Alpha", text.FillColor.Alpha)
                                            : null)
                                    : null,
                                text.Runs.Count > 0
                                    ? text.Runs.Select(run => new XElement(ns + "TextCode",
                                        new XAttribute("X", ToInvariant(run.XMillimeters)),
                                        new XAttribute("Y", ToInvariant(run.YMillimeters)),
                                        !string.IsNullOrWhiteSpace(run.DeltaX)
                                            ? new XAttribute("DeltaX", run.DeltaX)
                                            : null,
                                        !string.IsNullOrWhiteSpace(run.DeltaY)
                                            ? new XAttribute("DeltaY", run.DeltaY)
                                            : null,
                                        run.Text))
                                    : new[]
                                    {
                                        new XElement(ns + "TextCode",
                                            new XAttribute("X", "0"),
                                            new XAttribute("Y", ToInvariant(Math.Max(text.FontSizeMillimeters, 1d))),
                                            text.Text)
                                    });
                        }

                        layer.Add(textObject);
                    }
                    else if (element is OfdImageElement image)
                    {
                        var resource = imageResources[image];
                        XElement imageObject;
                        if (!string.IsNullOrWhiteSpace(image.SourceXml))
                        {
                            imageObject = XElement.Parse(image.SourceXml!, LoadOptions.PreserveWhitespace);
                            imageObject.SetAttributeValue("ID", objectId);
                            imageObject.SetAttributeValue("ResourceID", resource.Id);
                        }
                        else
                        {
                            imageObject = new XElement(ns + "ImageObject",
                                new XAttribute("ID", objectId),
                                new XAttribute("Boundary", BuildBox(image.XMillimeters, image.YMillimeters, image.WidthMillimeters, image.HeightMillimeters)),
                                new XAttribute("CTM", BuildMatrix(image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0)),
                                new XAttribute("ResourceID", resource.Id));
                        }

                        layer.Add(imageObject);
                    }
                    else if (element is OfdPathElement path)
                    {
                        var pathObject = string.IsNullOrWhiteSpace(path.SourceXml)
                            ? new XElement(ns + "PathObject")
                            : XElement.Parse(path.SourceXml!, LoadOptions.PreserveWhitespace);
                        pathObject.SetAttributeValue("ID", objectId);
                        pathObject.SetAttributeValue(
                            "Boundary",
                            BuildBox(path.XMillimeters, path.YMillimeters, path.WidthMillimeters, path.HeightMillimeters));
                        pathObject.SetAttributeValue("LineWidth", ToInvariant(path.LineWidthMillimeters));
                        pathObject.SetAttributeValue("Stroke", path.Stroke);
                        pathObject.SetAttributeValue("Fill", path.Fill);
                        pathObject.SetAttributeValue(
                            "CTM",
                            path.Transform is { Length: 6 }
                                ? string.Join(" ", path.Transform.Select(ToInvariant))
                                : null);
                        SetPathColor(pathObject, ns, "StrokeColor", path.Stroke ? path.StrokeColor : null);
                        SetPathColor(pathObject, ns, "FillColor", path.Fill ? path.FillColor : null);

                        var abbreviatedData = pathObject.Elements()
                            .FirstOrDefault(x => x.Name.LocalName == "AbbreviatedData");
                        if (abbreviatedData is null)
                        {
                            pathObject.Add(new XElement(ns + "AbbreviatedData", path.AbbreviatedData));
                        }
                        else
                        {
                            abbreviatedData.Value = path.AbbreviatedData;
                        }

                        layer.Add(pathObject);
                    }
                    else if (element is OfdRawElement raw && !string.IsNullOrWhiteSpace(raw.Xml))
                    {
                        var rawElement = XElement.Parse(raw.Xml, LoadOptions.PreserveWhitespace);
                        rawElement.SetAttributeValue("ID", objectId);
                        layer.Add(rawElement);
                    }
                }

                content.Add(layer);
            }

            if (!content.HasElements)
            {
                content.Add(new XElement(
                    ns + "Layer",
                    new XAttribute("ID", idAllocator.Allocate()),
                    new XAttribute("Type", "Body")));
            }

            var pageRoot = new XElement(
                ns + "Page",
                new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName),
                new XElement(ns + "Area", new XElement(ns + "PhysicalBox", BuildBox(
                    page.XMillimeters,
                    page.YMillimeters,
                    page.WidthMillimeters,
                    page.HeightMillimeters))));

            foreach (var preservedXml in page.PreservedPageElements)
            {
                if (!string.IsNullOrWhiteSpace(preservedXml))
                {
                    pageRoot.Add(XElement.Parse(preservedXml, LoadOptions.PreserveWhitespace));
                }
            }

            pageRoot.Add(content);
            var pageDocument = new XDocument(
                new XDeclaration("1.0", "UTF-8", null),
                pageRoot);
            entries[$"{pageFolder}/Content.xml"] = ToUtf8Bytes(pageDocument);
        }
    }

    private static string GetLayerKey(OfdElement element)
    {
        var id = element.LayerId ?? string.Empty;
        var type = string.IsNullOrWhiteSpace(element.LayerType) ? "Body" : element.LayerType;
        return $"{id}\u001f{type}";
    }

    private static void SetPathColor(
        XElement pathObject,
        XNamespace ns,
        string localName,
        OfdColor? color)
    {
        var colorElement = pathObject.Elements()
            .FirstOrDefault(x => x.Name.LocalName == localName);
        if (color is null)
        {
            colorElement?.Remove();
            return;
        }

        if (colorElement is null)
        {
            colorElement = new XElement(ns + localName);
            pathObject.AddFirst(colorElement);
        }

        colorElement.SetAttributeValue("Value", $"{color.Red} {color.Green} {color.Blue}");
        colorElement.SetAttributeValue("Alpha", color.Alpha == 255 ? null : color.Alpha.ToString(CultureInfo.InvariantCulture));
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

                    var fileLoc = attachment.IsExternal
                        ? attachment.ExternalPath ?? string.Empty
                        : fileName;
                    return new XElement(ns + "Attachment",
                        new XAttribute("ID", attachment.Id ?? $"Attachment_{index + 1}"),
                        new XAttribute("Name", attachment.Name),
                        new XAttribute("Format", attachment.Format ?? attachment.MediaType),
                        attachment.CreationDate.HasValue
                            ? new XAttribute("CreationDate", attachment.CreationDate.Value.ToString("O", CultureInfo.InvariantCulture))
                            : null,
                        attachment.ModificationDate.HasValue
                            ? new XAttribute("ModDate", attachment.ModificationDate.Value.ToString("O", CultureInfo.InvariantCulture))
                            : null,
                        attachment.SizeKilobytes.HasValue
                            ? new XAttribute("Size", ToInvariant(attachment.SizeKilobytes.Value))
                            : null,
                        attachment.Visible.HasValue
                            ? new XAttribute("Visible", attachment.Visible.Value)
                            : null,
                        !string.IsNullOrWhiteSpace(attachment.Usage)
                            ? new XAttribute("Usage", attachment.Usage)
                            : null,
                        new XElement(ns + "FileLoc", fileLoc));
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
                    new XAttribute("TypeID", "EMR"),
                    new XAttribute("NameSpace", "urn:ofdrw-net:custom-tags:emr"),
                    new XElement(ns + "FileLoc", "CustomTag_EMR.xml"))));

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

    private static string ResolveEntryPath(string basePath, string relativePath)
    {
        var normalizedBase = basePath.Replace('\\', '/').TrimStart('/');
        var separator = normalizedBase.LastIndexOf('/');
        var baseDirectory = separator < 0 ? string.Empty : normalizedBase.Substring(0, separator);
        var combined = relativePath.StartsWith("/", StringComparison.Ordinal)
            ? relativePath.TrimStart('/')
            : $"{baseDirectory}/{relativePath}";
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

    private static byte[] ToUtf8Bytes(XDocument xml)
    {
        using var ms = new MemoryStream();
        using (var writer = new StreamWriter(ms, new UTF8Encoding(false), 1024, leaveOpen: true))
        {
            xml.Save(writer);
        }

        return ms.ToArray();
    }

    private static long GetPreservedMaxId(IReadOnlyDictionary<string, byte[]> entries)
    {
        var maxId = 0L;
        foreach (var entry in entries)
        {
            if (!entry.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            try
            {
                using var stream = new MemoryStream(entry.Value, writable: false);
                var document = XDocument.Load(stream, LoadOptions.None);
                foreach (var attribute in document.Descendants()
                    .Attributes()
                    .Where(attribute => attribute.Name.LocalName is "ID" or "MaxUnitID"))
                {
                    if (long.TryParse(
                        attribute.Value,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out var parsed))
                    {
                        maxId = Math.Max(maxId, parsed);
                    }
                }

                foreach (var maxUnitId in document.Descendants()
                    .Where(element => element.Name.LocalName == "MaxUnitID"))
                {
                    if (long.TryParse(
                        maxUnitId.Value,
                        NumberStyles.None,
                        CultureInfo.InvariantCulture,
                        out var parsed))
                    {
                        maxId = Math.Max(maxId, parsed);
                    }
                }
            }
            catch
            {
                // Preserved XML can include vendor extensions not accepted by LINQ to XML.
            }
        }

        return maxId;
    }

    private sealed class OfdIdAllocator
    {
        private readonly HashSet<long> _usedIds = new();
        private long _nextId = 1;

        public long MaxId { get; private set; }

        public string Allocate()
        {
            while (_usedIds.Contains(_nextId))
            {
                _nextId++;
            }

            var value = _nextId++;
            _usedIds.Add(value);
            MaxId = Math.Max(MaxId, value);
            return value.ToString(CultureInfo.InvariantCulture);
        }

        public string AllocatePreferred(string? preferred)
        {
            if (long.TryParse(preferred, NumberStyles.None, CultureInfo.InvariantCulture, out var parsed) &&
                parsed > 0 &&
                _usedIds.Add(parsed))
            {
                MaxId = Math.Max(MaxId, parsed);
                return parsed.ToString(CultureInfo.InvariantCulture);
            }

            return Allocate();
        }

        public void AdvancePast(long maximumExistingId)
        {
            if (maximumExistingId <= 0)
            {
                return;
            }

            MaxId = Math.Max(MaxId, maximumExistingId);
            _nextId = Math.Max(_nextId, maximumExistingId + 1);
        }
    }
}
