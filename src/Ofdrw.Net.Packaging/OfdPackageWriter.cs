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
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Packaging;

public sealed class OfdPackageWriter
{
    public async Task WriteAsync(OfdDocumentPackage package, Stream destination, CancellationToken cancellationToken = default)
    {
        _ = await WriteWithResultAsync(package, destination, cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Writes the package and reports removed page resources and invalidated signatures.</summary>
    public async Task<OfdPackageWriteResult> WriteWithResultAsync(
        OfdDocumentPackage package, Stream destination, CancellationToken cancellationToken = default)
    {
        if (package is null)
        {
            throw new ArgumentNullException(nameof(package));
        }

        if (destination is null)
        {
            throw new ArgumentNullException(nameof(destination));
        }

        cancellationToken.ThrowIfCancellationRequested();
        var entries = BuildEntries(package);
        var result = OfdPackagePruner.Prune(package, entries);
        using var zip = new ZipArchive(destination, ZipArchiveMode.Create, leaveOpen: true);

        foreach (var entry in entries.OrderBy(x => x.Key, StringComparer.OrdinalIgnoreCase))
        {
            var zipEntry = zip.CreateEntry(entry.Key, package.Options.EnableDeflateCompression ? CompressionLevel.Optimal : CompressionLevel.NoCompression);
            using var stream = zipEntry.Open();
            await stream.WriteAsync(entry.Value, 0, entry.Value.Length, cancellationToken).ConfigureAwait(false);
        }
        return result;
    }

    private Dictionary<string, byte[]> BuildEntries(OfdDocumentPackage package)
    {
        var entries = new Dictionary<string, byte[]>(package.PreservedEntries, StringComparer.OrdinalIgnoreCase);
        var ns = XNamespace.Get(package.Options.Namespace);
        var documentPath = package.DocumentEntryPath ?? $"{package.Options.DocumentId}/Document.xml";
        var docId = OfdPackagePath.GetDirectory(documentPath);

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
                    new XElement(ns + "DocRoot", documentPath),
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
        var pagePaths = BuildPagePaths(orderedPages, docId);
        var elementIds = new Dictionary<OfdElement, string>();
        var layerIds = BuildPageObjectIds(orderedPages, elementIds, idAllocator);
        var imageResources = BuildImageResources(orderedPages, idAllocator);
        var fonts = BuildFonts(package, idAllocator);
        var publicResourceLocation = package.PublicResourceLocation ?? "PublicRes.xml";
        var documentResourceLocation = package.DocumentResourceLocation ??
            (imageResources.Count > 0 ? "DocumentRes.xml" : null);
        var resources = new OfdResourceCatalog(entries);
        var publicPath = OfdPackagePath.Resolve(documentPath, publicResourceLocation);
        resources.EnsureDocument(publicPath, ns);
        foreach (var font in fonts.Resources) resources.WriteFont(font.Id, font.Resource, publicPath, ns);
        if (!string.IsNullOrWhiteSpace(documentResourceLocation))
        {
            var resourcePath = OfdPackagePath.Resolve(documentPath, documentResourceLocation!);
            resources.EnsureDocument(resourcePath, ns);
            foreach (var image in imageResources.Values.GroupBy(image => image.Id).Select(group => group.First()))
                resources.WriteImage(image.Id, image.Image, image.Format, resourcePath, ns);
        }
        resources.Flush();
        BuildPages(orderedPages, entries, ns, pagePaths, pageIds, layerIds, elementIds, fonts.TextIds, imageResources, idAllocator);
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
                            new XAttribute("BaseLoc", OfdPackagePath.RelativeTo(documentPath, pagePaths[page])))
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

        entries[documentPath] = ToUtf8Bytes(docXml);
        return entries;
    }

    private static Dictionary<OfdPage, string> BuildPagePaths(IReadOnlyList<OfdPage> pages, string documentDirectory)
    {
        var result = new Dictionary<OfdPage, string>();
        var used = new HashSet<string>(pages.Where(page => !string.IsNullOrWhiteSpace(page.SourceEntryPath))
            .Select(page => OfdPackagePath.Resolve("OFD.xml", "/" + page.SourceEntryPath)), StringComparer.OrdinalIgnoreCase);
        foreach (var page in pages)
        {
            var path = page.SourceEntryPath;
            if (string.IsNullOrWhiteSpace(path))
            {
                var index = page.Index;
                do
                {
                    path = OfdPackagePath.Resolve("OFD.xml", $"/{documentDirectory}/Pages/Page_{index++}/Content.xml");
                } while (!used.Add(path));
            }
            else
            {
                path = OfdPackagePath.Resolve("OFD.xml", "/" + path);
            }

            if (result.Values.Contains(path, StringComparer.OrdinalIgnoreCase))
                throw new InvalidDataException("Multiple pages share the same source entry; clone repeated pages with a new source path.");
            result.Add(page, path!);
        }

        return result;
    }

    private sealed class FontBinding
    {
        internal OfdFontResource Resource { get; set; } = new();
        internal string Id { get; set; } = string.Empty;
    }

    private sealed class FontBindings
    {
        internal List<FontBinding> Resources { get; } = new();
        internal Dictionary<OfdTextElement, string> TextIds { get; } = new();
    }

    private static FontBindings BuildFonts(OfdDocumentPackage package, OfdIdAllocator idAllocator)
    {
        var result = new FontBindings();
        var byId = new Dictionary<string, FontBinding>(StringComparer.Ordinal);
        foreach (var font in package.Fonts)
        {
            if (!string.IsNullOrEmpty(font.Id) && byId.ContainsKey(font.Id))
                throw new InvalidDataException($"Duplicate font resource ID '{font.Id}'.");
            var binding = new FontBinding { Resource = font, Id = idAllocator.AllocatePreferred(font.Id) };
            result.Resources.Add(binding);
            if (!string.IsNullOrEmpty(font.Id)) byId.Add(font.Id, binding);
        }

        foreach (var text in package.Pages.SelectMany(page => page.Elements).OfType<OfdTextElement>())
        {
            FontBinding? binding = null;
            if (!string.IsNullOrEmpty(text.FontResourceId)) byId.TryGetValue(text.FontResourceId!, out binding);
            binding ??= result.Resources.FirstOrDefault(font =>
                string.Equals(font.Resource.FontName, text.FontName, StringComparison.OrdinalIgnoreCase) &&
                !font.Resource.Bold && !font.Resource.Italic)
                ?? result.Resources.FirstOrDefault(font =>
                    string.Equals(font.Resource.FontName, text.FontName, StringComparison.OrdinalIgnoreCase));
            if (binding is null)
            {
                binding = new FontBinding
                {
                    Resource = new OfdFontResource { FontName = text.FontName },
                    Id = idAllocator.AllocatePreferred(text.FontResourceId)
                };
                result.Resources.Add(binding);
                if (!string.IsNullOrEmpty(text.FontResourceId)) byId[text.FontResourceId!] = binding;
            }
            result.TextIds[text] = binding.Id;
        }
        return result;
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

        var byContent = new Dictionary<string, ImageResource>(StringComparer.Ordinal);
        var hashes = new Dictionary<byte[], string>();
        foreach (var image in pages.SelectMany(x => x.Elements).OfType<OfdImageElement>())
        {
            var format = ToOfdImageFormat(image.MediaType, image.FileName);
            if (!hashes.TryGetValue(image.Data, out var hash)) hashes[image.Data] = hash = BinaryIdentity.Hash(image.Data);
            var key = format + ":" + hash;
            if (!byContent.TryGetValue(key, out var resource))
            {
                resource = new ImageResource
                {
                    Image = image,
                    Id = idAllocator.AllocatePreferred(image.ResourceId),
                    Format = format
                };
                byContent.Add(key, resource);
            }
            resources[image] = resource;
        }

        return resources;
    }

    private static void BuildPages(
        List<OfdPage> pages,
        IDictionary<string, byte[]> entries,
        XNamespace ns,
        IReadOnlyDictionary<OfdPage, string> pagePaths,
        IReadOnlyDictionary<OfdPage, string> pageIds,
        IReadOnlyDictionary<OfdPage, Dictionary<string, string>> layerIds,
        IReadOnlyDictionary<OfdElement, string> elementIds,
        IReadOnlyDictionary<OfdTextElement, string> fontIds,
        IReadOnlyDictionary<OfdImageElement, ImageResource> imageResources,
        OfdIdAllocator idAllocator)
    {
        for (var i = 0; i < pages.Count; i++)
        {
            var page = pages[i];
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
                        var fontId = fontIds.TryGetValue(text, out var resolvedFontId)
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

                        AssignNestedIds(textObject, idAllocator);
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

                        imageObject.SetAttributeValue("Boundary", BuildBox(image.XMillimeters, image.YMillimeters, image.WidthMillimeters, image.HeightMillimeters));
                        imageObject.SetAttributeValue("CTM", image.Transform is { Length: 6 }
                            ? string.Join(" ", image.Transform.Select(ToInvariant))
                            : BuildMatrix(image.WidthMillimeters, 0, 0, image.HeightMillimeters, 0, 0));
                        imageObject.SetAttributeValue("Alpha", image.Alpha == 255 ? null : (object)Math.Max(0, Math.Min(255, image.Alpha)));
                        imageObject.Elements().Where(child => child.Name.LocalName == "Clips").Remove();
                        if (!string.IsNullOrWhiteSpace(image.ClipsXml))
                            imageObject.Add(XElement.Parse(image.ClipsXml!, LoadOptions.PreserveWhitespace));

                        AssignNestedIds(imageObject, idAllocator);
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

                        AssignNestedIds(pathObject, idAllocator);
                        layer.Add(pathObject);
                    }
                    else if (element is OfdRawElement raw && !string.IsNullOrWhiteSpace(raw.Xml))
                    {
                        var rawElement = XElement.Parse(raw.Xml, LoadOptions.PreserveWhitespace);
                        rawElement.SetAttributeValue("ID", objectId);
                        AssignNestedIds(rawElement, idAllocator);
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
            entries[pagePaths[page]] = ToUtf8Bytes(pageDocument);
        }
    }

    private static string GetLayerKey(OfdElement element)
    {
        var id = element.LayerId ?? string.Empty;
        var type = string.IsNullOrWhiteSpace(element.LayerType) ? "Body" : element.LayerType;
        return $"{id}\u001f{type}";
    }

    private static void AssignNestedIds(XElement root, OfdIdAllocator allocator)
    {
        foreach (var attribute in root.Descendants().Attributes("ID"))
            attribute.Value = allocator.AllocatePreferred(attribute.Value);
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
