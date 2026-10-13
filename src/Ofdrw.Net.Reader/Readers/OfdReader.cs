using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Interfaces;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Reader.Readers;

public sealed class OfdReader : IOfdReader
{
    private readonly OfdPackageLoader _loader = new();

    public async Task<OfdDocumentPackage> ReadAsync(Stream ofdStream, CancellationToken cancellationToken = default)
    {
        return await ReadAsync(ofdStream, new OfdPackageLoadOptions(), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Reads an OFD with explicit compressed and expanded resource budgets.</summary>
    public async Task<OfdDocumentPackage> ReadAsync(
        Stream ofdStream,
        OfdPackageLoadOptions options,
        CancellationToken cancellationToken = default)
    {
        var archive = await _loader.LoadAsync(ofdStream, options, cancellationToken).ConfigureAwait(false);
        var package = new OfdDocumentPackage();
        foreach (var entryName in archive.EntryNames)
        {
            package.PreservedEntries[entryName] = archive.GetBytes(entryName);
        }

        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(package, cancellationToken);
        material.GuardPreservedXml(package, cancellationToken);
        var ofdXml = material.ReadXml(OfdConstants.OfdRootFile, archive.GetBytes(OfdConstants.OfdRootFile), cancellationToken);
        var ofdNs = ofdXml.Root?.Name.Namespace ?? XNamespace.Get(OfdConstants.Namespace);
        var docType = ofdXml.Root?.Attribute("DocType")?.Value ?? OfdConstants.DefaultDocType;

        var docRoot = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .Elements(ofdNs + "DocRoot")
            .Select(x => x.Value)
            .FirstOrDefault() ?? "Doc_0/Document.xml";

        docRoot = material.ResolvePath("OFD.xml", docRoot, cancellationToken);
        package.Options.DocType = docType;
        package.DocumentEntryPath = docRoot;
        package.Options.Namespace = ofdNs.NamespaceName;
        package.Options.DocumentId = docRoot.Split('/')[0];

        var selectedDocBody = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .FirstOrDefault(body => string.Equals(
                material.ResolvePath("OFD.xml", body.Element(ofdNs + "DocRoot")?.Value ?? string.Empty, cancellationToken),
                docRoot,
                StringComparison.Ordinal));
        var info = selectedDocBody?.Element(ofdNs + "DocInfo");
        if (info is not null)
        {
            package.Options.Metadata.Title = info.Element(ofdNs + "Title")?.Value;
            package.Options.Metadata.Author = info.Element(ofdNs + "Author")?.Value;
            package.Options.Metadata.Subject = info.Element(ofdNs + "Subject")?.Value;
            package.Options.Metadata.Keywords = info.Element(ofdNs + "Keywords")?.Value;
            package.Options.Metadata.Creator = info.Element(ofdNs + "Creator")?.Value;
        }

        foreach (var bodyElement in selectedDocBody?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (bodyElement.Name != ofdNs + "DocInfo" && bodyElement.Name != ofdNs + "DocRoot")
            {
                package.PreservedDocBodyElements.Add(
                    OfdXmlEnvelope.Capture(bodyElement, material, cancellationToken));
            }
        }

        var documentXml = material.ReadXml(docRoot, archive.GetBytes(docRoot), cancellationToken);
        var docNs = documentXml.Root?.Name.Namespace ?? ofdNs;
        var commonData = documentXml.Root?.Element(docNs + "CommonData");
        foreach (var commonElement in commonData?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (commonElement.Name != docNs + "MaxUnitID" && commonElement.Name != docNs + "PageArea" &&
                commonElement.Name != docNs + "PublicRes" && commonElement.Name != docNs + "DocumentRes")
            {
                package.PreservedCommonDataElements.Add(
                    OfdXmlEnvelope.Capture(commonElement, material, cancellationToken));
            }
        }

        foreach (var documentElement in documentXml.Root?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (documentElement.Name != docNs + "CommonData" && documentElement.Name != docNs + "Pages" &&
                documentElement.Name != docNs + "Attachments" && documentElement.Name != docNs + "CustomTags")
            {
                package.PreservedDocumentElements.Add(
                    OfdXmlEnvelope.Capture(documentElement, material, cancellationToken));
            }
        }

        var defaultPageBox = ParseBox(commonData?
            .Element(docNs + "PageArea")?
            .Element(docNs + "PhysicalBox")?.Value);
        if (defaultPageBox.w > 0 && defaultPageBox.h > 0)
        {
            package.Options.DefaultPageWidthMillimeters = defaultPageBox.w;
            package.Options.DefaultPageHeightMillimeters = defaultPageBox.h;
        }
        var fontMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documentMediaMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documentMediaTypeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var publicResLoc = commonData?.Element(docNs + "PublicRes")?.Value;
        package.PublicResourceLocation = publicResLoc;
        if (!string.IsNullOrWhiteSpace(publicResLoc))
        {
            var publicResPath = material.ResolvePath(docRoot, publicResLoc!, cancellationToken);
            if (archive.Contains(publicResPath))
            {
                ReadFonts(archive, publicResPath, package, fontMap, material, cancellationToken);
                ReadMediaResources(archive, publicResPath, docNs, documentMediaMap, documentMediaTypeMap, material, cancellationToken);
            }
        }

        var documentResLoc = commonData?.Element(docNs + "DocumentRes")?.Value;
        package.DocumentResourceLocation = documentResLoc;
        if (!string.IsNullOrWhiteSpace(documentResLoc))
        {
            var documentResPath = material.ResolvePath(docRoot, documentResLoc!, cancellationToken);
            ReadMediaResources(archive, documentResPath, docNs, documentMediaMap, documentMediaTypeMap, material, cancellationToken);
            ReadFonts(archive, documentResPath, package, fontMap, material, cancellationToken);
        }

        var documentFonts = new Dictionary<string, string>(fontMap, StringComparer.OrdinalIgnoreCase);
        var templateLocations = commonData?.Elements(docNs + "TemplatePage")
            .Select(x => new
            {
                Id = x.Attribute("ID")?.Value,
                BaseLoc = x.Attribute("BaseLoc")?.Value,
                ZOrder = x.Attribute("ZOrder")?.Value
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.BaseLoc))
            .ToDictionary(x => x.Id!, x => (BaseLoc: x.BaseLoc!, ZOrder: x.ZOrder), StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, (string BaseLoc, string? ZOrder)>(StringComparer.OrdinalIgnoreCase);

        var pages = documentXml.Root?
            .Element(docNs + "Pages")?
            .Elements(docNs + "Page")
            .Select((x, index) => new
            {
                Index = index,
                Id = x.Attribute("ID")?.Value,
                BaseLoc = x.Attribute("BaseLoc")?.Value
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.BaseLoc))
            .Take(options.MaxPageCount == int.MaxValue ? int.MaxValue : options.MaxPageCount + 1)
            .ToList() ?? [];

        if (pages.Count > options.MaxPageCount)
            throw new InvalidDataException("OFD document exceeds the configured page count limit.");

        var annotationIndex = IndexAnnotationFiles(archive, documentXml, docRoot,
            new HashSet<string>(pages.Select(page => page.Id).OfType<string>().Where(id => !string.IsNullOrWhiteSpace(id)), StringComparer.Ordinal), material, cancellationToken);
        var annotationIdCounts = new Dictionary<long, int>();
        foreach (var entry in package.PreservedEntries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var reservedIds = material.Reservations(entry.Value, out _, cancellationToken);
            if (!reservedIds.Any() || material.OpaqueEntries.Contains(entry.Key)) continue;
            foreach (var attribute in material.ReadXml(entry.Key, entry.Value, cancellationToken).Descendants().Attributes("ID"))
            {
                material.Structure.Reference(); var id = OfdSourceIdentity.ParseId(attribute.Value);
                if (id != 0) annotationIdCounts[id] = annotationIdCounts.TryGetValue(id, out var count) ? count + 1 : 1;
            }
        }
        package.AnnotationIndexSource = annotationIndex.Source; annotationIndex.Source.DocumentOwner = package;
        var annotationDocuments = new Dictionary<string, XDocument?>(StringComparer.OrdinalIgnoreCase);
        var annotationBudget = new AnnotationReadBudget(options);
        foreach (var pageRef in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            fontMap = new Dictionary<string, string>(documentFonts, StringComparer.OrdinalIgnoreCase);
            var contentPath = material.ResolvePath(docRoot, pageRef.BaseLoc!, cancellationToken);
            var pageXml = material.ReadXml(contentPath, archive.GetBytes(contentPath), cancellationToken);
            var pageNs = pageXml.Root?.Name.Namespace ?? docNs;

            var box = ParseBox(pageXml.Root?
                .Element(pageNs + "Area")?
                .Element(pageNs + "PhysicalBox")?.Value);

            var page = new OfdPage
            {
                SourceEntryPath = contentPath,
                Id = pageRef.Id,
                Index = pageRef.Index,
                XMillimeters = box.w > 0 ? box.x : defaultPageBox.x,
                YMillimeters = box.h > 0 ? box.y : defaultPageBox.y,
                WidthMillimeters = box.w > 0 ? box.w : defaultPageBox.w,
                HeightMillimeters = box.h > 0 ? box.h : defaultPageBox.h
            };

            foreach (var pageElement in pageXml.Root?.Elements() ?? Enumerable.Empty<XElement>())
            {
                if (pageElement.Name != pageNs + "Area" && pageElement.Name != pageNs + "Content")
                {
                    page.PreservedPageElements.Add(OfdXmlEnvelope.Capture(pageElement, material, cancellationToken));
                }
            }

            var pageDir = GetDirectory(contentPath);
            var pageResPath = $"{pageDir}/PageRes.xml";
            var mediaMap = new Dictionary<string, string>(documentMediaMap, StringComparer.OrdinalIgnoreCase);
            var mediaTypeMap = new Dictionary<string, string>(documentMediaTypeMap, StringComparer.OrdinalIgnoreCase);

            if (archive.Contains(pageResPath))
            {
                ReadMediaResources(archive, pageResPath, docNs, mediaMap, mediaTypeMap, material, cancellationToken);
                ReadFonts(archive, pageResPath, package, fontMap, material, cancellationToken);
            }

            foreach (var element in ParsePageObjects(
                archive,
                pageXml,
                fontMap,
                mediaMap,
                mediaTypeMap,
                pageResPath, material, cancellationToken))
            {
                page.Elements.Add(element);
            }

            var annotationFonts = new Dictionary<string, string>(fontMap, StringComparer.OrdinalIgnoreCase);
            foreach (var templateRef in pageXml.Root?.Elements(pageNs + "Template") ?? Enumerable.Empty<XElement>())
            {
                var templateId = templateRef.Attribute("TemplateID")?.Value;
                if (string.IsNullOrWhiteSpace(templateId) ||
                    !templateLocations.TryGetValue(templateId!, out var templateLoc))
                {
                    continue;
                }

                var templatePath = material.ResolvePath(docRoot, templateLoc.BaseLoc, cancellationToken);
                if (!archive.Contains(templatePath))
                {
                    continue;
                }

                var templateXml = material.ReadXml(templatePath, archive.GetBytes(templatePath), cancellationToken);
                var templateResPath = $"{GetDirectory(templatePath)}/PageRes.xml";
                var templateMediaMap = new Dictionary<string, string>(documentMediaMap, StringComparer.OrdinalIgnoreCase);
                var templateMediaTypeMap = new Dictionary<string, string>(documentMediaTypeMap, StringComparer.OrdinalIgnoreCase);
                ReadMediaResources(archive, templateResPath, docNs, templateMediaMap, templateMediaTypeMap, material, cancellationToken);
                var templateFonts = new Dictionary<string, string>(fontMap, StringComparer.OrdinalIgnoreCase);
                ReadFonts(archive, templateResPath, package, templateFonts, material, cancellationToken);

                var template = new OfdTemplateContent
                {
                    TemplateId = templateId!,
                    ZOrder = SupportedTemplateOrder(templateRef.Attribute("ZOrder")?.Value)
                        ?? SupportedTemplateOrder(templateLoc.ZOrder) ?? "Background",
                    BaseLocation = templateLoc.BaseLoc
                };
                foreach (var element in ParsePageObjects(
                    archive,
                    templateXml,
                    templateFonts,
                    templateMediaMap,
                    templateMediaTypeMap,
                    templateResPath, material, cancellationToken))
                {
                    template.Elements.Add(element);
                }

                page.Templates.Add(template);
            }

            ReadAnnotationAppearances(archive, annotationIndex, annotationDocuments, annotationBudget, page, annotationFonts, mediaMap, mediaTypeMap, material, cancellationToken, annotationIdCounts);
            package.Pages.Add(page);
        }

        var attachmentsLoc = documentXml.Root?.Elements()
            .FirstOrDefault(x => x.Name.LocalName == "Attachments")?.Value;
        if (!string.IsNullOrWhiteSpace(attachmentsLoc))
        {
            var attachmentsPath = material.ResolvePath(docRoot, attachmentsLoc!, cancellationToken);
            if (archive.Contains(attachmentsPath))
            {
                var attachmentsXml = material.ReadXml(attachmentsPath, archive.GetBytes(attachmentsPath), cancellationToken);
                foreach (var node in attachmentsXml.Root?.Elements()
                    .Where(x => x.Name.LocalName == "Attachment") ?? Enumerable.Empty<XElement>())
                {
                    var fileLoc = node.Elements().FirstOrDefault(x => x.Name.LocalName == "FileLoc")?.Value
                        ?? node.Attribute("FileLoc")?.Value
                        ?? string.Empty;
                    var contentPath = material.ResolvePath(attachmentsPath, fileLoc, cancellationToken);
                    var external = bool.TryParse(node.Attribute("External")?.Value, out var legacyExternal) && legacyExternal;
                    external = external || !archive.Contains(contentPath);

                    var attachment = new OfdAttachment
                    {
                        Id = node.Attribute("ID")?.Value,
                        Name = node.Attribute("Name")?.Value ?? "Attachment",
                        MediaType = node.Attribute("MediaType")?.Value
                            ?? node.Attribute("Format")?.Value
                            ?? "application/octet-stream",
                        Format = node.Attribute("Format")?.Value,
                        CreationDate = ParseDateTime(node.Attribute("CreationDate")?.Value),
                        ModificationDate = ParseDateTime(node.Attribute("ModDate")?.Value),
                        SizeKilobytes = ParseNullableDouble(node.Attribute("Size")?.Value),
                        Visible = ParseNullableBoolean(node.Attribute("Visible")?.Value),
                        Usage = node.Attribute("Usage")?.Value,
                        IsExternal = external,
                        ExternalPath = external ? fileLoc : null
                    };

                    if (!external && archive.TryGetBytes(contentPath, out var bytes))
                    {
                        attachment.Data = bytes;
                    }

                    package.Attachments.Add(attachment);
                }
            }
        }

        var customTagsLoc = documentXml.Root?.Elements()
            .FirstOrDefault(x => x.Name.LocalName == "CustomTags")?.Value;
        if (!string.IsNullOrWhiteSpace(customTagsLoc))
        {
            var customTagsPath = material.ResolvePath(docRoot, customTagsLoc!, cancellationToken);
            if (archive.Contains(customTagsPath))
            {
                var customTagsXml = material.ReadXml(customTagsPath, archive.GetBytes(customTagsPath), cancellationToken);
                foreach (var customTag in customTagsXml.Root?.Elements()
                    .Where(x => x.Name.LocalName == "CustomTag") ?? Enumerable.Empty<XElement>())
                {
                    var detailLoc = customTag.Elements().FirstOrDefault(x => x.Name.LocalName == "FileLoc")?.Value
                        ?? customTag.Attribute("FileLoc")?.Value;
                    if (string.IsNullOrWhiteSpace(detailLoc))
                    {
                        continue;
                    }

                    var detailPath = material.ResolvePath(customTagsPath, detailLoc!, cancellationToken);
                    ReadKeyValueCustomTags(archive, detailPath, package.CustomTags, material, cancellationToken);
                }
            }
        }
        else
        {
            var legacyCustomTagPath = $"{package.Options.DocumentId}/Tags/CustomTag_EMR.xml";
            ReadKeyValueCustomTags(archive, legacyCustomTagPath, package.CustomTags, material, cancellationToken);
        }

        OfdNavigationXml.Read(package, documentXml.Root ?? throw new InvalidDataException("Missing document root."), material, cancellationToken);
        package.SourceEnvelope = OfdDocumentEnvelope.Capture(package, selectedDocBody, documentXml.Root!, material, cancellationToken);
        package.SourceIdentity = OfdSourceIdentity.Capture(package, material, cancellationToken);
        material.InspectStorage(package, cancellationToken);
        return package;
    }

    private sealed class AnnotationFiles
    {
        internal readonly List<string> Paths = new();
        internal readonly HashSet<string> Known = new(StringComparer.OrdinalIgnoreCase);
        internal readonly List<string> UnmodeledRecords = new();
        internal readonly List<string> Records = new();
        internal string? UnmodeledXml;
    }

    private sealed class AnnotationIndex
    {
        internal readonly OfdAnnotationIndexSource Source = new();
        internal XNamespace Namespace = XNamespace.None;
        internal readonly Dictionary<string, AnnotationFiles> Files = new(StringComparer.Ordinal);
        internal readonly List<string> UnmodeledLists = new();
        internal string? UnmodeledXml;
    }

    private static bool IsAnnotationRoot(XDocument xml, string localName, XNamespace documentNamespace) =>
        xml.Root?.Name.LocalName == localName && (xml.Root.Name.Namespace == documentNamespace ||
            xml.Root.Name.NamespaceName == OfdConstants.Namespace || xml.Root.Name.NamespaceName == OfdConstants.StandardNamespace);

    private static string? SupportedTemplateOrder(string? value) => value is "Background" or "Foreground" ? value : null;

    private static bool PlainContainer(XElement node, params string[] attributes) =>
        node.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && attributes.Contains(attribute.Name.LocalName)) &&
        node.Nodes().All(child => child is XElement or XComment || child is XText text && string.IsNullOrWhiteSpace(text.Value));

    private static bool LiteralAnnotationValue(XElement node, params string[] attributes) =>
        !node.HasElements && node.Nodes().All(child => child is XText or XComment) &&
        node.Attributes().All(attribute => attribute.IsNamespaceDeclaration ||
            attribute.Name.Namespace == XNamespace.None && attributes.Contains(attribute.Name.LocalName));

    private static bool SupportedAnnotationWrapper(XElement annotation)
    {
        var ns = annotation.Name.Namespace;
        if (!PlainContainer(annotation, "ID", "Type", "Creator", "LastModDate", "Subtype", "Visible", "Print", "NoZoom", "NoRotate", "ReadOnly") ||
            annotation.Elements(ns + "Appearance").Count() != 1 || annotation.Elements(ns + "Remark").Count() > 1 ||
            annotation.Elements(ns + "Parameters").Count() > 1) return false;
        foreach (var child in annotation.Elements())
        {
            if (child.Name == ns + "Appearance") continue;
            if (child.Name == ns + "Remark" && LiteralAnnotationValue(child)) continue;
            if (child.Name == ns + "Parameters" && PlainContainer(child) && child.Elements().All(parameter =>
                parameter.Name == ns + "Parameter" && LiteralAnnotationValue(parameter, "Name"))) continue;
            return false;
        }
        return true;
    }

    private static AnnotationIndex IndexAnnotationFiles(OfdPackageArchive archive, XDocument document,
        string documentPath, ISet<string> loadedPageIds, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        var result = new AnnotationIndex { Namespace = document.Root!.Name.Namespace };
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in document.Root!.Elements().Where(node => node.Name.LocalName == "Annotations"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            result.Source.Declarations.Add(OfdXmlEnvelope.Capture(declaration, material, cancellationToken));
            if (declaration.Name.Namespace != result.Namespace)
            {
                result.UnmodeledLists.Add(declaration.ToString(SaveOptions.DisableFormatting));
                continue;
            }
            var literal = string.Concat(declaration.Nodes().OfType<XText>().Select(text => text.Value));
            if (declaration.HasElements || declaration.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                declaration.Nodes().Any(node => node is not XText and not XComment))
                result.UnmodeledLists.Add(declaration.ToString(SaveOptions.DisableFormatting));
            if (string.IsNullOrWhiteSpace(literal))
            {
                result.UnmodeledLists.Add(declaration.ToString(SaveOptions.DisableFormatting));
                continue;
            }
            var listPath = material.ResolvePath(documentPath, literal, cancellationToken);
            if (!visited.Add(listPath)) continue;
            result.Source.Capture(listPath, archive.TryGetBytes(listPath, out var indexBytes) ? indexBytes : null, material, cancellationToken);
            if (!archive.Contains(listPath))
            {
                result.UnmodeledLists.Add(declaration.ToString(SaveOptions.DisableFormatting));
                continue;
            }
            XDocument list;
            try { list = material.ReadXml(listPath, archive.GetBytes(listPath), cancellationToken); }
            catch (XmlException)
            {
                result.UnmodeledLists.Add(declaration.ToString(SaveOptions.DisableFormatting));
                continue;
            }
            if (!IsAnnotationRoot(list, "Annotations", result.Namespace))
            {
                result.UnmodeledLists.Add(declaration.ToString(SaveOptions.DisableFormatting));
                continue;
            }
            if (!PlainContainer(list.Root!) || list.Nodes().Any(node => node is not XElement and not XComment && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                result.UnmodeledLists.Add(list.ToString(SaveOptions.DisableFormatting));
            foreach (var record in list.Root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var pageId = record.Attribute("PageID")?.Value;
                var locations = record.Elements(list.Root.Name.Namespace + "FileLoc").ToList();
                var location = locations.Count == 1 && !locations[0].HasElements &&
                    locations[0].Nodes().All(node => node is XText or XComment)
                    ? string.Concat(locations[0].Nodes().OfType<XText>().Select(text => text.Value)) : null;
                AnnotationFiles? files = null;
                var raw = record.ToString(SaveOptions.DisableFormatting);
                if (!string.IsNullOrWhiteSpace(pageId))
                {
                    if (!result.Files.TryGetValue(pageId!, out files)) result.Files[pageId!] = files = new AnnotationFiles();
                    files.Records.Add(raw);
                }
                if (record.Name != list.Root.Name.Namespace + "Page" || string.IsNullOrWhiteSpace(pageId) || string.IsNullOrWhiteSpace(location))
                {
                    if (files is null) result.UnmodeledLists.Add(raw); else files.UnmodeledRecords.Add(raw);
                    continue;
                }
                if (record.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name != XName.Get("PageID")) ||
                    locations[0].Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                    record.Elements().Count() != 1 || record.Nodes().Any(node => node is not XElement and not XComment &&
                        (node is not XText text || !string.IsNullOrWhiteSpace(text.Value)))) files!.UnmodeledRecords.Add(raw);
                var path = material.ResolvePath(listPath, location!, cancellationToken);
                if (files!.Known.Add(path)) files.Paths.Add(path);
            }
        }
        foreach (var orphan in result.Files.Where(pair => !loadedPageIds.Contains(pair.Key)).ToList())
        {
            result.UnmodeledLists.AddRange(orphan.Value.Records);
            result.Files.Remove(orphan.Key);
        }
        result.UnmodeledXml = WrapUnmodeled(result.UnmodeledLists);
        foreach (var files in result.Files.Values) files.UnmodeledXml = WrapUnmodeled(files.UnmodeledRecords);
        return result;
    }

    private static string? WrapUnmodeled(List<string> records) => records.Count == 0 ? null :
        "<UnmodeledAnnotations>" + string.Concat(records) + "</UnmodeledAnnotations>";

    private sealed class AnnotationReadBudget(OfdPackageLoadOptions options)
    {
        private int remainingObjects = options.MaxAnnotationObjectCount;
        private long remainingXmlBytes = options.MaxAnnotationXmlBytes;

        internal void ChargeXml(long bytes)
        {
            if (bytes > remainingXmlBytes)
                throw new InvalidDataException("Annotation materialization exceeds the cumulative XML byte budget.");
            remainingXmlBytes -= bytes;
        }

        internal void ChargeObject()
        {
            if (remainingObjects == 0)
                throw new InvalidDataException("Annotation materialization exceeds the cumulative object budget.");
            remainingObjects--;
        }
    }

    private static void ReadAnnotationAppearances(OfdPackageArchive archive, AnnotationIndex index,
        IDictionary<string, XDocument?> documents, AnnotationReadBudget budget, OfdPage page, IReadOnlyDictionary<string, string> fonts,
        IReadOnlyDictionary<string, string> media, IReadOnlyDictionary<string, string> mediaTypes,
        OfdModelMaterialization material, CancellationToken cancellationToken, IReadOnlyDictionary<long, int> idCounts)
    {
        var typed = new OfdAnnotationCollection();
        if (index.UnmodeledXml is not null)
            page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "Annotations", Xml = index.UnmodeledXml });
        if (page.Id is null || !index.Files.TryGetValue(page.Id, out var files))
        { page.AnnotationOwnership.Bind(page, null, material, cancellationToken); return; }
        if (files.UnmodeledXml is not null)
            page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "Annotations", Xml = files.UnmodeledXml });
        foreach (var annotationPath in files.Paths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Cached XML still produces new typed/raw strings on every page.
            // Charge each occurrence before parsing or materializing its content.
            if (archive.TryGetBytes(annotationPath, out var annotationBytes)) budget.ChargeXml(annotationBytes.LongLength);
            if (!index.Source.IndexDigests.ContainsKey(annotationPath)) index.Source.Capture(annotationPath, annotationBytes, material, cancellationToken);
            if (!documents.TryGetValue(annotationPath, out var annotations))
            {
                annotations = null;
                if (archive.Contains(annotationPath))
                    try { annotations = material.ReadXml(annotationPath, archive.GetBytes(annotationPath), cancellationToken); }
                    catch (XmlException) { }
                documents[annotationPath] = annotations;
            }
            if (annotations is null || !IsAnnotationRoot(annotations, "PageAnnot", index.Namespace))
            {
                page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance",
                    Xml = new XElement("UnmodeledAnnotation", new XAttribute("FileLoc", annotationPath)).ToString() });
                continue;
            }
            page.AnnotationOwnership.Entries.Add(annotationPath, OfdXmlEnvelope.Capture(annotations.Root!, material, cancellationToken));
            if (annotations.Root!.Attributes().Any(a => !a.IsNamespaceDeclaration) || annotations.Root.Nodes().Any(n => n is XProcessingInstruction or XComment || n is XText t && !string.IsNullOrWhiteSpace(t.Value)) || annotations.Nodes().Any(n => n is XProcessingInstruction or XComment))
                page.AnnotationOwnership.UnsafeDeletionEntries.Add(annotationPath);
            var shared = index.Files.Values.Count(candidate => candidate.Known.Contains(annotationPath)) > 1;
            var annotationSlot = -1;
            if (!PlainContainer(annotations.Root!) || annotations.Nodes().Any(node => node is not XElement and not XComment && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = annotations.ToString(SaveOptions.DisableFormatting) });
            foreach (var annotation in annotations.Root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                annotationSlot++;
                // Strict caller quota is charged before the first local typed materialization.
                foreach (var appearanceNode in annotation.Elements(annotation.Name.Namespace + "Appearance"))
                foreach (var graphic in appearanceNode.Descendants())
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (graphic.Name.Namespace == appearanceNode.Name.Namespace && graphic.Name.LocalName is "TextObject" or "ImageObject" or "PathObject") budget.ChargeObject();
                }
                var projectionStart = page.AnnotationAppearances.Count;
                var admitted = shared ? null : ReadTypedAnnotation(archive, annotation, fonts, media, mediaTypes,
                    annotationPath, idCounts, material, cancellationToken);
                try
                {
                if (annotation.Name != annotations.Root.Name.Namespace + "Annot")
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = annotation.ToString(SaveOptions.DisableFormatting) });
                    continue;
                }
                if (!SupportedAnnotationWrapper(annotation))
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = annotation.ToString(SaveOptions.DisableFormatting) });
                if (annotation.Attribute("Visible")?.Value is "false" or "0")
                {
                    var hiddenAppearance = annotation.Element(annotations.Root.Name.Namespace + "Appearance");
                    if (hiddenAppearance is not null)
                    {
                        var hiddenNodes = FlattenAnnotationNodes(hiddenAppearance.Elements(), hiddenAppearance.Name.Namespace, budget, cancellationToken);
                        if (!PlainContainer(hiddenAppearance, "Boundary", "ID", "CTM") || hiddenNodes is null ||
                            !OfdBoxParser.TryParse(hiddenAppearance.Attribute("Boundary")?.Value, out _) ||
                            hiddenNodes.SelectMany(node => node.DescendantsAndSelf().Attributes()).Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute) ||
                                attribute.Name == "Boundary" && !OfdBoxParser.TryParse(attribute.Value, out _)))
                            page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = annotation.ToString(SaveOptions.DisableFormatting) });
                    }
                    continue;
                }
                var appearances = annotation.Elements(annotations.Root.Name.Namespace + "Appearance").ToList();
                if (appearances.Count > 1)
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = annotation.ToString(SaveOptions.DisableFormatting) });
                    continue;
                }
                var appearance = appearances.Count == 1 ? appearances[0] : null;
                if (appearance is null)
                {
                    if (annotation.Elements().Any(node => node.Name.LocalName == "Appearance"))
                        page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = annotation.ToString(SaveOptions.DisableFormatting) });
                    continue;
                }
                var appearanceTransform = ParseMatrix(appearance.Attribute("CTM")?.Value);
                var hasUnmodeledMetadata = !PlainContainer(appearance, "Boundary", "ID", "CTM") ||
                    !OfdBoxParser.TryParse(appearance.Attribute("Boundary")?.Value, out _);
                if (appearance.Attribute("CTM") is not null && (appearanceTransform is null || appearanceTransform.Any(value => double.IsNaN(value) || double.IsInfinity(value))))
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = appearance.ToString(SaveOptions.DisableFormatting) });
                    continue;
                }
                var ns = appearance.Name.Namespace;
                var box = ParseBox(appearance.Attribute("Boundary")?.Value);
                var nodes = FlattenAnnotationNodes(appearance.Elements(), appearance.Name.Namespace, budget, cancellationToken);
                if (nodes is not null && nodes.SelectMany(node => node.DescendantsAndSelf().Attributes("Boundary")).Any(attribute => !OfdBoxParser.TryParse(attribute.Value, out _)))
                    hasUnmodeledMetadata = true;
                if (nodes is null || box.w <= 0 || box.h <= 0 || new[] { box.x, box.y, box.w, box.h }.Any(value => double.IsNaN(value) || double.IsInfinity(value)))
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = appearance.ToString(SaveOptions.DisableFormatting) });
                    continue;
                }
                var xml = new XDocument(new XElement(ns + "Page", new XElement(ns + "Content",
                    new XElement(ns + "Layer", new XAttribute("ID", "annotation-" + annotation.Attribute("ID")?.Value), nodes))));
                var staged = new List<OfdElement>();
                if (hasUnmodeledMetadata) staged.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = appearance.ToString(SaveOptions.DisableFormatting) });
                if (nodes.SelectMany(node => node.DescendantsAndSelf().Attributes()).Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute)))
                    staged.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = appearance.ToString(SaveOptions.DisableFormatting) });
                foreach (var element in ParsePageObjects(archive, xml, fonts, media, mediaTypes, annotationPath, material, cancellationToken))
                {
                    if (element is OfdImageElement { Data.Length: 0 })
                    {
                        staged.Clear();
                        staged.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = appearance.ToString(SaveOptions.DisableFormatting) });
                        break;
                    }
                    try
                    {
                        staged.Add(OfdAnnotationProjection.Project(element, new OfdAnnotationAppearance
                        {
                            XMillimeters = box.x, YMillimeters = box.y, WidthMillimeters = box.w, HeightMillimeters = box.h,
                            Transform = appearanceTransform
                        }, ns, material, cancellationToken));
                    }
                    catch (Exception exception) when (exception is NotSupportedException or InvalidDataException && !OfdTraversalBudget.IsLimit(exception))
                    {
                        staged.Clear();
                        staged.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = appearance.ToString(SaveOptions.DisableFormatting) });
                        break;
                    }
                }
                page.AnnotationAppearances.AddRange(staged);
                }
                finally
                {
                    if (admitted is not null)
                    {
                        admitted.Source = new OfdAnnotationSource(annotationPath, annotationSlot, projectionStart,
                            page.AnnotationAppearances.Count - projectionStart,
                            OfdModelSnapshot.Capture(admitted.Appearance.Elements, material, cancellationToken), annotations.Root.Name.NamespaceName);
                        typed.Items.Add(admitted);
                    }
                    else page.AnnotationOwnership.HasUnclaimedSource = true;
                }
            }
        }
        page.AnnotationOwnership.Bind(page, typed.Items.Count == 0 ? null : typed, material, cancellationToken);
    }

    private static OfdAnnotation? ReadTypedAnnotation(OfdPackageArchive archive, XElement annotation,
        IReadOnlyDictionary<string, string> fonts, IReadOnlyDictionary<string, string> media,
        IReadOnlyDictionary<string, string> mediaTypes, string path, IReadOnlyDictionary<long, int> idCounts,
        OfdModelMaterialization material, CancellationToken token)
    {
        var ns = annotation.Name.Namespace;
        if (annotation.Name != ns + "Annot" || (ns.NamespaceName != OfdConstants.Namespace && ns.NamespaceName != OfdConstants.StandardNamespace) ||
            !Unique(annotation.Attribute("ID")?.Value) ||
            !Enum.TryParse<OfdAnnotationType>(annotation.Attribute("Type")?.Value, false, out var type) || !Enum.IsDefined(typeof(OfdAnnotationType), type) ||
            string.IsNullOrWhiteSpace(annotation.Attribute("Creator")?.Value) ||
            !DateTime.TryParseExact(annotation.Attribute("LastModDate")?.Value, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date) ||
            annotation.Elements(ns + "Appearance").Count() != 1 || annotation.Elements(ns + "Remark").Count() > 1 ||
            annotation.Nodes().Any(n => n is XText t && !string.IsNullOrWhiteSpace(t.Value))) return null;
        var flags = new Dictionary<string, bool?>();
        foreach (var name in new[] { "Visible", "Print", "NoZoom", "NoRotate", "ReadOnly" })
        {
            var literal = annotation.Attribute(name)?.Value;
            if (literal is not null && literal is not ("true" or "false" or "1" or "0")) return null;
            flags[name] = literal is null ? null : literal is "true" or "1";
        }
        var appearance = annotation.Element(ns + "Appearance")!;
        if (!OfdBoxParser.TryParse(appearance.Attribute("Boundary")?.Value, out _) ||
            appearance.Nodes().Any(n => n is XText t && !string.IsNullOrWhiteSpace(t.Value)) ||
            appearance.Elements().Any(n => n.Name.Namespace != ns || n.Name.LocalName is not ("TextObject" or "ImageObject" or "PathObject") ||
                !Unique(n.Attribute("ID")?.Value) || !IsSupportedAnnotationPrimitive(n, ns))) return null;
        foreach (var primitive in appearance.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (!OfdBoxParser.TryParse(primitive.Attribute("Boundary")?.Value, out var primitiveBox) ||
                !OfdBoxParser.IsWritablePositiveSide(primitiveBox.w) || !OfdBoxParser.IsWritablePositiveSide(primitiveBox.h)) return null;
            var primitiveMatrix = ParseMatrix(primitive.Attribute("CTM")?.Value);
            if (primitive.Attribute("CTM") is not null && (primitiveMatrix is null || primitiveMatrix.Any(v => double.IsNaN(v) || double.IsInfinity(v)))) return null;
            if (primitive.Name.LocalName == "TextObject" && (!double.TryParse(primitive.Attribute("Size")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) || !(size > 0) || double.IsInfinity(size))) return null;
        }
        var matrix = ParseMatrix(appearance.Attribute("CTM")?.Value);
        if (appearance.Attribute("CTM") is not null && (matrix is null || matrix.Any(v => double.IsNaN(v) || double.IsInfinity(v)))) return null;
        var box = ParseBox(appearance.Attribute("Boundary")?.Value);
        if (box.w <= 0 || box.h <= 0) return null;
        var remark = annotation.Element(ns + "Remark");
        if (remark is not null && (remark.HasElements || remark.Nodes().Any(n => n is not XText and not XComment))) return null;
        var result = new OfdAnnotation
        {
            Id = annotation.Attribute("ID")!.Value, Type = type, Creator = annotation.Attribute("Creator")!.Value,
            LastModifiedDate = DateTime.SpecifyKind(date, DateTimeKind.Unspecified), Subtype = annotation.Attribute("Subtype")?.Value,
            Visible = flags["Visible"], Print = flags["Print"], NoZoom = flags["NoZoom"], NoRotate = flags["NoRotate"], ReadOnly = flags["ReadOnly"],
            Remark = remark?.Value, SourceXml = OfdXmlEnvelope.Capture(annotation, material, token),
            Appearance = new OfdAnnotationAppearance { XMillimeters = box.x, YMillimeters = box.y, WidthMillimeters = box.w,
                HeightMillimeters = box.h, Transform = matrix, SourceXml = OfdXmlEnvelope.Capture(appearance, material, token) }
        };
        // Capture inherits namespace bindings before attaching nodes to the parser's narrow synthetic container.
        var nodes = appearance.Elements().Select(n => OfdXmlEnvelope.Parse(OfdXmlEnvelope.Capture(n, material, token))).ToArray();
        var xml = new XDocument(new XElement(ns + "Page", new XElement(ns + "Content", new XElement(ns + "Layer", nodes))));
        foreach (var element in ParsePageObjects(archive, xml, fonts, media, mediaTypes, path, material, token))
        {
            if (element is OfdRawElement || element is OfdImageElement { Data.Length: 0 }) return null;
            result.Appearance.Elements.Add(element);
        }
        return result;
        bool Unique(string? value) => OfdSourceIdentity.ParseId(value) is var id && id != 0 && idCounts.TryGetValue(id, out var count) && count == 1;
    }

    private static List<XElement>? FlattenAnnotationNodes(IEnumerable<XElement> nodes, XNamespace ns, AnnotationReadBudget budget, CancellationToken token)
    {
        var result = new List<XElement>();
        bool Append(IEnumerable<XElement> children, int depth)
        {
            if (depth > 32) return false;
            foreach (var node in children)
            {
                token.ThrowIfCancellationRequested();
                if (node.Name.Namespace != ns) return false;
                if (node.Name == ns + "PageBlock")
                {
                    if (!PlainContainer(node, "ID") ||
                        !Append(node.Elements(), depth + 1)) return false;
                }
                else if (node.Name.LocalName is "TextObject" or "ImageObject" or "PathObject")
                {
                    if (!IsSupportedAnnotationPrimitive(node, ns)) return false;
                    result.Add(node);
                }
                else return false;
            }
            return true;
        }
        return Append(nodes, 0) ? result : null;
    }

    private static bool IsSupportedAnnotationPrimitive(XElement root, XNamespace ns)
    {
        if (root.Name == ns + "TextObject" && !root.Elements(ns + "TextCode").Any()) return false;
        return OfdGraphicXmlContract.HasKnownChildren(root) && !OfdGraphicXmlContract.HasUnsupportedReferences(root);
    }

    private static void ReadFonts(
        OfdPackageArchive archive, string resourcePath, OfdDocumentPackage package,
        IDictionary<string, string> fontMap, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (!archive.Contains(resourcePath)) return;
        var xml = material.ReadXml(resourcePath, archive.GetBytes(resourcePath), cancellationToken);
        var baseLocation = xml.Root?.Attribute("BaseLoc")?.Value;
        foreach (var font in xml.Descendants().Where(node => node.Name.LocalName == "Font"))
        {
            var id = font.Attribute("ID")?.Value;
            var name = font.Attribute("FontName")?.Value ?? font.Attribute("FamilyName")?.Value;
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)) continue;
            fontMap[id!] = name!;
            if (package.Fonts.Any(existing => existing.Id == id)) continue;
            var file = font.Elements().FirstOrDefault(node => node.Name.LocalName == "FontFile")?.Value;
            var resource = new OfdFontResource
            {
                Id = id!, FontName = name!, FamilyName = font.Attribute("FamilyName")?.Value,
                Charset = font.Attribute("Charset")?.Value,
                Bold = ParseBoolean(font.Attribute("Bold")?.Value, false),
                Italic = ParseBoolean(font.Attribute("Italic")?.Value, false), FileName = file
            };
            if (!string.IsNullOrWhiteSpace(file))
            {
                var reference = string.IsNullOrWhiteSpace(baseLocation) || file!.StartsWith("/", StringComparison.Ordinal)
                    ? file! : baseLocation!.TrimEnd('/') + "/" + file;
                if (archive.TryGetBytes(material.ResolvePath(resourcePath, reference, cancellationToken), out var bytes)) resource.Data = bytes;
            }
            package.Fonts.Add(resource);
        }
    }

    private static void ReadMediaResources(
        OfdPackageArchive archive,
        string resPath,
        XNamespace fallbackNs,
        IDictionary<string, string> mediaMap,
        IDictionary<string, string> mediaTypeMap, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (!archive.Contains(resPath))
        {
            return;
        }

        var resXml = material.ReadXml(resPath, archive.GetBytes(resPath), cancellationToken);
        var resNs = resXml.Root?.Name.Namespace ?? fallbackNs;
        var baseLoc = resXml.Root?.Attribute("BaseLoc")?.Value;

        foreach (var media in resXml.Descendants(resNs + "MultiMedia"))
        {
            var id = media.Attribute("ID")?.Value;
            var file = media.Attribute("MediaFile")?.Value ?? media.Element(resNs + "MediaFile")?.Value;
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(file))
            {
                continue;
            }

            var mediaLoc = string.IsNullOrWhiteSpace(baseLoc) || file!.StartsWith("/", StringComparison.Ordinal)
                ? file!
                : $"{baseLoc!.TrimEnd('/')}/{file.TrimStart('/')}";

            mediaMap[id!] = material.ResolvePath(resPath, mediaLoc, cancellationToken);
            mediaTypeMap[id!] = ToMediaType(media.Attribute("Format")?.Value);
        }
    }

    private static IEnumerable<OfdElement> ParsePageObjects(
        OfdPackageArchive archive,
        XDocument pageXml,
        IReadOnlyDictionary<string, string> fontMap,
        IReadOnlyDictionary<string, string> mediaMap,
        IReadOnlyDictionary<string, string> mediaTypeMap,
        string pageResPath, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        var layers = pageXml.Root?
            .Elements()
            .FirstOrDefault(x => x.Name.LocalName == "Content")?
            .Elements()
            .Where(x => x.Name.LocalName == "Layer")
            .ToList() ?? [];

        foreach (var layer in layers)
        {
            var layerId = layer.Attribute("ID")?.Value;
            var layerType = layer.Attribute("Type")?.Value ?? "Body";
            foreach (var node in layer.Elements())
            {
                var localName = node.Name.LocalName;
                if (string.Equals(localName, "TextObject", StringComparison.OrdinalIgnoreCase))
                {
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value);
                    var textCodes = node.Elements(node.Name.Namespace + "TextCode").ToList();
                    var text = new OfdTextElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        Text = string.Concat(textCodes.Select(ReadDirectText)),
                        FontResourceId = node.Attribute("Font")?.Value,
                        FontName = ResolveFontName(node.Attribute("Font")?.Value, fontMap),
                        FontSizeMillimeters = ParseDouble(node.Attribute("Size")?.Value, 4d),
                        Weight = (int)ParseDouble(node.Attribute("Weight")?.Value, OfdTextElement.DefaultWeight),
                        Italic = ParseBoolean(node.Attribute("Italic")?.Value, false),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value),
                        FillColor = ParseColor(node.Elements()
                            .FirstOrDefault(x => x.Name.LocalName == "FillColor")) ?? OfdColor.Black,
                        SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken),
                        ClippingXml = node.Element(node.Name.Namespace + "Clips")?.ToString(SaveOptions.DisableFormatting)
                    };
                    foreach (var textCode in textCodes)
                    {
                        text.Runs.Add(new OfdTextRun
                        {
                            Text = ReadDirectText(textCode),
                            XMillimeters = ParseDouble(textCode.Attribute("X")?.Value, 0d),
                            YMillimeters = ParseDouble(textCode.Attribute("Y")?.Value, text.FontSizeMillimeters),
                            DeltaX = textCode.Attribute("DeltaX")?.Value,
                            DeltaY = textCode.Attribute("DeltaY")?.Value
                        });
                    }

                    yield return text;
                    continue;
                }

                if (string.Equals(localName, "ImageObject", StringComparison.OrdinalIgnoreCase))
                {
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value);
                    var resourceId = node.Attribute("ResourceID")?.Value ?? string.Empty;
                    var image = new OfdImageElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        ResourceId = resourceId,
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value),
                        Alpha = (int)Math.Max(0, Math.Min(255, ParseDouble(node.Attribute("Alpha")?.Value, 255))),
                        ClipsXml = node.Element(node.Name.Namespace + "Clips")?.ToString(SaveOptions.DisableFormatting),
                        MediaType = mediaTypeMap.TryGetValue(resourceId, out var mediaType) ? mediaType : "image/png",
                        SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken)
                    };

                    if (mediaMap.TryGetValue(resourceId, out var mediaFile))
                    {
                        var mediaPath = mediaFile.Contains('/') ? mediaFile : material.ResolvePath(pageResPath, mediaFile, cancellationToken);
                        if (archive.TryGetBytes(mediaPath, out var bytes))
                        {
                            image.Data = bytes;
                            image.FileName = Path.GetFileName(mediaPath);
                        }
                    }

                    yield return image;
                    continue;
                }

                if (string.Equals(localName, "PathObject", StringComparison.OrdinalIgnoreCase))
                {
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value);
                    var strokeColor = ParseColor(node.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "StrokeColor"));
                    var fillColor = ParseColor(node.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "FillColor"));
                    yield return new OfdPathElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        AbbreviatedData = ReadDirectText(node.Element(node.Name.Namespace + "AbbreviatedData")),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value),
                        LineWidthMillimeters = ParseDouble(node.Attribute("LineWidth")?.Value, 0.353d),
                        Stroke = ParseBoolean(node.Attribute("Stroke")?.Value, true),
                        Fill = ParseBoolean(node.Attribute("Fill")?.Value, false),
                        StrokeColor = strokeColor ?? OfdColor.Black,
                        FillColor = fillColor,
                        SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken),
                        ClippingXml = node.Element(node.Name.Namespace + "Clips")?.ToString(SaveOptions.DisableFormatting)
                    };
                    continue;
                }

                var rawBoundary = ParseBox(node.Attribute("Boundary")?.Value);
                yield return new OfdRawElement
                {
                    ObjectId = node.Attribute("ID")?.Value,
                    LayerId = layerId,
                    LayerType = layerType,
                    LocalName = localName,
                    Xml = OfdXmlEnvelope.Capture(node, material, cancellationToken),
                    XMillimeters = rawBoundary.x,
                    YMillimeters = rawBoundary.y,
                    WidthMillimeters = rawBoundary.w,
                    HeightMillimeters = rawBoundary.h
                };
            }
        }
    }

    private static string ReadDirectText(XElement? node) => node is null ? string.Empty : string.Concat(node.Nodes().OfType<XText>().Where(text => text is XCData || !string.IsNullOrWhiteSpace(text.Value) || node.AncestorsAndSelf().Select(element => element.Attribute(XNamespace.Xml + "space")?.Value).FirstOrDefault(value => value is not null) == "preserve").Select(text => text.Value));

    private static void ReadKeyValueCustomTags(
        OfdPackageArchive archive,
        string detailPath,
        IDictionary<string, string> destination, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (!archive.Contains(detailPath))
        {
            return;
        }

        var tagsXml = material.ReadXml(detailPath, archive.GetBytes(detailPath), cancellationToken);
        foreach (var tag in tagsXml.Root?.Elements().Where(x => x.Name.LocalName == "Tag")
            ?? Enumerable.Empty<XElement>())
        {
            var key = tag.Attribute("Key")?.Value;
            var value = tag.Attribute("Value")?.Value;
            if (!string.IsNullOrWhiteSpace(key) && value is not null)
            {
                destination[key!] = value;
            }
        }
    }

    private static string ResolveFontName(string? fontRef, IReadOnlyDictionary<string, string> fontMap)
    {
        if (!string.IsNullOrWhiteSpace(fontRef) && fontMap.TryGetValue(fontRef!, out var fontName))
        {
            return fontName;
        }

        return string.IsNullOrWhiteSpace(fontRef) ? "SimSun" : fontRef!;
    }

    private static string ToMediaType(string? format)
    {
        return format?.Trim().ToUpperInvariant() switch
        {
            "JPG" => "image/jpeg",
            "JPEG" => "image/jpeg",
            "BMP" => "image/bmp",
            "TIFF" => "image/tiff",
            _ => "image/png"
        };
    }


    private static string GetDirectory(string path)
    {
        var normalized = path.Replace('\\', '/').TrimStart('/');
        var index = normalized.LastIndexOf('/');
        return index < 0 ? string.Empty : normalized.Substring(0, index);
    }

    private static (double x, double y, double w, double h) ParseBox(string? value)
    {
        if (value is null) return default;
        // Ordinary display keeps the legacy first four finite values, while
        // rewrite preflight requires exactly four and preserves the raw XML.
        if (OfdBoxParser.TryParse(value, out var box, allowTrailingTokens: true)) return box;
        throw new InvalidDataException("OFD box must start with four finite numbers.");
    }

    private static double ParseDouble(string? value, double fallback)
    {
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : fallback;
    }

    private static double? ParseNullableDouble(string? value)
    {
        return double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var parsed)
            ? parsed
            : null;
    }

    private static bool? ParseNullableBoolean(string? value)
    {
        return bool.TryParse(value, out var parsed) ? parsed : null;
    }

    private static bool ParseBoolean(string? value, bool fallback)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return fallback;
        }

        return value == "1" ||
            (value != "0" && bool.TryParse(value, out var parsed) ? parsed : fallback);
    }

    private static double[]? ParseMatrix(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var parts = value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 6)
        {
            return null;
        }

        var matrix = new double[6];
        for (var i = 0; i < matrix.Length; i++)
        {
            if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out matrix[i]))
            {
                return null;
            }
        }

        return matrix;
    }

    private static OfdColor? ParseColor(XElement? colorElement)
    {
        var value = colorElement?.Attribute("Value")?.Value;
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        var parts = value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        var values = parts
            .Select(x => int.TryParse(x, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed)
                ? parsed
                : -1)
            .ToArray();
        if (values.Any(x => x < 0))
        {
            return null;
        }

        var alpha = (int)ParseDouble(colorElement?.Attribute("Alpha")?.Value, 255d);
        return values.Length switch
        {
            1 => new OfdColor(values[0], values[0], values[0], alpha),
            >= 3 => new OfdColor(values[0], values[1], values[2], alpha),
            _ => null
        };
    }

    private static DateTimeOffset? ParseDateTime(string? value)
    {
        return DateTimeOffset.TryParse(
            value,
            CultureInfo.InvariantCulture,
            DateTimeStyles.AllowWhiteSpaces | DateTimeStyles.AssumeLocal,
            out var parsed)
            ? parsed
            : null;
    }
}
