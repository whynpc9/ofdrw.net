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
        return ReadArchive(archive, options, new OfdModelMaterialization(new OfdModelLimits()), cancellationToken);
    }

    internal async Task<OfdDocumentPackage> ReadControlledAsync(Stream ownedSpool, OfdPackageLoadOptions options,
        OfdControlledEditContext context, CancellationToken cancellationToken = default)
    {
        if (context is null) throw new ArgumentNullException(nameof(context));
        try
        {
            context.AssertActive(); cancellationToken.ThrowIfCancellationRequested();
            context.Material.ChargeCopy(128, cancellationToken);
            var snapshot = OfdPackageLoader.SnapshotOptions(options);
            using var scope = context.Material.BeginControlledRead(context, cancellationToken);
            var archive = await _loader.LoadAccountedAsync(ownedSpool, snapshot, context.Material, cancellationToken).ConfigureAwait(false);
            context.AssertActive();
            var package = ReadArchive(archive, snapshot, context.Material, cancellationToken);
            context.AssertActive(); cancellationToken.ThrowIfCancellationRequested(); return package;
        }
        catch { context.Fail(); throw; }
    }

    // Dedicated actual Reader bridge. Generic controlled/public routes retain
    // their guards and loading behavior; only this route records private census.
    internal async Task<OfdDocumentPackage> ReadOwnedInitialAsync(Stream actualSnapshot,
        OfdPackageLoadOptions options, OfdControlledEditContext context,
        OfdOwnedReadAcquisition acquisition, CancellationToken token)
    {
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            context.Material.ChargeCopy(128, token);
            var snapshot = OfdPackageLoader.SnapshotOptions(options);
            using var scope = context.Material.BeginControlledRead(context, token);
            var archive = await _loader.LoadAccountedAsync(actualSnapshot, snapshot, context.Material,
                token, ownedAcquisition: acquisition).ConfigureAwait(false);
            context.AssertActive();
            var package = ReadArchive(archive, snapshot, context.Material, token);
            context.AssertActive(); token.ThrowIfCancellationRequested(); return package;
        }
        catch { context.Fail(); throw; }
    }

    internal async Task<OfdDocumentPackage> ReadOwnedOutputAsync(Stream actualOutput,
        OfdControlledEditContext context, OfdOwnedZipCensus census, CancellationToken token)
    {
        try
        {
            context.AssertActive(); context.Material.ChargeCopy(512, token);
            var options = OfdPackageLoader.SnapshotOptions(new OfdPackageLoadOptions());
            using var scope = context.Material.BeginControlledRead(context, token);
            var archive = await _loader.LoadAccountedAsync(actualOutput, options, context.Material,
                token, ownedCensus: census).ConfigureAwait(false);
            context.AssertActive();
            var actualR = ReadArchive(archive, options, context.Material, token);
            context.AssertActive(); token.ThrowIfCancellationRequested(); return actualR;
        }
        catch { context.Fail(); throw; }
    }

    private static OfdDocumentPackage ReadArchive(OfdPackageArchive archive, OfdPackageLoadOptions options,
        OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        material.ReserveReadModel(1024L + archive.EntryNames.Count * 160L, cancellationToken);
        var package = new OfdDocumentPackage();
        foreach (var entryName in archive.EntryNames)
        {
            material.ReserveReadModel(160, cancellationToken);
            package.PreservedEntries[entryName] = archive.GetBytes(entryName);
        }
        material.InspectStorage(package, cancellationToken);
        material.GuardPreservedXml(package, cancellationToken);
        var ofdXml = material.ReadXml(OfdConstants.OfdRootFile, archive.GetBytes(OfdConstants.OfdRootFile), cancellationToken);
        var ofdNs = ofdXml.Root?.Name.Namespace ?? XNamespace.Get(OfdConstants.Namespace);
        var docType = ofdXml.Root?.Attribute("DocType")?.Value ?? OfdConstants.DefaultDocType;

        var docRoot = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .Elements(ofdNs + "DocRoot")
            .Select(x => ReadValue(x, material, cancellationToken)!)
            .FirstOrDefault() ?? "Doc_0/Document.xml";

        docRoot = material.ResolvePath("OFD.xml", docRoot, cancellationToken);
        package.Options.DocType = docType;
        package.DocumentEntryPath = docRoot;
        package.Options.Namespace = ofdNs.NamespaceName;
        material.ReserveReadModel(256L + docRoot.Length * 16L, cancellationToken);
        var docIdStorage = material.IsControlledRead ? material.ReserveReadStorage(2L * docRoot.Length, cancellationToken) : null;
        package.Options.DocumentId = docRoot.Split('/')[0];
        if (docIdStorage is not null) material.AdoptReadString(package.Options.DocumentId, docIdStorage, cancellationToken);

        var selectedDocBody = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .FirstOrDefault(body => string.Equals(
                material.ResolvePath("OFD.xml", ReadValue(body.Element(ofdNs + "DocRoot"), material, cancellationToken) ?? string.Empty, cancellationToken),
                docRoot,
                StringComparison.Ordinal));
        var info = selectedDocBody?.Element(ofdNs + "DocInfo");
        if (info is not null)
        {
            package.Options.Metadata.Title = ReadValue(info.Element(ofdNs + "Title"), material, cancellationToken);
            package.Options.Metadata.Author = ReadValue(info.Element(ofdNs + "Author"), material, cancellationToken);
            package.Options.Metadata.Subject = ReadValue(info.Element(ofdNs + "Subject"), material, cancellationToken);
            package.Options.Metadata.Keywords = ReadValue(info.Element(ofdNs + "Keywords"), material, cancellationToken);
            package.Options.Metadata.Creator = ReadValue(info.Element(ofdNs + "Creator"), material, cancellationToken);
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
            .Element(docNs + "PhysicalBox") is XElement physical ? ReadValue(physical, material, cancellationToken) : null, material, cancellationToken);
        if (defaultPageBox.w > 0 && defaultPageBox.h > 0)
        {
            package.Options.DefaultPageWidthMillimeters = defaultPageBox.w;
            package.Options.DefaultPageHeightMillimeters = defaultPageBox.h;
        }
        material.ReserveReadModel(1024, cancellationToken);
        var fontMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documentMediaMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documentMediaTypeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var publicResLoc = ReadValue(commonData?.Element(docNs + "PublicRes"), material, cancellationToken);
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

        var documentResLoc = ReadValue(commonData?.Element(docNs + "DocumentRes"), material, cancellationToken);
        package.DocumentResourceLocation = documentResLoc;
        if (!string.IsNullOrWhiteSpace(documentResLoc))
        {
            var documentResPath = material.ResolvePath(docRoot, documentResLoc!, cancellationToken);
            ReadMediaResources(archive, documentResPath, docNs, documentMediaMap, documentMediaTypeMap, material, cancellationToken);
            ReadFonts(archive, documentResPath, package, fontMap, material, cancellationToken);
        }

        material.ReserveReadModel(256L + (commonData?.Elements().Count() ?? 0) * 160L, cancellationToken);
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

        if (material.IsControlledRead && (documentXml.Root?.Element(docNs + "Pages")?.Elements(docNs + "Page").Count(node => !string.IsNullOrWhiteSpace(node.Attribute("BaseLoc")?.Value)) ?? 0) > options.MaxPageCount)
            throw new InvalidDataException("OFD document exceeds the configured page count limit.");
        material.ReserveReadModel(256L + (documentXml.Root?.Element(docNs + "Pages")?.Elements().Count() ?? 0) * 128L, cancellationToken);
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

        material.ReserveReadModel(512L + pages.Count * 128L, cancellationToken);
        var annotationIndex = IndexAnnotationFiles(archive, documentXml, docRoot,
            new HashSet<string>(pages.Select(page => page.Id).OfType<string>().Where(id => !string.IsNullOrWhiteSpace(id)), StringComparer.Ordinal), material, cancellationToken);
        var annotationDocuments = new Dictionary<string, XDocument?>(StringComparer.OrdinalIgnoreCase);
        var annotationBudget = new AnnotationReadBudget(options);
        foreach (var pageRef in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var contentPath = material.ResolvePath(docRoot, pageRef.BaseLoc!, cancellationToken);
            var pageXml = material.ReadXml(contentPath, archive.GetBytes(contentPath), cancellationToken);
            var pageNs = pageXml.Root?.Name.Namespace ?? docNs;

            var box = ParseBox(pageXml.Root?
                .Element(pageNs + "Area")?
                .Element(pageNs + "PhysicalBox") is XElement physicalBox ? ReadValue(physicalBox, material, cancellationToken) : null, material, cancellationToken);

            material.ReserveReadModel(1024, cancellationToken);
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

            material.ReserveReadModel(1024L + contentPath.Length * 16L + documentMediaMap.Count * 160L + documentMediaTypeMap.Count * 160L, cancellationToken);
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
                material.ReserveReadModel(1024L + templatePath.Length * 16L + documentMediaMap.Count * 160L + documentMediaTypeMap.Count * 160L, cancellationToken);
                var templateResPath = $"{GetDirectory(templatePath)}/PageRes.xml";
                var templateMediaMap = new Dictionary<string, string>(documentMediaMap, StringComparer.OrdinalIgnoreCase);
                var templateMediaTypeMap = new Dictionary<string, string>(documentMediaTypeMap, StringComparer.OrdinalIgnoreCase);
                ReadMediaResources(archive, templateResPath, docNs, templateMediaMap, templateMediaTypeMap, material, cancellationToken);
                ReadFonts(archive, templateResPath, package, fontMap, material, cancellationToken);

                material.ReserveReadModel(512, cancellationToken);
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
                    fontMap,
                    templateMediaMap,
                    templateMediaTypeMap,
                    templateResPath, material, cancellationToken))
                {
                    template.Elements.Add(element);
                }

                page.Templates.Add(template);
            }

            ReadAnnotationAppearances(archive, annotationIndex, annotationDocuments, annotationBudget, page, fontMap, mediaMap, mediaTypeMap, material, cancellationToken);
            package.Pages.Add(page);
        }

        var attachmentsLoc = ReadValue(documentXml.Root?.Elements().FirstOrDefault(x => x.Name.LocalName == "Attachments"), material, cancellationToken);
        if (!string.IsNullOrWhiteSpace(attachmentsLoc))
        {
            var attachmentsPath = material.ResolvePath(docRoot, attachmentsLoc!, cancellationToken);
            if (archive.Contains(attachmentsPath))
            {
                var attachmentsXml = material.ReadXml(attachmentsPath, archive.GetBytes(attachmentsPath), cancellationToken);
                foreach (var node in attachmentsXml.Root?.Elements()
                    .Where(x => x.Name.LocalName == "Attachment") ?? Enumerable.Empty<XElement>())
                {
                    var fileLoc = ReadValue(node.Elements().FirstOrDefault(x => x.Name.LocalName == "FileLoc"), material, cancellationToken)
                        ?? node.Attribute("FileLoc")?.Value
                        ?? string.Empty;
                    var contentPath = material.ResolvePath(attachmentsPath, fileLoc, cancellationToken);
                    var external = bool.TryParse(node.Attribute("External")?.Value, out var legacyExternal) && legacyExternal;
                    external = external || !archive.Contains(contentPath);

                    material.ReserveReadModel(512, cancellationToken);
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

        var customTagsLoc = ReadValue(documentXml.Root?.Elements().FirstOrDefault(x => x.Name.LocalName == "CustomTags"), material, cancellationToken);
        if (!string.IsNullOrWhiteSpace(customTagsLoc))
        {
            var customTagsPath = material.ResolvePath(docRoot, customTagsLoc!, cancellationToken);
            if (archive.Contains(customTagsPath))
            {
                var customTagsXml = material.ReadXml(customTagsPath, archive.GetBytes(customTagsPath), cancellationToken);
                foreach (var customTag in customTagsXml.Root?.Elements()
                    .Where(x => x.Name.LocalName == "CustomTag") ?? Enumerable.Empty<XElement>())
                {
                    var detailLoc = ReadValue(customTag.Elements().FirstOrDefault(x => x.Name.LocalName == "FileLoc"), material, cancellationToken)
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
            var legacyCustomTagPath = material.ReadConcat(package.Options.DocumentId, "/Tags/CustomTag_EMR.xml", cancellationToken);
            ReadKeyValueCustomTags(archive, legacyCustomTagPath, package.CustomTags, material, cancellationToken);
        }

        material.ReserveReadModel(1024L + (documentXml.Root?.DescendantNodesAndSelf().Count() ?? 0) * 4096L, cancellationToken);
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
        material.ReserveReadModel(1024, cancellationToken);
        var result = new AnnotationIndex { Namespace = document.Root!.Name.Namespace };
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in document.Root!.Elements().Where(node => node.Name.LocalName == "Annotations"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (declaration.Name.Namespace != result.Namespace)
            {
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(declaration, material, cancellationToken));
                continue;
            }
            var literal = ReadJoinText(declaration.Nodes().OfType<XText>().Select(text => text.Value), material, cancellationToken);
            if (declaration.HasElements || declaration.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration) ||
                declaration.Nodes().Any(node => node is not XText and not XComment))
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(declaration, material, cancellationToken));
            if (string.IsNullOrWhiteSpace(literal))
            {
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(declaration, material, cancellationToken));
                continue;
            }
            var listPath = material.ResolvePath(documentPath, literal, cancellationToken);
            if (!visited.Add(listPath)) continue;
            if (!archive.Contains(listPath))
            {
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(declaration, material, cancellationToken));
                continue;
            }
            XDocument list;
            try { list = material.ReadXml(listPath, archive.GetBytes(listPath), cancellationToken); }
            catch (XmlException)
            {
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(declaration, material, cancellationToken));
                continue;
            }
            if (!IsAnnotationRoot(list, "Annotations", result.Namespace))
            {
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(declaration, material, cancellationToken));
                continue;
            }
            if (!PlainContainer(list.Root!) || list.Nodes().Any(node => node is not XElement and not XComment && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                result.UnmodeledLists.Add(OfdXmlEnvelope.SerializeForRead(list, material, cancellationToken));
            foreach (var record in list.Root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                var pageId = record.Attribute("PageID")?.Value;
                material.ReserveReadModel(256L + record.Elements().Count() * 32L, cancellationToken);
                var locations = record.Elements(list.Root.Name.Namespace + "FileLoc").ToList();
                var location = locations.Count == 1 && !locations[0].HasElements &&
                    locations[0].Nodes().All(node => node is XText or XComment)
                    ? ReadJoinText(locations[0].Nodes().OfType<XText>().Select(text => text.Value), material, cancellationToken) : null;
                AnnotationFiles? files = null;
                var raw = OfdXmlEnvelope.SerializeForRead(record, material, cancellationToken);
                if (!string.IsNullOrWhiteSpace(pageId))
                {
                    material.ReserveReadModel(512, cancellationToken);
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
        material.ReserveReadModel(256L + result.Files.Count * 64L, cancellationToken);
        foreach (var orphan in result.Files.Where(pair => !loadedPageIds.Contains(pair.Key)).ToList())
        {
            result.UnmodeledLists.AddRange(orphan.Value.Records);
            result.Files.Remove(orphan.Key);
        }
        result.UnmodeledXml = WrapUnmodeled(result.UnmodeledLists, material, cancellationToken);
        foreach (var files in result.Files.Values) files.UnmodeledXml = WrapUnmodeled(files.UnmodeledRecords, material, cancellationToken);
        return result;
    }

    private static string? WrapUnmodeled(List<string> records, OfdModelMaterialization material, CancellationToken token)
    {
        if (records.Count == 0) return null;
        if (!material.IsControlledRead) return "<UnmodeledAnnotations>" + string.Concat(records) + "</UnmodeledAnnotations>";
        material.ReserveReadModel(128L + records.Count * 32L, token);
        return ReadJoinText(new[] { "<UnmodeledAnnotations>", ReadJoinText(records, material, token), "</UnmodeledAnnotations>" }, material, token);
    }

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
        OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        material.ReserveReadModel(512, cancellationToken);
        if (index.UnmodeledXml is not null)
            page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "Annotations", Xml = index.UnmodeledXml });
        if (page.Id is null || !index.Files.TryGetValue(page.Id, out var files)) return;
        if (files.UnmodeledXml is not null)
            page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "Annotations", Xml = files.UnmodeledXml });
        foreach (var annotationPath in files.Paths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            material.ReserveReadModel(1024L + 16L * annotationPath.Length, cancellationToken);
            // Cached XML still produces new typed/raw strings on every page.
            // Charge each occurrence before parsing or materializing its content.
            if (archive.TryGetBytes(annotationPath, out var annotationBytes)) budget.ChargeXml(annotationBytes.LongLength);
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
                    Xml = ReadXmlString(new XElement("UnmodeledAnnotation", new XAttribute("FileLoc", annotationPath)), material, cancellationToken)! });
                continue;
            }
            if (!PlainContainer(annotations.Root!) || annotations.Nodes().Any(node => node is not XElement and not XComment && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(annotations, material, cancellationToken) });
            foreach (var annotation in annotations.Root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                material.ReserveReadModel(1024, cancellationToken);
                if (annotation.Name != annotations.Root.Name.Namespace + "Annot")
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken) });
                    continue;
                }
                if (!SupportedAnnotationWrapper(annotation))
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken) });
                if (annotation.Attribute("Visible")?.Value is "false" or "0")
                {
                    var hiddenAppearance = annotation.Element(annotations.Root.Name.Namespace + "Appearance");
                    if (hiddenAppearance is not null)
                    {
                        var hiddenNodes = FlattenAnnotationNodes(hiddenAppearance.Elements(), hiddenAppearance.Name.Namespace, budget, material, cancellationToken);
                        if (!PlainContainer(hiddenAppearance, "Boundary", "ID", "CTM") || hiddenNodes is null ||
                            !TryParseBox(hiddenAppearance.Attribute("Boundary")?.Value, material, cancellationToken, out _) ||
                            hiddenNodes.SelectMany(node => node.DescendantsAndSelf().Attributes()).Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute) ||
                                attribute.Name == "Boundary" && !TryParseBox(attribute.Value, material, cancellationToken, out _)))
                            page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken) });
                    }
                    continue;
                }
                material.ReserveReadModel(256L + annotation.Elements().Count() * 32L, cancellationToken);
                var appearances = annotation.Elements(annotations.Root.Name.Namespace + "Appearance").ToList();
                if (appearances.Count > 1)
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken) });
                    continue;
                }
                var appearance = appearances.Count == 1 ? appearances[0] : null;
                if (appearance is null)
                {
                    if (annotation.Elements().Any(node => node.Name.LocalName == "Appearance"))
                        page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken) });
                    continue;
                }
                var appearanceTransform = ParseMatrix(appearance.Attribute("CTM")?.Value, material, cancellationToken);
                var hasUnmodeledMetadata = !PlainContainer(appearance, "Boundary", "ID", "CTM") ||
                    !TryParseBox(appearance.Attribute("Boundary")?.Value, material, cancellationToken, out _);
                if (appearance.Attribute("CTM") is not null && (appearanceTransform is null || appearanceTransform.Any(value => double.IsNaN(value) || double.IsInfinity(value))))
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
                    continue;
                }
                var ns = appearance.Name.Namespace;
                var box = ParseBox(appearance.Attribute("Boundary")?.Value, material, cancellationToken);
                var nodes = FlattenAnnotationNodes(appearance.Elements(), appearance.Name.Namespace, budget, material, cancellationToken);
                if (nodes is not null && nodes.SelectMany(node => node.DescendantsAndSelf().Attributes("Boundary")).Any(attribute => !TryParseBox(attribute.Value, material, cancellationToken, out _)))
                    hasUnmodeledMetadata = true;
                if (nodes is null || box.w <= 0 || box.h <= 0 || new[] { box.x, box.y, box.w, box.h }.Any(value => double.IsNaN(value) || double.IsInfinity(value)))
                {
                    page.AnnotationAppearances.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
                    continue;
                }
                var layerId = material.ReadConcat("annotation-", annotation.Attribute("ID")?.Value ?? string.Empty, cancellationToken);
                material.PrepareReadAnnotationClone(nodes, cancellationToken);
                var xml = new XDocument(new XElement(ns + "Page", new XElement(ns + "Content",
                    new XElement(ns + "Layer", new XAttribute("ID", layerId), nodes))));
                var staged = new List<OfdElement>();
                if (hasUnmodeledMetadata) staged.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
                if (nodes.SelectMany(node => node.DescendantsAndSelf().Attributes()).Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute)))
                    staged.Add(new OfdRawElement { LocalName = "UnmodeledAnnotationMetadata", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
                foreach (var element in ParsePageObjects(archive, xml, fonts, media, mediaTypes, annotationPath, material, cancellationToken))
                {
                    if (element is OfdImageElement { Data.Length: 0 })
                    {
                        staged.Clear();
                        staged.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
                        break;
                    }
                    material.ReserveReadModel(16384, cancellationToken);
                    var transform = appearanceTransform ?? new double[] { 1, 0, 0, 1, 0, 0 };
                    var x = element.XMillimeters; var y = element.YMillimeters;
                    var originalWidth = element.WidthMillimeters; var originalHeight = element.HeightMillimeters;
                    var minX = Math.Min(0, transform[0] * originalWidth) + Math.Min(0, transform[2] * originalHeight);
                    var minY = Math.Min(0, transform[1] * originalWidth) + Math.Min(0, transform[3] * originalHeight);
                    element.XMillimeters = box.x + transform[0] * x + transform[2] * y + transform[4] + minX;
                    element.YMillimeters = box.y + transform[1] * x + transform[3] * y + transform[5] + minY;
                    element.WidthMillimeters = Math.Abs(transform[0]) * originalWidth + Math.Abs(transform[2]) * originalHeight;
                    element.HeightMillimeters = Math.Abs(transform[1]) * originalWidth + Math.Abs(transform[3]) * originalHeight;
                    if (appearanceTransform is not null)
                    {
                        var inner = element switch
                        {
                            OfdTextElement text => text.Transform,
                            OfdPathElement path => path.Transform,
                            OfdImageElement image => image.Transform ?? new double[] { originalWidth, 0, 0, originalHeight, 0, 0 },
                            _ => null
                        } ?? new double[] { 1, 0, 0, 1, 0, 0 };
                        var combined = new double[]
                        {
                            transform[0] * inner[0] + transform[2] * inner[1], transform[1] * inner[0] + transform[3] * inner[1],
                            transform[0] * inner[2] + transform[2] * inner[3], transform[1] * inner[2] + transform[3] * inner[3],
                            transform[0] * inner[4] + transform[2] * inner[5] - minX, transform[1] * inner[4] + transform[3] * inner[5] - minY
                        };
                        if (element is OfdTextElement transformedText) transformedText.Transform = combined;
                        if (element is OfdImageElement transformedImage) transformedImage.Transform = combined;
                        if (element is OfdPathElement transformedPath) transformedPath.Transform = combined;
                    }
                    try
                    {
                        element.ClippingXml = TransformAnnotationClips(element.ClippingXml, ns, transform, minX, minY,
                            box.x + transform[4] - element.XMillimeters, box.y + transform[5] - element.YMillimeters, box.w, box.h,
                            element is OfdImageElement ? originalWidth : 0, element is OfdImageElement ? originalHeight : 0, material, cancellationToken);
                    }
                    catch (Exception exception) when (exception is NotSupportedException or InvalidDataException && !OfdTraversalBudget.IsLimit(exception))
                    {
                        staged.Clear();
                        staged.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
                        break;
                    }
                    // Keep the serialized boundary consistent with the translated typed geometry.
                    var source = element switch { OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, _ => null };
                    if (source is not null)
                    {
                        material.ReserveReadModel(16384L + source.Length * 8L, cancellationToken);
                        material.PrepareReadXmlParser(checked(2L * source.Length), cancellationToken);
                        var node = XElement.Parse(source);
                        node.SetAttributeValue("Boundary", string.Join(" ", new[] { element.XMillimeters, element.YMillimeters, element.WidthMillimeters, element.HeightMillimeters }.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));
                        if (appearanceTransform is not null)
                        {
                            var objectTransform = element switch { OfdTextElement matrixText => matrixText.Transform, OfdImageElement matrixImage => matrixImage.Transform, OfdPathElement matrixPath => matrixPath.Transform, _ => null };
                            if (objectTransform is not null) node.SetAttributeValue("CTM", string.Join(" ", objectTransform.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));
                        }
                        source = OfdXmlEnvelope.Capture(node, material, cancellationToken);
                        if (element is OfdTextElement text) text.SourceXml = source;
                        if (element is OfdImageElement image) image.SourceXml = source;
                        if (element is OfdPathElement path) path.SourceXml = source;
                    }
                    staged.Add(element);
                }
                page.AnnotationAppearances.AddRange(staged);
            }
        }
    }

    private static List<XElement>? FlattenAnnotationNodes(IEnumerable<XElement> nodes, XNamespace ns, AnnotationReadBudget budget, OfdModelMaterialization material, CancellationToken token)
    {
        material.ReserveReadModel(128, token);
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
                    budget.ChargeObject();
                    if (!IsSupportedAnnotationPrimitive(node, ns)) return false;
                    material.ReserveReadModel(64, token);
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

    private static string TransformAnnotationClips(string? original, XNamespace ns, double[] outer,
        double minX, double minY, double outerX, double outerY, double width, double height, double imageWidth, double imageHeight,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ReserveReadModel(32768L + (original?.Length ?? 0) * 32L, token);
        if (original is not null) material.GuardFragmentXml(original, token);
        var clips = new XElement(ns + "Clips");
        if (!string.IsNullOrWhiteSpace(original))
        {
            material.PrepareReadXmlParser(checked(2L * original!.Length), token);
            var source = XElement.Parse(original!);
            if (source.Name.Namespace != ns || !OfdGraphicXmlContract.HasKnownChildren(source) || OfdGraphicXmlContract.HasUnsupportedReferences(source) ||
                source.DescendantsAndSelf().Attributes().Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute)))
                throw new NotSupportedException("Unsupported annotation clip cannot be flattened safely.");
            material.PrepareReadXmlParser(checked(2L * original!.Length), token);
            foreach (var region in OfdClipGeometry.Read(original))
            {
                var area = new XElement(ns + "Area");
                foreach (var path in region.Paths)
                {
                    var inner = path.Transform ?? new double[] { 1, 0, 0, 1, 0, 0 };
                    var matrix = new[] { outer[0] * inner[0] + outer[2] * inner[1], outer[1] * inner[0] + outer[3] * inner[1],
                        outer[0] * inner[2] + outer[2] * inner[3], outer[1] * inner[2] + outer[3] * inner[3],
                        outer[0] * inner[4] + outer[2] * inner[5] - minX, outer[1] * inner[4] + outer[3] * inner[5] - minY };
                    area.Add(ClipPath(path.AbbreviatedData, matrix, region.EvenOdd));
                }
                clips.Add(new XElement(ns + "Clip", area));
            }
        }
        if (imageWidth > 0 && imageHeight > 0)
        {
            var imageRectangle = $"M 0 0 L {Number(imageWidth)} 0 L {Number(imageWidth)} {Number(imageHeight)} L 0 {Number(imageHeight)} C";
            clips.Add(new XElement(ns + "Clip", new XElement(ns + "Area", ClipPath(imageRectangle,
                new[] { outer[0], outer[1], outer[2], outer[3], -minX, -minY }, false))));
        }
        // Separate Clip means intersection with existing clipping, not union.
        var rectangle = $"M 0 0 L {Number(width)} 0 L {Number(width)} {Number(height)} L 0 {Number(height)} C";
        clips.Add(new XElement(ns + "Clip", new XElement(ns + "Area", ClipPath(rectangle,
            new[] { outer[0], outer[1], outer[2], outer[3], outerX, outerY }, false))));
        return OfdXmlEnvelope.SerializeForRead(clips, material, token);

        string Number(double value) => value.ToString("R", CultureInfo.InvariantCulture);
        XElement ClipPath(string data, double[] matrix, bool evenOdd) => new(ns + "Path",
            new XAttribute("ID", "annotation-clip"), new XAttribute("CTM", string.Join(" ", matrix.Select(Number))),
            evenOdd ? new XAttribute("Rule", "Even-Odd") : null, new XElement(ns + "AbbreviatedData", data));
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
            var file = ReadValue(font.Elements().FirstOrDefault(node => node.Name.LocalName == "FontFile"), material, cancellationToken);
            material.ReserveReadModel(512, cancellationToken);
            var resource = new OfdFontResource
            {
                Id = id!, FontName = name!, FamilyName = font.Attribute("FamilyName")?.Value,
                Charset = font.Attribute("Charset")?.Value,
                Bold = ParseBoolean(font.Attribute("Bold")?.Value, false),
                Italic = ParseBoolean(font.Attribute("Italic")?.Value, false), FileName = file
            };
            if (!string.IsNullOrWhiteSpace(file))
            {
                material.ReserveReadModel(256L + 16L * ((baseLocation?.Length ?? 0) + file!.Length), cancellationToken);
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
            var file = media.Attribute("MediaFile")?.Value ?? ReadValue(media.Element(resNs + "MediaFile"), material, cancellationToken);
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(file))
            {
                continue;
            }

            material.ReserveReadModel(512L + 16L * ((baseLoc?.Length ?? 0) + file!.Length + (media.Attribute("Format")?.Value.Length ?? 0)), cancellationToken);
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
        material.ReserveReadModel(256L + (pageXml.Root?.Descendants().Count() ?? 0) * 32L, cancellationToken);
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
                material.ReserveReadModel(512, cancellationToken);
                var localName = node.Name.LocalName;
                if (string.Equals(localName, "TextObject", StringComparison.OrdinalIgnoreCase))
                {
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
                    material.ReserveReadModel(256L + node.Elements().Count() * 160L, cancellationToken);
                    var textCodes = node.Elements(node.Name.Namespace + "TextCode").ToList();
                    var textParts = textCodes.Select(code => ReadDirectText(code, material, cancellationToken)).ToArray();
                    var text = new OfdTextElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        Text = ReadJoinText(textParts, material, cancellationToken),
                        FontResourceId = node.Attribute("Font")?.Value,
                        FontName = ResolveFontName(node.Attribute("Font")?.Value, fontMap),
                        FontSizeMillimeters = ParseDouble(node.Attribute("Size")?.Value, 4d),
                        Weight = (int)ParseDouble(node.Attribute("Weight")?.Value, OfdTextElement.DefaultWeight),
                        Italic = ParseBoolean(node.Attribute("Italic")?.Value, false),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value, material, cancellationToken),
                        FillColor = ParseColor(node.Elements()
                            .FirstOrDefault(x => x.Name.LocalName == "FillColor"), material, cancellationToken) ?? OfdColor.Black,
                        SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken),
                        ClippingXml = ReadXmlString(node.Element(node.Name.Namespace + "Clips"), material, cancellationToken)
                    };
                    foreach (var textCode in textCodes)
                    {
                        material.ReserveReadModel(256, cancellationToken);
                        text.Runs.Add(new OfdTextRun
                        {
                            Text = ReadDirectText(textCode, material, cancellationToken),
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
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
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
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value, material, cancellationToken),
                        Alpha = (int)Math.Max(0, Math.Min(255, ParseDouble(node.Attribute("Alpha")?.Value, 255))),
                        ClipsXml = ReadXmlString(node.Element(node.Name.Namespace + "Clips"), material, cancellationToken),
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
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
                    var strokeColor = ParseColor(node.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "StrokeColor"), material, cancellationToken);
                    var fillColor = ParseColor(node.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "FillColor"), material, cancellationToken);
                    yield return new OfdPathElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        AbbreviatedData = ReadDirectText(node.Element(node.Name.Namespace + "AbbreviatedData"), material, cancellationToken),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value, material, cancellationToken),
                        LineWidthMillimeters = ParseDouble(node.Attribute("LineWidth")?.Value, 0.353d),
                        Stroke = ParseBoolean(node.Attribute("Stroke")?.Value, true),
                        Fill = ParseBoolean(node.Attribute("Fill")?.Value, false),
                        StrokeColor = strokeColor ?? OfdColor.Black,
                        FillColor = fillColor,
                        SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken),
                        ClippingXml = ReadXmlString(node.Element(node.Name.Namespace + "Clips"), material, cancellationToken)
                    };
                    continue;
                }

                var rawBoundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
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

    private static string ReadDirectText(XElement? node, OfdModelMaterialization material, CancellationToken token)
    {
        if (node is null) return string.Empty;
        material.ReserveReadModel(512, token);
        var values = node.Nodes().OfType<XText>().Where(text => text is XCData || !string.IsNullOrWhiteSpace(text.Value) ||
            node.AncestorsAndSelf().Select(element => element.Attribute(XNamespace.Xml + "space")?.Value).FirstOrDefault(value => value is not null) == "preserve").Select(text => text.Value);
        return ReadJoinText(values, material, token);
    }
    private static string ReadJoinText(IEnumerable<string> values, OfdModelMaterialization material, CancellationToken token)
    {
        if (!material.IsControlledRead) return string.Concat(values);
        long length = 0;
        foreach (var value in values) { token.ThrowIfCancellationRequested(); length = checked(length + value.Length); }
        material.ChargeCopy(checked(256L + 6L * length), token); material.ChargeWork(checked(4L * length), token);
        var storage = material.ReserveReadStorage(checked(2L * length), token);
        var result = string.Concat(values); material.AdoptReadString(result, storage, token); return result;
    }
    private static string? ReadValue(XElement? node, OfdModelMaterialization material, CancellationToken token)
        => node is null ? null : material.ReadXmlValue(node, token);
    private static string? ReadXmlString(XNode? node, OfdModelMaterialization material, CancellationToken token)
        => node is null ? null : OfdXmlEnvelope.SerializeForRead(node, material, token);


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
                material.ReserveReadModel(160, cancellationToken);
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

    private static bool TryParseBox(string? value, OfdModelMaterialization material, CancellationToken token,
        out (double x, double y, double w, double h) box)
    {
        material.PrepareReadNumericParse(value, token);
        return OfdBoxParser.TryParse(value, out box);
    }

    private static (double x, double y, double w, double h) ParseBox(string? value, OfdModelMaterialization material, CancellationToken token)
    {
        if (value is null) return default;
        material.PrepareReadNumericParse(value, token);
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

    private static double[]? ParseMatrix(string? value, OfdModelMaterialization material, CancellationToken token)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        material.ReserveReadModel(512L + value!.Length * 32L, token);
        var parts = value!.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length != 6)
        {
            return null;
        }

        var matrixStorage = material.IsControlledRead ? material.ReserveReadStorage(48, token) : null;
        var matrix = new double[6];
        for (var i = 0; i < matrix.Length; i++)
        {
            if (!double.TryParse(parts[i], NumberStyles.Float, CultureInfo.InvariantCulture, out matrix[i]))
            {
                return null;
            }
        }

        if (matrixStorage is not null) material.AdoptReadStorage(matrix, 48, matrixStorage, token);
        return matrix;
    }

    private static OfdColor? ParseColor(XElement? colorElement, OfdModelMaterialization material, CancellationToken token)
    {
        var value = colorElement?.Attribute("Value")?.Value;
        if (string.IsNullOrWhiteSpace(value))
        {
            return null;
        }

        material.ReserveReadModel(512L + value!.Length * 32L, token);
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
