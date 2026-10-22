using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Interfaces;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Reader.Readers;

public sealed class OfdReader : IOfdReader
{
    private readonly OfdPackageLoader _loader = new();
    private static OfdRawElement AnnotationMetadata(string xml) => new()
    {
        LocalName = "UnmodeledAnnotationMetadata", Xml = xml, ReaderAnnotationMetadataXml = xml
    };
    private static OfdRawElement AnnotationIndexMetadata(string xml) => new()
    {
        LocalName = "Annotations", Xml = xml, ReaderAnnotationMetadataXml = xml
    };

    /// <summary>Lists DocBody declarations in source order without materializing their pages.</summary>
    public async Task<IReadOnlyList<OfdDocumentDescriptor>> ListDocumentsAsync(Stream ofdStream,
        OfdPackageLoadOptions? options = null, CancellationToken cancellationToken = default)
    {
        var archive = await _loader.LoadAsync(ofdStream, options ?? new OfdPackageLoadOptions(), cancellationToken).ConfigureAwait(false);
        var material = new OfdModelMaterialization(new OfdModelLimits());
        var xml = material.ReadXml("OFD.xml", archive.GetBytes("OFD.xml"), cancellationToken);
        var documents = DocumentBodies(xml, material, cancellationToken).Select(body => body.Descriptor).ToArray();
        if (documents.Any(document => !archive.Contains(document.DocumentEntryPath)))
            throw new InvalidDataException("A DocBody references a missing document entry.");
        return documents;
    }

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
        try
        {
            return await ReadCoreAsync(ofdStream, options, cancellationToken).ConfigureAwait(false);
        }
        catch (XmlException error)
        {
            throw new InvalidDataException("OFD package contains malformed XML.", error);
        }
        catch (KeyNotFoundException error)
        {
            throw new InvalidDataException("OFD package references a missing entry.", error);
        }
    }

    private async Task<OfdDocumentPackage> ReadCoreAsync(
        Stream ofdStream,
        OfdPackageLoadOptions options,
        CancellationToken cancellationToken)
    {
        var snapshot = OfdPackageLoader.SnapshotOptions(options);
        var archive = await _loader.LoadAsync(ofdStream, snapshot, cancellationToken).ConfigureAwait(false);
        return ReadArchive(archive, snapshot, new OfdModelMaterialization(new OfdModelLimits()), cancellationToken);
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

        var bodies = DocumentBodies(ofdXml, material, cancellationToken);
        if (options.DocumentIndex >= bodies.Count && !(options.DocumentIndex == 0 && bodies.Count == 0))
            throw new ArgumentOutOfRangeException(nameof(options), "DocumentIndex is outside the DocBody list.");
        var docRoot = bodies.Count == 0 ? "Doc_0/Document.xml" : bodies[options.DocumentIndex].Descriptor.DocumentEntryPath;
        package.Options.DocType = docType;
        package.DocumentEntryPath = docRoot;
        package.Options.Namespace = ofdNs.NamespaceName;
        material.ReserveReadModel(256L + docRoot.Length * 16L, cancellationToken);
        var docIdStorage = material.IsControlledRead ? material.ReserveReadStorage(2L * docRoot.Length, cancellationToken) : null;
        package.Options.DocumentId = docRoot.Split('/')[0];
        if (docIdStorage is not null) material.AdoptReadString(package.Options.DocumentId, docIdStorage, cancellationToken);

        var selectedDocBody = bodies.Count == 0 ? null : bodies[options.DocumentIndex].Element;
        var info = selectedDocBody?.Element(ofdNs + "DocInfo");
        if (info is not null)
        {
            package.Options.Metadata.Title = ReadValue(info.Element(ofdNs + "Title"), material, cancellationToken);
            package.Options.Metadata.Author = ReadValue(info.Element(ofdNs + "Author"), material, cancellationToken);
            package.Options.Metadata.Subject = ReadValue(info.Element(ofdNs + "Subject"), material, cancellationToken);
            package.Options.Metadata.Keywords = ReadValue(info.Element(ofdNs + "Keywords"), material, cancellationToken);
            package.Options.Metadata.Creator = ReadValue(info.Element(ofdNs + "Creator"), material, cancellationToken);
            OfdStandardXml.ReadMetadata(package, info, material, cancellationToken);
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
        var documentAmbiguousMedia = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var documentPaint = new OfdPaintResources();

        var publicResLoc = ReadValue(commonData?.Element(docNs + "PublicRes"), material, cancellationToken);
        package.PublicResourceLocation = publicResLoc;
        if (!string.IsNullOrWhiteSpace(publicResLoc))
        {
            var publicResPath = material.ResolvePath(docRoot, publicResLoc!, cancellationToken);
            if (archive.Contains(publicResPath))
            {
                ReadFonts(archive, publicResPath, package, fontMap, material, cancellationToken);
                ReadMediaResources(archive, publicResPath, documentMediaMap, documentMediaTypeMap, documentAmbiguousMedia, material, cancellationToken);
                documentPaint.Read(archive, publicResPath, material, cancellationToken);
            }
        }

        var documentResLoc = ReadValue(commonData?.Element(docNs + "DocumentRes"), material, cancellationToken);
        package.DocumentResourceLocation = documentResLoc;
        if (!string.IsNullOrWhiteSpace(documentResLoc))
        {
            var documentResPath = material.ResolvePath(docRoot, documentResLoc!, cancellationToken);
            ReadMediaResources(archive, documentResPath, documentMediaMap, documentMediaTypeMap, documentAmbiguousMedia, material, cancellationToken);
            ReadFonts(archive, documentResPath, package, fontMap, material, cancellationToken);
            documentPaint.Read(archive, documentResPath, material, cancellationToken);
        }

        var compositeMap = ReadCompositeResources(archive, package, docRoot, docNs, fontMap, documentMediaMap,
            documentMediaTypeMap, documentAmbiguousMedia, documentPaint, material, cancellationToken);
        material.ReserveReadModel(256L + fontMap.Count * 160L, cancellationToken);
        var documentFonts = new Dictionary<string, string>(fontMap, StringComparer.OrdinalIgnoreCase);
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
        material.ReserveReadModel(256, cancellationToken);
        var annotationIdCounts = new Dictionary<long, int>();
        foreach (var entry in package.PreservedEntries)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var reservedIds = material.Reservations(entry.Value, out _, cancellationToken);
            if (!reservedIds.Any() || material.OpaqueEntries.Contains(entry.Key)) continue;
            foreach (var attribute in material.ReadXml(entry.Key, entry.Value, cancellationToken).Descendants().Attributes("ID"))
            {
                material.Structure.Reference(); var id = OfdSourceIdentity.ParseId(attribute.Value);
                if (id != 0)
                {
                    if (!annotationIdCounts.ContainsKey(id)) material.ReserveReadModel(128, cancellationToken);
                    annotationIdCounts[id] = annotationIdCounts.TryGetValue(id, out var count) ? count + 1 : 1;
                }
            }
        }
        package.AnnotationIndexSource = annotationIndex.Source; annotationIndex.Source.DocumentOwner = package;
        var annotationDocuments = new Dictionary<string, XDocument?>(StringComparer.OrdinalIgnoreCase);
        var annotationBudget = new AnnotationReadBudget(options);
        foreach (var pageRef in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            material.ReserveReadModel(256L + documentFonts.Count * 160L, cancellationToken);
            fontMap = new Dictionary<string, string>(documentFonts, StringComparer.OrdinalIgnoreCase);
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
            var ambiguousMedia = new HashSet<string>(documentAmbiguousMedia, StringComparer.OrdinalIgnoreCase);
            var pagePaint = new OfdPaintResources(documentPaint);

            if (archive.Contains(pageResPath))
            {
                ReadMediaResources(archive, pageResPath, mediaMap, mediaTypeMap, ambiguousMedia, material, cancellationToken);
                ReadFonts(archive, pageResPath, package, fontMap, material, cancellationToken);
                pagePaint.Read(archive, pageResPath, material, cancellationToken);
            }
            var pageCompositeMap = pageXml.Root is { } pageRoot && OfdPageXmlContract.ContainsPageBlock(pageRoot)
                ? ReadScopedCompositeResources(archive, package, pageResPath, docNs, compositeMap,
                    fontMap, mediaMap, mediaTypeMap, ambiguousMedia, pagePaint, material, cancellationToken)
                : compositeMap;

            foreach (var element in ParsePageObjects(
                archive,
                pageXml,
                fontMap,
                mediaMap,
                mediaTypeMap,
                pageResPath, material, cancellationToken, compositeMap, pagePaint, ambiguousMedia, pageCompositeMap))
            {
                page.Elements.Add(element);
            }

            material.ReserveReadModel(256L + fontMap.Count * 160L, cancellationToken);
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
                material.ReserveReadModel(1024L + templatePath.Length * 16L + documentMediaMap.Count * 160L + documentMediaTypeMap.Count * 160L, cancellationToken);
                var templateResPath = $"{GetDirectory(templatePath)}/PageRes.xml";
                var templateMediaMap = new Dictionary<string, string>(documentMediaMap, StringComparer.OrdinalIgnoreCase);
                var templateMediaTypeMap = new Dictionary<string, string>(documentMediaTypeMap, StringComparer.OrdinalIgnoreCase);
                var templateAmbiguousMedia = new HashSet<string>(documentAmbiguousMedia, StringComparer.OrdinalIgnoreCase);
                var templatePaint = new OfdPaintResources(documentPaint);
                ReadMediaResources(archive, templateResPath, templateMediaMap, templateMediaTypeMap, templateAmbiguousMedia, material, cancellationToken);
                material.ReserveReadModel(256L + documentFonts.Count * 160L, cancellationToken);
                var templateFonts = new Dictionary<string, string>(documentFonts, StringComparer.OrdinalIgnoreCase);
                ReadFonts(archive, templateResPath, package, templateFonts, material, cancellationToken);
                templatePaint.Read(archive, templateResPath, material, cancellationToken);
                var templateCompositeMap = templateXml.Root is { } templateRoot && OfdPageXmlContract.ContainsPageBlock(templateRoot)
                    ? ReadScopedCompositeResources(archive, package, templateResPath, docNs, compositeMap,
                        templateFonts, templateMediaMap, templateMediaTypeMap, templateAmbiguousMedia, templatePaint, material, cancellationToken)
                    : compositeMap;

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
                    templateFonts,
                    templateMediaMap,
                    templateMediaTypeMap,
                    templateResPath, material, cancellationToken, compositeMap, templatePaint, templateAmbiguousMedia, templateCompositeMap))
                {
                    template.Elements.Add(element);
                }

                page.Templates.Add(template);
            }

            ReadAnnotationAppearances(archive, annotationIndex, annotationDocuments, annotationBudget, page, annotationFonts, mediaMap, mediaTypeMap, material, cancellationToken, annotationIdCounts);
            package.Pages.Add(page);
        }

        XElement? attachmentDeclaration = null;
        foreach (var declaration in documentXml.Root!.Elements(docNs + "Attachments"))
        {
            cancellationToken.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (attachmentDeclaration is not null) throw new InvalidDataException("Duplicate attachment descriptor references.");
            attachmentDeclaration = declaration;
        }
        var attachmentsLoc = ReadValue(attachmentDeclaration, material, cancellationToken);
        if (!string.IsNullOrWhiteSpace(attachmentsLoc))
        {
            var attachmentsPath = material.ResolvePath(docRoot, attachmentsLoc!, cancellationToken);
            if (archive.Contains(attachmentsPath))
            {
                var attachmentsXml = material.ReadXml(attachmentsPath, archive.GetBytes(attachmentsPath), cancellationToken);
                foreach (var node in attachmentsXml.Root?.Elements()
                    .Where(x => x.Name == attachmentsXml.Root.Name.Namespace + "Attachment") ?? Enumerable.Empty<XElement>())
                {
                    var fileLoc = ReadValue(node.Element(node.Name.Namespace + "FileLoc"), material, cancellationToken)
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
                package.AttachmentSource = OfdAttachmentSource.Capture(package, attachmentDeclaration!, attachmentsPath, attachmentsXml.Root!, material, cancellationToken);
            }
        }

        var customTagDeclarations = documentXml.Root!.Elements(docNs + "CustomTags").ToArray();
        var customTagIndexes = new List<string>();
        var customTagDetails = new List<string>();
        foreach (var declaration in customTagDeclarations)
        {
            var location = ReadValue(declaration, material, cancellationToken);
            if (string.IsNullOrWhiteSpace(location)) continue;
            var listPath = material.ResolvePath(docRoot, location!, cancellationToken);
            customTagIndexes.Add(listPath);
            if (!archive.Contains(listPath)) continue;
            var list = material.ReadXml(listPath, archive.GetBytes(listPath), cancellationToken);
            if (list.Root?.Name != docNs + "CustomTags") continue;
            foreach (var customTag in list.Root.Elements(docNs + "CustomTag") ?? Enumerable.Empty<XElement>())
            {
                // Other schemas may use Tag/Key/Value too; their payloads remain opaque.
                if (customTag.Attribute("TypeID")?.Value != "EMR" ||
                    customTag.Attribute("NameSpace")?.Value != OfdCustomTagSource.EmrNamespace) continue;
                var locationNode = customTag.Element(docNs + "FileLoc");
                var detailLoc = ReadValue(locationNode, material, cancellationToken) ?? customTag.Attribute("FileLoc")?.Value;
                if (string.IsNullOrWhiteSpace(detailLoc)) continue;
                var detailPath = material.ResolvePath(listPath, detailLoc!, cancellationToken);
                if (ReadKeyValueCustomTags(archive, detailPath, package.CustomTags, docNs, material, cancellationToken))
                    customTagDetails.Add(detailPath);
            }
        }
        if (customTagDeclarations.Length == 0)
        {
            var directory = OfdPackagePath.GetDirectory(docRoot);
            var legacyPath = (directory.Length == 0 ? "" : directory + "/") + "Tags/CustomTag_EMR.xml";
            if (ReadKeyValueCustomTags(archive, legacyPath, package.CustomTags, docNs, material, cancellationToken))
                customTagDetails.Add(legacyPath);
        }
        if (customTagDeclarations.Length > 0 || customTagDetails.Count > 0) package.CustomTagSource = new OfdCustomTagSource(package,
            customTagDeclarations.Select(node => OfdXmlEnvelope.Capture(node, material, cancellationToken)).ToArray(),
            customTagIndexes, customTagDetails, material, cancellationToken);

        material.ReserveReadModel(1024L + (documentXml.Root?.DescendantNodesAndSelf().Count() ?? 0) * 4096L, cancellationToken);
        OfdNavigationXml.Read(package, documentXml.Root ?? throw new InvalidDataException("Missing document root."), material, cancellationToken);
        OfdStandardXml.ReadDocument(package, selectedDocBody, documentXml.Root!, material, cancellationToken);
        package.SourceEnvelope = OfdDocumentEnvelope.Capture(package, selectedDocBody, documentXml.Root!, material, cancellationToken);
        package.SourceIdentity = OfdSourceIdentity.Capture(package, material, cancellationToken);
        material.InspectStorage(package, cancellationToken);
        return package;
    }

    private static IReadOnlyList<(XElement Element, OfdDocumentDescriptor Descriptor)> DocumentBodies(
        XDocument xml, OfdModelMaterialization material, CancellationToken token)
    {
        var root = xml.Root ?? throw new InvalidDataException("OFD.xml has no root element.");
        var ns = root.Name.Namespace;
        var result = new List<(XElement, OfdDocumentDescriptor)>();
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var body in root.Elements(ns + "DocBody"))
        {
            token.ThrowIfCancellationRequested();
            var roots = body.Elements(ns + "DocRoot").ToArray();
            if (roots.Length != 1 || string.IsNullOrWhiteSpace(roots[0].Value))
                throw new InvalidDataException("Each DocBody must have exactly one nonblank DocRoot.");
            var path = material.ResolvePath("OFD.xml", ReadValue(roots[0], material, token)!, token);
            if (!paths.Add(path)) throw new InvalidDataException("DocBody roots are duplicated or case-aliased.");
            var signatures = body.Elements(ns + "Signatures")
                .Select(node => material.ResolvePath("OFD.xml", ReadValue(node, material, token) ?? string.Empty, token))
                .ToArray();
            var title = ReadValue(body.Element(ns + "DocInfo")?.Element(ns + "Title"), material, token);
            result.Add((body, new OfdDocumentDescriptor(result.Count, path, title, signatures)));
        }
        return result;
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
        material.ReserveReadModel(1024, cancellationToken);
        var result = new AnnotationIndex { Namespace = document.Root!.Name.Namespace };
        var visited = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var declaration in document.Root!.Elements().Where(node => node.Name.LocalName == "Annotations"))
        {
            cancellationToken.ThrowIfCancellationRequested();
            result.Source.Declarations.Add(OfdXmlEnvelope.Capture(declaration, material, cancellationToken));
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
            result.Source.Capture(listPath, archive.TryGetBytes(listPath, out var indexBytes) ? indexBytes : null, material, cancellationToken);
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
        OfdModelMaterialization material, CancellationToken cancellationToken, IReadOnlyDictionary<long, int> idCounts)
    {
        material.ReserveReadModel(512, cancellationToken);
        var typed = new OfdAnnotationCollection();
        if (index.UnmodeledXml is not null)
            page.AnnotationAppearances.Add(AnnotationIndexMetadata(index.UnmodeledXml));
        if (page.Id is null || !index.Files.TryGetValue(page.Id, out var files))
        { page.AnnotationOwnership.Bind(page, null, material, cancellationToken); return; }
        if (files.UnmodeledXml is not null)
            page.AnnotationAppearances.Add(AnnotationIndexMetadata(files.UnmodeledXml));
        foreach (var annotationPath in files.Paths)
        {
            cancellationToken.ThrowIfCancellationRequested();
            material.ReserveReadModel(1024L + 16L * annotationPath.Length, cancellationToken);
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
                    Xml = ReadXmlString(new XElement("UnmodeledAnnotation", new XAttribute("FileLoc", annotationPath)), material, cancellationToken)! });
                continue;
            }
            // Charge each actual page/file occurrence before its first persistent
            // root/source XML copy. Cached DOMs do not cache caller quota usage.
            foreach (var sourceAnnotation in annotations.Root!.Elements())
            foreach (var appearanceNode in sourceAnnotation.Elements(sourceAnnotation.Name.Namespace + "Appearance"))
            foreach (var graphic in appearanceNode.Descendants())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (graphic.Name.Namespace == appearanceNode.Name.Namespace && graphic.Name.LocalName is "TextObject" or "ImageObject" or "PathObject") budget.ChargeObject();
            }
            material.ReadAnnotationFileAdmissionCheckpoint();
            page.AnnotationOwnership.Entries.Add(annotationPath, OfdXmlEnvelope.Capture(annotations.Root!, material, cancellationToken));
            if (annotations.Root!.Attributes().Any(a => !a.IsNamespaceDeclaration) || annotations.Root.Nodes().Any(n => n is XProcessingInstruction or XComment || n is XText t && !string.IsNullOrWhiteSpace(t.Value)) || annotations.Nodes().Any(n => n is XProcessingInstruction or XComment))
                page.AnnotationOwnership.UnsafeDeletionEntries.Add(annotationPath);
            var shared = index.Files.Values.Count(candidate => candidate.Known.Contains(annotationPath)) > 1;
            var annotationSlot = -1;
            if (!PlainContainer(annotations.Root!) || annotations.Nodes().Any(node => node is not XElement and not XComment && (node is not XText text || !string.IsNullOrWhiteSpace(text.Value))))
                page.AnnotationAppearances.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(annotations, material, cancellationToken)));
            foreach (var annotation in annotations.Root.Elements())
            {
                cancellationToken.ThrowIfCancellationRequested();
                material.ReserveReadModel(1024, cancellationToken);
                annotationSlot++;
                var projectionStart = page.AnnotationAppearances.Count;
                var admitted = shared || annotation.Name != annotations.Root.Name.Namespace + "Annot" ? null : ReadTypedAnnotation(archive, annotation, fonts, media, mediaTypes,
                    annotationPath, idCounts, material, cancellationToken);
                try
                {
                if (annotation.Name != annotations.Root.Name.Namespace + "Annot")
                {
                    page.AnnotationAppearances.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken)));
                    continue;
                }
                if (!SupportedAnnotationWrapper(annotation))
                    page.AnnotationAppearances.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken)));
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
                            page.AnnotationAppearances.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken)));
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
                        page.AnnotationAppearances.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(annotation, material, cancellationToken)));
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
                if (hasUnmodeledMetadata) staged.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken)));
                if (nodes.SelectMany(node => node.DescendantsAndSelf().Attributes()).Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute)))
                    staged.Add(AnnotationMetadata(OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken)));
                foreach (var element in ParsePageObjects(archive, xml, fonts, media, mediaTypes, annotationPath, material, cancellationToken))
                {
                    if (element is OfdImageElement { Data.Length: 0 })
                    {
                        staged.Clear();
                        staged.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
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
                        staged.Add(new OfdRawElement { LocalName = "UnsupportedAnnotationAppearance", Xml = OfdXmlEnvelope.SerializeForRead(appearance, material, cancellationToken) });
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
        material.ReserveReadModel(1024, token);
        var flags = new Dictionary<string, bool?>();
        foreach (var name in new[] { "Visible", "Print", "NoZoom", "NoRotate", "ReadOnly" })
        {
            var literal = annotation.Attribute(name)?.Value;
            if (literal is not null && literal is not ("true" or "false" or "1" or "0")) return null;
            flags[name] = literal is null ? null : literal is "true" or "1";
        }
        var appearance = annotation.Element(ns + "Appearance")!;
        if (!TryParseBox(appearance.Attribute("Boundary")?.Value, material, token, out _) ||
            appearance.Nodes().Any(n => n is XText t && !string.IsNullOrWhiteSpace(t.Value)) ||
            appearance.Elements().Any(n => n.Name.Namespace != ns || n.Name.LocalName is not ("TextObject" or "ImageObject" or "PathObject") ||
                !Unique(n.Attribute("ID")?.Value) || !IsSupportedAnnotationPrimitive(n, ns))) return null;
        foreach (var primitive in appearance.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (!TryParseBox(primitive.Attribute("Boundary")?.Value, material, token, out var primitiveBox) ||
                !OfdBoxParser.IsWritablePositiveSide(primitiveBox.w) || !OfdBoxParser.IsWritablePositiveSide(primitiveBox.h)) return null;
            var primitiveMatrix = ParseMatrix(primitive.Attribute("CTM")?.Value, material, token);
            if (primitive.Attribute("CTM") is not null && (primitiveMatrix is null || primitiveMatrix.Any(v => double.IsNaN(v) || double.IsInfinity(v)))) return null;
            if (primitive.Name.LocalName == "TextObject" && (!double.TryParse(primitive.Attribute("Size")?.Value, NumberStyles.Float, CultureInfo.InvariantCulture, out var size) || !(size > 0) || double.IsInfinity(size))) return null;
        }
        var matrix = ParseMatrix(appearance.Attribute("CTM")?.Value, material, token);
        if (appearance.Attribute("CTM") is not null && (matrix is null || matrix.Any(v => double.IsNaN(v) || double.IsInfinity(v)))) return null;
        var box = ParseBox(appearance.Attribute("Boundary")?.Value, material, token);
        if (box.w <= 0 || box.h <= 0) return null;
        var remark = annotation.Element(ns + "Remark");
        if (remark is not null && (remark.HasElements || remark.Nodes().Any(n => n is not XText and not XComment))) return null;
        material.ReserveReadModel(2048, token);
        var result = new OfdAnnotation
        {
            Id = annotation.Attribute("ID")!.Value, Type = type, Creator = annotation.Attribute("Creator")!.Value,
            LastModifiedDate = DateTime.SpecifyKind(date, DateTimeKind.Unspecified), Subtype = annotation.Attribute("Subtype")?.Value,
            Visible = flags["Visible"], Print = flags["Print"], NoZoom = flags["NoZoom"], NoRotate = flags["NoRotate"], ReadOnly = flags["ReadOnly"],
            Remark = ReadValue(remark, material, token), SourceXml = OfdXmlEnvelope.Capture(annotation, material, token),
            Appearance = new OfdAnnotationAppearance { XMillimeters = box.x, YMillimeters = box.y, WidthMillimeters = box.w,
                HeightMillimeters = box.h, Transform = matrix, SourceXml = OfdXmlEnvelope.Capture(appearance, material, token) }
        };
        // Capture inherits namespace bindings before attaching nodes to the parser's narrow synthetic container.
        var primitiveCount = appearance.Elements().Count();
        material.ReserveReadModel(512L + primitiveCount * 32L, token);
        var nodes = new List<XElement>(primitiveCount);
        foreach (var primitive in appearance.Elements())
        {
            var source = OfdXmlEnvelope.Capture(primitive, material, token);
            material.PrepareReadXmlParser(checked(2L * source.Length), token);
            material.ReserveReadModel(32768L + 128L * source.Length, token);
            nodes.Add(OfdXmlEnvelope.Parse(source));
        }
        material.ReserveReadModel(2048, token);
        var xml = new XDocument(new XElement(ns + "Page", new XElement(ns + "Content", new XElement(ns + "Layer", nodes))));
        foreach (var element in ParsePageObjects(archive, xml, fonts, media, mediaTypes, path, material, token))
        {
            if (element is OfdRawElement || element is OfdImageElement { Data.Length: 0 }) return null;
            result.Appearance.Elements.Add(element);
        }
        return result;
        bool Unique(string? value) => OfdSourceIdentity.ParseId(value) is var id && id != 0 && idCounts.TryGetValue(id, out var count) && count == 1;
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

    private static void ReadFonts(
        OfdPackageArchive archive, string resourcePath, OfdDocumentPackage package,
        IDictionary<string, string> fontMap, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (!archive.Contains(resourcePath)) return;
        var xml = material.ReadXml(resourcePath, archive.GetBytes(resourcePath), cancellationToken);
        var root = xml.Root;
        if (root is null || root.Name.LocalName != "Res" ||
            root.Name.Namespace != XNamespace.None && root.Name.NamespaceName is not (OfdConstants.Namespace or OfdConstants.StandardNamespace))
            return;
        var ns = root.Name.Namespace;
        var baseLocation = root.Attribute("BaseLoc")?.Value;
        var containers = root.Elements(ns + "Fonts").ToArray();
        var validContainer = OfdResourceXmlContract.IsPlainRoot(xml) && containers.Length == 1 && PlainContainer(containers[0]) &&
            containers[0].Elements().All(child => child.Name == ns + "Font");
        foreach (var font in containers.SelectMany(fonts => fonts.Elements(ns + "Font")))
        {
            var id = font.Attribute("ID")?.Value;
            var name = font.Attribute("FontName")?.Value ?? font.Attribute("FamilyName")?.Value;
            if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(name)) continue;
            var key = CanonicalResourceId(id);
            if (key is null) continue;
            var validBold = TryParseFontFlag(font.Attribute("Bold")?.Value, out var bold);
            var validItalic = TryParseFontFlag(font.Attribute("Italic")?.Value, out var italic);
            var files = font.Elements(ns + "FontFile").ToArray();
            var validDeclaration = validContainer &&
                PlainContainer(font, "ID", "FontName", "FamilyName", "Charset", "Bold", "Italic", "Serif", "FixedWidth") &&
                font.Elements().All(child => child.Name == ns + "FontFile") && files.Length <= 1 &&
                files.All(child => child.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
                    child.Nodes().All(node => node is XText or XComment));
            var fontFile = files.FirstOrDefault();
            var file = ReadValue(fontFile, material, cancellationToken);
            material.ReserveReadModel(512, cancellationToken);
            var resource = new OfdFontResource
            {
                Id = id!, FontName = name!, FamilyName = font.Attribute("FamilyName")?.Value,
                Charset = font.Attribute("Charset")?.Value,
                Bold = bold, Italic = italic, FileName = file
            };
            var resolvedFile = fontFile is null;
            if (fontFile is not null && !string.IsNullOrWhiteSpace(file))
            {
                material.ReserveReadModel(256L + 16L * ((baseLocation?.Length ?? 0) + file!.Length), cancellationToken);
                var reference = string.IsNullOrWhiteSpace(baseLocation) || file!.StartsWith("/", StringComparison.Ordinal)
                    ? file! : baseLocation!.TrimEnd('/') + "/" + file;
                if (archive.TryGetBytes(material.ResolvePath(resourcePath, reference, cancellationToken), out var bytes) && bytes.Length > 0)
                { resource.Data = bytes; resolvedFile = true; }
            }
            // A name-only font deliberately permits host fallback. An explicit
            // FontFile must resolve before it can satisfy PageBlock admission.
            if (validDeclaration && validBold && validItalic && resolvedFile) fontMap[key] = name!;
            else fontMap.Remove(key);
            if (package.Fonts.Any(existing => existing.Id == id)) continue;
            package.Fonts.Add(resource);
        }
    }

    private static void ReadMediaResources(
        OfdPackageArchive archive,
        string resPath,
        IDictionary<string, string> mediaMap,
        IDictionary<string, string> mediaTypeMap, ISet<string> ambiguousMedia,
        OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (!archive.Contains(resPath))
        {
            return;
        }

        var resXml = material.ReadXml(resPath, archive.GetBytes(resPath), cancellationToken);
        var root = resXml.Root;
        if (root is null || root.Name.LocalName != "Res" ||
            root.Name.Namespace != XNamespace.None && root.Name.NamespaceName is not (OfdConstants.Namespace or OfdConstants.StandardNamespace))
            return;
        var resNs = root.Name.Namespace;
        var baseLoc = root.Attribute("BaseLoc")?.Value;
        var containers = root.Elements(resNs + "MultiMedias").ToArray();
        var validContainer = OfdResourceXmlContract.IsPlainRoot(resXml) && containers.Length == 1 && PlainContainer(containers[0]) &&
            containers[0].Elements().All(child => child.Name == resNs + "MultiMedia");
        foreach (var media in containers.SelectMany(group => group.Elements(resNs + "MultiMedia")))
        {
            var id = media.Attribute("ID")?.Value;
            var key = CanonicalResourceId(id);
            if (key is null) continue;
            var fileAttribute = media.Attribute("MediaFile");
            var files = media.Elements(resNs + "MediaFile").ToArray();
            var validDeclaration = validContainer &&
                PlainContainer(media, "ID", "Type", "Format", "MediaFile") &&
                media.Elements().All(child => child.Name == resNs + "MediaFile") &&
                files.Length + (fileAttribute is null ? 0 : 1) == 1 &&
                files.All(child => child.Attributes().All(attribute => attribute.IsNamespaceDeclaration) &&
                    child.Nodes().All(node => node is XText or XComment));
            var declaredFormat = media.Attribute("Format")?.Value;
            var format = declaredFormat?.Trim().ToUpperInvariant();
            var validFormat = format is null or "PNG" or "JPG" or "JPEG" or "BMP" or "TIFF";
            var file = fileAttribute?.Value ?? ReadValue(files.FirstOrDefault(), material, cancellationToken);
            if (string.IsNullOrWhiteSpace(file))
            {
                ambiguousMedia.Add(key); continue;
            }

            var declaredType = media.Attribute("Type")?.Value;
            if (declaredType is not null && !string.Equals(declaredType, "Image", StringComparison.Ordinal))
            {
                // A page-level resource can shadow a document-level image ID.
                // Keep an explicit non-image declaration unresolved for ImageObject.
                ambiguousMedia.Add(key);
                mediaMap.Remove(key);
                mediaTypeMap.Remove(key);
                continue;
            }

            material.ReserveReadModel(512L + 16L * ((baseLoc?.Length ?? 0) + file!.Length + (media.Attribute("Format")?.Value.Length ?? 0)), cancellationToken);
            var mediaLoc = string.IsNullOrWhiteSpace(baseLoc) || file!.StartsWith("/", StringComparison.Ordinal)
                ? file!
                : $"{baseLoc!.TrimEnd('/')}/{file.TrimStart('/')}";

            var mediaPath = material.ResolvePath(resPath, mediaLoc, cancellationToken);
            // Preserve permissive legacy top-level imports. PageBlock admission
            // separately requires an explicit format to match the actual bytes.
            var validPayload = format is null ||
                archive.TryGetBytes(mediaPath, out var payload) && MatchesMediaFormat(payload, format, cancellationToken);
            if (validDeclaration && validFormat && validPayload) ambiguousMedia.Remove(key);
            else ambiguousMedia.Add(key);
            mediaMap[key] = mediaPath;
            mediaTypeMap[key] = ToMediaType(media.Attribute("Format")?.Value);
        }
    }

    private static bool MatchesMediaFormat(byte[] payload, string format, CancellationToken token)
    {
        var mime = format switch
        {
            "PNG" => "image/png",
            "JPG" or "JPEG" => "image/jpeg",
            _ => null
        };
        if (mime is null) return false;
        try { EncodedImageAdmission.Validate(payload, mime, token); return true; }
        catch (Exception error) when (error is InvalidDataException or IOException or NotSupportedException or OverflowException or IndexOutOfRangeException)
        { return false; }
    }

    private static Dictionary<string, OfdCompositeResource> ReadCompositeResources(OfdPackageArchive archive, OfdDocumentPackage package,
        string documentPath, XNamespace ns, IReadOnlyDictionary<string, string> fonts, IReadOnlyDictionary<string, string> media,
        IReadOnlyDictionary<string, string> mediaTypes, ISet<string> ambiguousMedia,
        OfdPaintResources paint, OfdModelMaterialization material, CancellationToken token)
    {
        var paths = new[] { package.PublicResourceLocation, package.DocumentResourceLocation }
            .Where(location => !string.IsNullOrWhiteSpace(location)).Distinct()
            .Select(location => material.ResolvePath(documentPath, location!, token));
        var (map, definitions) = ReadCompositeResourceScope(archive, package, paths, ns, null,
            fonts, media, mediaTypes, ambiguousMedia, paint, material, token);
        if (definitions.Count == 0) return map;
        package.CompositeResources = new();
        foreach (var resource in definitions) package.CompositeResources.Items.Add(resource);
        material.ReserveReadModel(32L * package.CompositeResources.Items.Count, token);
        package.OriginalCompositeOrder = package.CompositeResources.Items.ToArray();
        package.OriginalCompositeResources = package.CompositeResources;
        package.CompositeBaseline = OfdModelSnapshot.Capture(package.CompositeResources, material, token);
        OfdCompositeGraph.Validate(package, material, token); return map;
    }

    private static IReadOnlyDictionary<string, OfdCompositeResource> ReadScopedCompositeResources(
        OfdPackageArchive archive, OfdDocumentPackage package, string path, XNamespace ns,
        IReadOnlyDictionary<string, OfdCompositeResource> inherited,
        IReadOnlyDictionary<string, string> fonts, IReadOnlyDictionary<string, string> media,
        IReadOnlyDictionary<string, string> mediaTypes, ISet<string> ambiguousMedia,
        OfdPaintResources paint, OfdModelMaterialization material, CancellationToken token)
    {
        if (!archive.Contains(path)) return inherited;
        var (map, definitions) = ReadCompositeResourceScope(archive, package, new[] { path }, ns, inherited,
            fonts, media, mediaTypes, ambiguousMedia, paint, material, token);
        if (definitions.Count == 0) return map;
        // Page and template definitions are read-only PageBlock dependencies.
        // Validate their graph without granting Writer ownership of PageRes.
        var validation = new OfdDocumentPackage { CompositeResources = new OfdCompositeResourceCollection() };
        var reachable = new HashSet<OfdCompositeResource>();
        var pending = new Stack<OfdCompositeResource>(definitions);
        while (pending.Count > 0)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            var resource = pending.Pop();
            if (!reachable.Add(resource)) continue;
            validation.CompositeResources.Items.Add(resource);
            foreach (var placement in resource.Elements.OfType<OfdCompositeElement>()) pending.Push(placement.Resource);
        }
        OfdCompositeGraph.Validate(validation, material, token);
        return map;
    }

    private static (Dictionary<string, OfdCompositeResource> Map, List<OfdCompositeResource> Definitions) ReadCompositeResourceScope(
        OfdPackageArchive archive, OfdDocumentPackage package, IEnumerable<string> paths, XNamespace documentNs,
        IReadOnlyDictionary<string, OfdCompositeResource>? inherited,
        IReadOnlyDictionary<string, string> fonts, IReadOnlyDictionary<string, string> media,
        IReadOnlyDictionary<string, string> mediaTypes, ISet<string> ambiguousMedia,
        OfdPaintResources paint, OfdModelMaterialization material, CancellationToken token)
    {
        material.ReserveReadModel(512, token);
        var map = new Dictionary<string, OfdCompositeResource>(StringComparer.Ordinal);
        if (inherited is not null) foreach (var pair in inherited) map.Add(pair.Key, pair.Value);
        var localIds = new HashSet<string>(StringComparer.Ordinal);
        var definitions = new List<(OfdCompositeResource Resource, XElement Node, string Path)>();
        foreach (var path in paths)
        {
            if (!archive.Contains(path)) continue;
            var xml = material.ReadXml(path, archive.GetBytes(path), token); var root = xml.Root!;
            if (root.Name.LocalName != "Res" ||
                (inherited is null ? root.Name.Namespace != documentNs : !SupportedGraphicNamespace(root.Name.Namespace))) continue;
            var ns = root.Name.Namespace;
            var containers = root.Elements(ns + "CompositeGraphicUnits").ToArray();
            var supportedContainers = OfdResourceXmlContract.IsPlainRoot(xml) && containers.Length == 1 && PlainContainer(containers[0]) &&
                containers[0].Elements().All(child => child.Name == ns + "CompositeGraphicUnit");
            foreach (var container in containers)
            foreach (var pair in container.Elements().Select((node, slot) => (node, slot)))
            {
                token.ThrowIfCancellationRequested(); material.ReserveReadModel(2048, token);
                var node = pair.node; if (node.Name != ns + "CompositeGraphicUnit") continue;
                var id = node.Attribute("ID")?.Value;
                var key = CanonicalResourceId(id);
                if (key is null || !localIds.Add(key)) throw new InvalidDataException("Composite resource ID is missing or duplicated.");
                var resource = new OfdCompositeResource { Id = id, WidthMillimeters = ParseDouble(node.Attribute("Width")?.Value, 0),
                    HeightMillimeters = ParseDouble(node.Attribute("Height")?.Value, 0), SourceEntryPath = path, SourceSlot = pair.slot,
                    SourceXml = OfdXmlEnvelope.Capture(node, material, token), SourceContainerSupported = supportedContainers };
                map[key] = resource; definitions.Add((resource, node, path));
            }
        }
        foreach (var definition in definitions)
        {
            material.ReserveReadModel(2048L + definition.Node.Elements().Count() * 128L, token);
            foreach (var element in ParseGraphicNodes(archive, definition.Node.Elements(), null, "Body", fonts, media, mediaTypes,
                definition.Path, material, token, map, definition.Node.Name.Namespace, paint, ambiguousMedia: ambiguousMedia))
            {
                if (element is OfdTextElement unresolvedText &&
                    !string.IsNullOrWhiteSpace(unresolvedText.FontResourceId) &&
                    !HasFont(fonts, unresolvedText.FontResourceId) &&
                    package.Fonts.Any(font => OfdSourceIdentity.ParseId(font.Id) == OfdSourceIdentity.ParseId(unresolvedText.FontResourceId)))
                {
                    // The font exists but its declaration was rejected for
                    // rendering. Preserve TextCode for logical extraction.
                    unresolvedText.RenderPaintError ??= "Composite text has an unresolved Font resource.";
                    definition.Resource.Elements.Add(unresolvedText);
                    continue;
                }
                if (element is OfdTextElement text && !HasFont(fonts, text.FontResourceId) ||
                    element is OfdImageElement { Data.Length: 0 })
                {
                    var source = element is OfdTextElement t ? t.SourceXml! : ((OfdImageElement)element).SourceXml!;
                    definition.Resource.Elements.Add(new OfdRawElement { LocalName = OfdXmlEnvelope.Parse(source).Name.LocalName, Xml = source,
                        ObjectId = element.ObjectId, XMillimeters = element.XMillimeters, YMillimeters = element.YMillimeters,
                        WidthMillimeters = element.WidthMillimeters, HeightMillimeters = element.HeightMillimeters });
                }
                else definition.Resource.Elements.Add(element);
            }
        }
        return (map, definitions.Select(item => item.Resource).ToList());
    }

    private static bool SupportedGraphicNamespace(XNamespace ns) => ns == XNamespace.None ||
        ns.NamespaceName is OfdConstants.Namespace or OfdConstants.StandardNamespace;

    private static IEnumerable<OfdElement> ParsePageObjects(
        OfdPackageArchive archive,
        XDocument pageXml,
        IReadOnlyDictionary<string, string> fontMap,
        IReadOnlyDictionary<string, string> mediaMap,
        IReadOnlyDictionary<string, string> mediaTypeMap,
        string pageResPath, OfdModelMaterialization material, CancellationToken cancellationToken,
        IReadOnlyDictionary<string, OfdCompositeResource>? compositeMap = null, OfdPaintResources? paint = null,
        ISet<string>? ambiguousMedia = null, IReadOnlyDictionary<string, OfdCompositeResource>? pageBlockCompositeMap = null)
    {
        material.ReserveReadModel(256L + (pageXml.Root?.Descendants().Count() ?? 0) * 32L, cancellationToken);
        var root = pageXml.Root;
        if (root is not null && (!SupportedGraphicNamespace(root.Name.Namespace) || root.Name.LocalName != "Page" ||
            OfdPageXmlContract.ContainsPageBlock(root) && !OfdPageXmlContract.IsPlainPageEnvelope(pageXml)))
        {
            yield return new OfdRawElement { LocalName = root.Name.LocalName,
                Xml = OfdXmlEnvelope.Capture(root, material, cancellationToken) };
            yield break;
        }
        var content = root?.Elements().FirstOrDefault(x => x.Name.LocalName == "Content");
        if (content is not null && (!SupportedGraphicNamespace(content.Name.Namespace) ||
            OfdPageXmlContract.ContainsPageBlock(content) && !OfdPageXmlContract.IsPlainContentEnvelope(content)))
        {
            yield return new OfdRawElement { LocalName = "Content",
                Xml = OfdXmlEnvelope.Capture(content, material, cancellationToken) };
            yield break;
        }
        var layers = content?
            .Elements()
            .Where(x => x.Name.LocalName == "Layer")
            .ToList() ?? [];

        foreach (var layer in layers)
        {
            if (!SupportedGraphicNamespace(layer.Name.Namespace) ||
                OfdPageXmlContract.HasPageBlock(layer) && !OfdPageXmlContract.IsPlainLayerEnvelope(layer))
            {
                yield return new OfdRawElement
                {
                    ObjectId = layer.Attribute("ID")?.Value,
                    LayerId = layer.Attribute("ID")?.Value,
                    LayerType = layer.Attribute("Type")?.Value ?? "Body",
                    LocalName = "Layer",
                    Xml = OfdXmlEnvelope.Capture(layer, material, cancellationToken)
                };
                continue;
            }
            foreach (var element in ParseGraphicNodes(archive, layer.Elements(), layer.Attribute("ID")?.Value, layer.Attribute("Type")?.Value ?? "Body",
                fontMap, mediaMap, mediaTypeMap, pageResPath, material, cancellationToken, compositeMap, layer.Name.Namespace, paint,
                ambiguousMedia: ambiguousMedia, pageBlockCompositeMap: pageBlockCompositeMap)) yield return element;
        }
    }

    private static IEnumerable<OfdElement> ParseGraphicNodes(OfdPackageArchive archive, IEnumerable<XElement> nodes, string? layerId, string layerType,
        IReadOnlyDictionary<string, string> fontMap, IReadOnlyDictionary<string, string> mediaMap, IReadOnlyDictionary<string, string> mediaTypeMap,
        string pageResPath, OfdModelMaterialization material, CancellationToken cancellationToken, IReadOnlyDictionary<string, OfdCompositeResource>? compositeMap, XNamespace structuralNamespace, OfdPaintResources? paint = null, int blockDepth = 0,
        ISet<string>? ambiguousMedia = null, IReadOnlyDictionary<string, OfdCompositeResource>? pageBlockCompositeMap = null)
    {
            foreach (var node in nodes)
            {
                material.ReserveReadModel(512, cancellationToken);
                var localName = node.Name.LocalName;
                // Existing vector imports can mix the two OFD namespaces or use
                // unqualified graphic names. Extension namespaces remain opaque.
                var knownGraphicNamespace = SupportedGraphicNamespace(node.Name.Namespace);
                if (knownGraphicNamespace && SupportedGraphicNamespace(structuralNamespace) && localName == "PageBlock")
                {
                    material.Structure.Depth(blockDepth + 1);
                    var xml = OfdXmlEnvelope.Capture(node, material, cancellationToken);
                    var block = new OfdRawElement
                    {
                        ObjectId = node.Attribute("ID")?.Value, LayerId = layerId, LayerType = layerType,
                        LocalName = "PageBlock", Xml = xml, PageBlockSourceXml = xml,
                        PageBlockChildren = new List<OfdElement>()
                    };
                    foreach (var child in ParseGraphicNodes(archive, node.Elements(), layerId, layerType,
                        fontMap, mediaMap, mediaTypeMap, pageResPath, material, cancellationToken,
                        pageBlockCompositeMap ?? compositeMap, node.Name.Namespace, paint, blockDepth + 1,
                        ambiguousMedia, pageBlockCompositeMap))
                        block.PageBlockChildren.Add(child);
                    yield return block;
                    continue;
                }
                if (knownGraphicNamespace && string.Equals(localName, "TextObject", StringComparison.OrdinalIgnoreCase))
                {
                    var sourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken);
                    var (renderNode, paintError) = paint?.Normalize(node, sourceXml.Length, material, cancellationToken) ?? (node, (string?)null);
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
                        SourceInheritedTextCodeAxes = blockDepth > 0,
                        FontResourceId = node.Attribute("Font")?.Value,
                        FontName = ResolveFontName(node.Attribute("Font")?.Value, fontMap),
                        FontSizeMillimeters = ParseDouble(node.Attribute("Size")?.Value, 4d),
                        Weight = (int)ParseDouble(node.Attribute("Weight")?.Value, OfdTextElement.DefaultWeight),
                        Italic = ParseBoolean(node.Attribute("Italic")?.Value, false),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value, material, cancellationToken),
                        FillColor = ParseColor(renderNode.Elements()
                            .FirstOrDefault(x => x.Name.LocalName == "FillColor"), material, cancellationToken) ?? OfdColor.Black,
                        SourceXml = sourceXml,
                        ClippingXml = ReadXmlString(node.Element(node.Name.Namespace + "Clips"), material, cancellationToken)
                    };
                    AttachPaint(text, node, renderNode, paintError, material, cancellationToken);
                    if (blockDepth > 0 && textCodes.Count == 0)
                        text.RenderPaintError ??= "PageBlock text requires at least one TextCode.";
                    if (blockDepth > 0 && !HasFont(fontMap, text.FontResourceId))
                        text.RenderPaintError ??= "PageBlock text has an unresolved Font resource.";
                    double? previousX = null, previousY = null;
                    foreach (var textCode in textCodes)
                    {
                        material.ReserveReadModel(256, cancellationToken);
                        if (blockDepth > 0 &&
                            (previousX is null && textCode.Attribute("X") is null ||
                             previousY is null && textCode.Attribute("Y") is null))
                            text.RenderPaintError ??= "The first PageBlock TextCode requires X and Y.";
                        var x = ParseDouble(textCode.Attribute("X")?.Value, blockDepth > 0 ? previousX ?? 0d : 0d);
                        var y = ParseDouble(textCode.Attribute("Y")?.Value,
                            blockDepth > 0 ? previousY ?? text.FontSizeMillimeters : text.FontSizeMillimeters);
                        text.Runs.Add(new OfdTextRun
                        {
                            Text = ReadDirectText(textCode, material, cancellationToken),
                            XMillimeters = x,
                            YMillimeters = y,
                            DeltaX = textCode.Attribute("DeltaX")?.Value,
                            DeltaY = textCode.Attribute("DeltaY")?.Value
                        });
                        previousX = x;
                        previousY = y;
                    }

                    OfdClipXml.Load(text, material, cancellationToken);
                    yield return text;
                    continue;
                }

                if (knownGraphicNamespace && string.Equals(localName, "ImageObject", StringComparison.OrdinalIgnoreCase))
                {
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
                    var resourceId = node.Attribute("ResourceID")?.Value ?? string.Empty;
                    var mediaKey = CanonicalResourceId(resourceId);
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
                        MediaType = mediaKey is not null && mediaTypeMap.TryGetValue(mediaKey, out var mediaType) ? mediaType : "image/png",
                        SourceMediaAmbiguous = mediaKey is not null && ambiguousMedia?.Contains(mediaKey) == true,
                        SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken)
                    };

                    if (mediaKey is not null && (blockDepth == 0 || ambiguousMedia?.Contains(mediaKey) != true) &&
                        mediaMap.TryGetValue(mediaKey, out var mediaPath))
                    {
                        if (archive.TryGetBytes(mediaPath, out var bytes))
                        {
                            image.Data = bytes;
                            image.FileName = Path.GetFileName(mediaPath);
                        }
                    }

                    OfdClipXml.Load(image, material, cancellationToken);
                    yield return image;
                    continue;
                }

                if (knownGraphicNamespace && string.Equals(localName, "PathObject", StringComparison.OrdinalIgnoreCase))
                {
                    var sourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken);
                    var (renderNode, paintError) = paint?.Normalize(node, sourceXml.Length, material, cancellationToken) ?? (node, (string?)null);
                    var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
                    var strokeColor = ParseColor(renderNode.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "StrokeColor"), material, cancellationToken);
                    var fillColor = ParseColor(renderNode.Elements()
                        .FirstOrDefault(x => x.Name.LocalName == "FillColor"), material, cancellationToken);
                    var path = new OfdPathElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        // Match XElement.Value used by path admission, including whitespace between XML comments.
                        AbbreviatedData = ReadDirectText(node.Element(node.Name.Namespace + "AbbreviatedData"), material, cancellationToken,
                            preserveWhitespace: true),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value, material, cancellationToken),
                        FillRule = node.Attribute("Rule")?.Value switch { "NonZero" => OfdFillRule.NonZero, "Even-Odd" => OfdFillRule.EvenOdd, _ => null },
                        LineWidthMillimeters = ParseDouble(renderNode.Attribute("LineWidth")?.Value, 0.353d),
                        Stroke = ParseBoolean(node.Attribute("Stroke")?.Value, true),
                        Fill = ParseBoolean(node.Attribute("Fill")?.Value, false),
                        StrokeColor = strokeColor ?? OfdColor.Black,
                        FillColor = fillColor,
                        SourceXml = sourceXml,
                        ClippingXml = ReadXmlString(node.Element(node.Name.Namespace + "Clips"), material, cancellationToken)
                    };
                    AttachPaint(path, node, renderNode, paintError, material, cancellationToken);
                    OfdClipXml.Load(path, material, cancellationToken);
                    yield return path;
                    continue;
                }

                // PageBlock admits mixed supported drawing namespaces; ordinary
                // composite placements retain their enclosing namespace contract.
                if (knownGraphicNamespace && SupportedGraphicNamespace(structuralNamespace) && localName == "CompositeObject" &&
                    (blockDepth > 0 || node.Name == structuralNamespace + "CompositeObject"))
                {
                    var resourceId = node.Attribute("ResourceID")?.Value;
                    var compositeKey = CanonicalResourceId(resourceId);
                    if (compositeKey is not null && compositeMap?.TryGetValue(compositeKey, out var resource) == true)
                    {
                        var boundary = ParseBox(node.Attribute("Boundary")?.Value, material, cancellationToken);
                        var placement = new OfdCompositeElement { ObjectId = node.Attribute("ID")?.Value, LayerId = layerId, LayerType = layerType,
                            XMillimeters = boundary.x, YMillimeters = boundary.y, WidthMillimeters = boundary.w, HeightMillimeters = boundary.h,
                            Resource = resource, Transform = ParseMatrix(node.Attribute("CTM")?.Value, material, cancellationToken),
                            SourceXml = OfdXmlEnvelope.Capture(node, material, cancellationToken),
                            ClippingXml = ReadXmlString(node.Element(node.Name.Namespace + "Clips"), material, cancellationToken) };
                        OfdClipXml.Load(placement, material, cancellationToken); yield return placement; continue;
                    }
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

    private static void AttachPaint(OfdElement element, XElement source, XElement render, string? error,
        OfdModelMaterialization material, CancellationToken token)
    {
        element.RenderPaintError = error;
        if (ReferenceEquals(source, render)) return;
        element.RenderSourceAuthority = element switch
        {
            OfdTextElement text => text.SourceXml,
            OfdPathElement path => path.SourceXml,
            _ => null
        };
        element.RenderSourceXml = OfdXmlEnvelope.Capture(render, material, token);
    }

    private static string ReadDirectText(XElement? node, OfdModelMaterialization material, CancellationToken token,
        bool preserveWhitespace = false)
    {
        if (node is null) return string.Empty;
        material.ReserveReadModel(512, token);
        var values = node.Nodes().OfType<XText>().Where(text => preserveWhitespace || text is XCData || !string.IsNullOrWhiteSpace(text.Value) ||
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


    private static bool ReadKeyValueCustomTags(
        OfdPackageArchive archive,
        string detailPath,
        IDictionary<string, string> destination, XNamespace ns, OfdModelMaterialization material, CancellationToken cancellationToken)
    {
        if (!archive.Contains(detailPath))
        {
            return false;
        }

        var tagsXml = material.ReadXml(detailPath, archive.GetBytes(detailPath), cancellationToken);
        if (tagsXml.Root?.Name != ns + "EMRTags") return false;
        foreach (var tag in tagsXml.Root.Elements(ns + "Tag"))
        {
            var key = tag.Attribute("Key")?.Value;
            var value = tag.Attribute("Value")?.Value;
            if (!string.IsNullOrWhiteSpace(key) && value is not null)
            {
                material.ReserveReadModel(160, cancellationToken);
                destination[key!] = value;
            }
        }
        return true;
    }

    private static string? CanonicalResourceId(string? id)
    {
        var numericId = OfdSourceIdentity.ParseId(id);
        return numericId == 0 ? null : numericId.ToString(CultureInfo.InvariantCulture);
    }

    private static bool HasFont(IReadOnlyDictionary<string, string> fontMap, string? fontRef)
        => CanonicalResourceId(fontRef) is { } key && fontMap.ContainsKey(key);

    private static string ResolveFontName(string? fontRef, IReadOnlyDictionary<string, string> fontMap)
    {
        if (CanonicalResourceId(fontRef) is { } key && fontMap.TryGetValue(key, out var fontName))
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

        var normalized = value!.Trim();
        return normalized switch
        {
            "1" => true,
            "0" => false,
            _ => bool.TryParse(normalized, out var parsed) ? parsed : fallback
        };
    }

    private static bool TryParseFontFlag(string? value, out bool result)
    {
        result = false;
        if (value is null) return true;
        var literal = value.Trim();
        if (literal == "1") { result = true; return true; }
        if (literal == "0") return true;
        return bool.TryParse(literal, out result);
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
        if (colorElement is null) return null;
        var value = colorElement.Attribute("Value")?.Value;
        if (value is null)
        {
            var defaultAlpha = 255;
            var literal = colorElement.Attribute("Alpha")?.Value;
            if (literal is not null && (!int.TryParse(literal, NumberStyles.Integer, CultureInfo.InvariantCulture, out defaultAlpha) ||
                defaultAlpha is < 0 or > 255)) return null;
            material.ReserveReadModel(512, token);
            return new OfdColor(0, 0, 0, defaultAlpha);
        }
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
