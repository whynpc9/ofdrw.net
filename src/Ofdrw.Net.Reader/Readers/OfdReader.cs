using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
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

        var ofdXml = XDocument.Parse(archive.ReadUtf8Text(OfdConstants.OfdRootFile));
        var ofdNs = ofdXml.Root?.Name.Namespace ?? XNamespace.Get(OfdConstants.Namespace);
        var docType = ofdXml.Root?.Attribute("DocType")?.Value ?? OfdConstants.DefaultDocType;

        var docRoot = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .Elements(ofdNs + "DocRoot")
            .Select(x => x.Value)
            .FirstOrDefault() ?? "Doc_0/Document.xml";

        docRoot = OfdPackagePath.Resolve("OFD.xml", docRoot);
        package.Options.DocType = docType;
        package.DocumentEntryPath = docRoot;
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

        var selectedDocBody = ofdXml.Root?
            .Elements(ofdNs + "DocBody")
            .FirstOrDefault(body => string.Equals(
                OfdPackagePath.Resolve("OFD.xml", body.Element(ofdNs + "DocRoot")?.Value ?? string.Empty),
                docRoot,
                StringComparison.Ordinal));
        foreach (var bodyElement in selectedDocBody?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (bodyElement.Name.LocalName is not "DocInfo" and not "DocRoot")
            {
                package.PreservedDocBodyElements.Add(
                    bodyElement.ToString(SaveOptions.DisableFormatting));
            }
        }

        var documentXml = XDocument.Parse(archive.ReadUtf8Text(docRoot));
        var docNs = documentXml.Root?.Name.Namespace ?? ofdNs;
        var commonData = documentXml.Root?.Element(docNs + "CommonData");
        foreach (var commonElement in commonData?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (commonElement.Name.LocalName is not "MaxUnitID" and
                not "PageArea" and
                not "PublicRes" and
                not "DocumentRes")
            {
                package.PreservedCommonDataElements.Add(
                    commonElement.ToString(SaveOptions.DisableFormatting));
            }
        }

        foreach (var documentElement in documentXml.Root?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (documentElement.Name.LocalName is not "CommonData" and
                not "Pages" and
                not "Attachments" and
                not "CustomTags")
            {
                package.PreservedDocumentElements.Add(
                    documentElement.ToString(SaveOptions.DisableFormatting));
            }
        }

        var defaultPageBox = ParseBox(commonData?
            .Element(docNs + "PageArea")?
            .Element(docNs + "PhysicalBox")?.Value);
        var fontMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documentMediaMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var documentMediaTypeMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        var publicResLoc = commonData?.Element(docNs + "PublicRes")?.Value;
        package.PublicResourceLocation = publicResLoc;
        if (!string.IsNullOrWhiteSpace(publicResLoc))
        {
            var publicResPath = Resolve(docRoot, publicResLoc!);
            if (archive.Contains(publicResPath))
            {
                ReadFonts(archive, publicResPath, package, fontMap);
                ReadMediaResources(archive, publicResPath, docNs, documentMediaMap, documentMediaTypeMap);
            }
        }

        var documentResLoc = commonData?.Element(docNs + "DocumentRes")?.Value;
        package.DocumentResourceLocation = documentResLoc;
        if (!string.IsNullOrWhiteSpace(documentResLoc))
        {
            var documentResPath = Resolve(docRoot, documentResLoc!);
            ReadMediaResources(archive, documentResPath, docNs, documentMediaMap, documentMediaTypeMap);
            ReadFonts(archive, documentResPath, package, fontMap);
        }

        var templateLocations = commonData?.Elements()
            .Where(x => x.Name.LocalName == "TemplatePage")
            .Select(x => new
            {
                Id = x.Attribute("ID")?.Value,
                BaseLoc = x.Attribute("BaseLoc")?.Value
            })
            .Where(x => !string.IsNullOrWhiteSpace(x.Id) && !string.IsNullOrWhiteSpace(x.BaseLoc))
            .ToDictionary(x => x.Id!, x => x.BaseLoc!, StringComparer.OrdinalIgnoreCase)
            ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

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

        foreach (var pageRef in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var contentPath = Resolve(docRoot, pageRef.BaseLoc!);
            var pageXml = XDocument.Parse(archive.ReadUtf8Text(contentPath));
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
                if (pageElement.Name.LocalName is not "Area" and not "Content")
                {
                    page.PreservedPageElements.Add(pageElement.ToString(SaveOptions.DisableFormatting));
                }
            }

            var pageDir = GetDirectory(contentPath);
            var pageResPath = $"{pageDir}/PageRes.xml";
            var mediaMap = new Dictionary<string, string>(documentMediaMap, StringComparer.OrdinalIgnoreCase);
            var mediaTypeMap = new Dictionary<string, string>(documentMediaTypeMap, StringComparer.OrdinalIgnoreCase);

            if (archive.Contains(pageResPath))
            {
                ReadMediaResources(archive, pageResPath, docNs, mediaMap, mediaTypeMap);
                ReadFonts(archive, pageResPath, package, fontMap);
            }

            foreach (var element in ParsePageObjects(
                archive,
                pageXml,
                fontMap,
                mediaMap,
                mediaTypeMap,
                pageResPath))
            {
                page.Elements.Add(element);
            }

            foreach (var templateRef in pageXml.Root?.Elements()
                .Where(x => x.Name.LocalName == "Template") ?? Enumerable.Empty<XElement>())
            {
                var templateId = templateRef.Attribute("TemplateID")?.Value;
                if (string.IsNullOrWhiteSpace(templateId) ||
                    !templateLocations.TryGetValue(templateId!, out var templateLoc))
                {
                    continue;
                }

                var templatePath = Resolve(docRoot, templateLoc);
                if (!archive.Contains(templatePath))
                {
                    continue;
                }

                var templateXml = XDocument.Parse(archive.ReadUtf8Text(templatePath));
                var templateResPath = $"{GetDirectory(templatePath)}/PageRes.xml";
                var templateMediaMap = new Dictionary<string, string>(documentMediaMap, StringComparer.OrdinalIgnoreCase);
                var templateMediaTypeMap = new Dictionary<string, string>(documentMediaTypeMap, StringComparer.OrdinalIgnoreCase);
                ReadMediaResources(archive, templateResPath, docNs, templateMediaMap, templateMediaTypeMap);
                ReadFonts(archive, templateResPath, package, fontMap);

                var template = new OfdTemplateContent
                {
                    TemplateId = templateId!,
                    ZOrder = templateRef.Attribute("ZOrder")?.Value ?? "Background",
                    BaseLocation = templateLoc
                };
                foreach (var element in ParsePageObjects(
                    archive,
                    templateXml,
                    fontMap,
                    templateMediaMap,
                    templateMediaTypeMap,
                    templateResPath))
                {
                    template.Elements.Add(element);
                }

                page.Templates.Add(template);
            }

            ReadAnnotationAppearances(archive, documentXml, docRoot, page, fontMap, mediaMap, mediaTypeMap, cancellationToken);
            package.Pages.Add(page);
        }

        var attachmentsLoc = documentXml.Root?.Elements()
            .FirstOrDefault(x => x.Name.LocalName == "Attachments")?.Value;
        if (!string.IsNullOrWhiteSpace(attachmentsLoc))
        {
            var attachmentsPath = Resolve(docRoot, attachmentsLoc!);
            if (archive.Contains(attachmentsPath))
            {
                var attachmentsXml = XDocument.Parse(archive.ReadUtf8Text(attachmentsPath));
                foreach (var node in attachmentsXml.Root?.Elements()
                    .Where(x => x.Name.LocalName == "Attachment") ?? Enumerable.Empty<XElement>())
                {
                    var fileLoc = node.Elements().FirstOrDefault(x => x.Name.LocalName == "FileLoc")?.Value
                        ?? node.Attribute("FileLoc")?.Value
                        ?? string.Empty;
                    var contentPath = Resolve(attachmentsPath, fileLoc);
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
            var customTagsPath = Resolve(docRoot, customTagsLoc!);
            if (archive.Contains(customTagsPath))
            {
                var customTagsXml = XDocument.Parse(archive.ReadUtf8Text(customTagsPath));
                foreach (var customTag in customTagsXml.Root?.Elements()
                    .Where(x => x.Name.LocalName == "CustomTag") ?? Enumerable.Empty<XElement>())
                {
                    var detailLoc = customTag.Elements().FirstOrDefault(x => x.Name.LocalName == "FileLoc")?.Value
                        ?? customTag.Attribute("FileLoc")?.Value;
                    if (string.IsNullOrWhiteSpace(detailLoc))
                    {
                        continue;
                    }

                    var detailPath = Resolve(customTagsPath, detailLoc!);
                    ReadKeyValueCustomTags(archive, detailPath, package.CustomTags);
                }
            }
        }
        else
        {
            var legacyCustomTagPath = $"{package.Options.DocumentId}/Tags/CustomTag_EMR.xml";
            ReadKeyValueCustomTags(archive, legacyCustomTagPath, package.CustomTags);
        }

        return package;
    }

    private static void ReadAnnotationAppearances(OfdPackageArchive archive, XDocument document, string documentPath,
        OfdPage page, IReadOnlyDictionary<string, string> fonts, IReadOnlyDictionary<string, string> media,
        IReadOnlyDictionary<string, string> mediaTypes, CancellationToken cancellationToken)
    {
        foreach (var declaration in document.Root!.Elements().Where(node => node.Name.LocalName == "Annotations"))
        {
            var listPath = OfdPackagePath.Resolve(documentPath, declaration.Value);
            var list = XDocument.Parse(archive.ReadUtf8Text(listPath));
            foreach (var record in list.Root!.Elements().Where(node => node.Name.LocalName == "Page" && node.Attribute("PageID")?.Value == page.Id))
            {
                var location = record.Elements().FirstOrDefault(node => node.Name.LocalName == "FileLoc")?.Value;
                if (string.IsNullOrWhiteSpace(location)) continue;
                var annotationPath = OfdPackagePath.Resolve(listPath, location!);
                var annotations = XDocument.Parse(archive.ReadUtf8Text(annotationPath));
                foreach (var annotation in annotations.Root!.Elements().Where(node => node.Name.LocalName == "Annot"))
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    if (annotation.Attribute("Visible")?.Value is "false" or "0") continue;
                    var appearance = annotation.Elements().FirstOrDefault(node => node.Name.LocalName == "Appearance");
                    if (appearance is null) continue;
                    if (appearance.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration && attribute.Name.LocalName is not ("Boundary" or "ID" or "CTM")))
                        throw new NotSupportedException("Unsupported annotation appearance attributes cannot be rendered safely.");
                    var appearanceTransform = ParseMatrix(appearance.Attribute("CTM")?.Value);
                    if (appearance.Attribute("CTM") is not null && (appearanceTransform is null || appearanceTransform.Any(value => double.IsNaN(value) || double.IsInfinity(value))))
                        throw new InvalidDataException("Invalid annotation appearance transform.");
                    var ns = appearance.Name.Namespace;
                    var xml = new XDocument(new XElement(ns + "Page", new XElement(ns + "Content",
                        new XElement(ns + "Layer", new XAttribute("ID", "annotation-" + annotation.Attribute("ID")?.Value), appearance.Elements()))));
                    var box = ParseBox(appearance.Attribute("Boundary")?.Value);
                    foreach (var element in ParsePageObjects(archive, xml, fonts, media, mediaTypes, annotationPath))
                    {
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
                        // Keep the serialized boundary consistent with the translated typed geometry.
                        var source = element switch { OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, _ => null };
                        if (source is not null)
                        {
                            var node = XElement.Parse(source);
                            node.SetAttributeValue("Boundary", string.Join(" ", new[] { element.XMillimeters, element.YMillimeters, element.WidthMillimeters, element.HeightMillimeters }.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));
                            if (appearanceTransform is not null)
                            {
                                var objectTransform = element switch { OfdTextElement matrixText => matrixText.Transform, OfdImageElement matrixImage => matrixImage.Transform, OfdPathElement matrixPath => matrixPath.Transform, _ => null };
                                if (objectTransform is not null) node.SetAttributeValue("CTM", string.Join(" ", objectTransform.Select(value => value.ToString("R", CultureInfo.InvariantCulture))));
                            }
                            source = node.ToString(SaveOptions.DisableFormatting);
                            if (element is OfdTextElement text) text.SourceXml = source;
                            if (element is OfdImageElement image) image.SourceXml = source;
                            if (element is OfdPathElement path) path.SourceXml = source;
                        }
                        page.AnnotationAppearances.Add(element);
                    }
                }
            }
        }
    }

    private static void ReadFonts(
        OfdPackageArchive archive, string resourcePath, OfdDocumentPackage package,
        IDictionary<string, string> fontMap)
    {
        if (!archive.Contains(resourcePath)) return;
        var xml = XDocument.Parse(archive.ReadUtf8Text(resourcePath));
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
                if (archive.TryGetBytes(Resolve(resourcePath, reference), out var bytes)) resource.Data = bytes;
            }
            package.Fonts.Add(resource);
        }
    }

    private static void ReadMediaResources(
        OfdPackageArchive archive,
        string resPath,
        XNamespace fallbackNs,
        IDictionary<string, string> mediaMap,
        IDictionary<string, string> mediaTypeMap)
    {
        if (!archive.Contains(resPath))
        {
            return;
        }

        var resXml = XDocument.Parse(archive.ReadUtf8Text(resPath));
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

            mediaMap[id!] = Resolve(resPath, mediaLoc);
            mediaTypeMap[id!] = ToMediaType(media.Attribute("Format")?.Value);
        }
    }

    private static IEnumerable<OfdElement> ParsePageObjects(
        OfdPackageArchive archive,
        XDocument pageXml,
        IReadOnlyDictionary<string, string> fontMap,
        IReadOnlyDictionary<string, string> mediaMap,
        IReadOnlyDictionary<string, string> mediaTypeMap,
        string pageResPath)
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
                    var textCodes = node.Descendants()
                        .Where(x => x.Name.LocalName == "TextCode")
                        .ToList();
                    var text = new OfdTextElement
                    {
                        ObjectId = node.Attribute("ID")?.Value,
                        LayerId = layerId,
                        LayerType = layerType,
                        XMillimeters = boundary.x,
                        YMillimeters = boundary.y,
                        WidthMillimeters = boundary.w,
                        HeightMillimeters = boundary.h,
                        Text = textCodes.Count == 0 ? node.Value : string.Concat(textCodes.Select(code => code.Value)),
                        FontResourceId = node.Attribute("Font")?.Value,
                        FontName = ResolveFontName(node.Attribute("Font")?.Value, fontMap),
                        FontSizeMillimeters = ParseDouble(node.Attribute("Size")?.Value, 4d),
                        Weight = (int)ParseDouble(node.Attribute("Weight")?.Value, OfdTextElement.DefaultWeight),
                        Italic = ParseBoolean(node.Attribute("Italic")?.Value, false),
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value),
                        FillColor = ParseColor(node.Elements()
                            .FirstOrDefault(x => x.Name.LocalName == "FillColor")) ?? OfdColor.Black,
                        SourceXml = node.ToString(SaveOptions.DisableFormatting)
                    };
                    foreach (var textCode in textCodes)
                    {
                        text.Runs.Add(new OfdTextRun
                        {
                            Text = textCode.Value,
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
                        ClipsXml = node.Elements().FirstOrDefault(child => child.Name.LocalName == "Clips")?.ToString(SaveOptions.DisableFormatting),
                        MediaType = mediaTypeMap.TryGetValue(resourceId, out var mediaType) ? mediaType : "image/png",
                        SourceXml = node.ToString(SaveOptions.DisableFormatting)
                    };

                    if (mediaMap.TryGetValue(resourceId, out var mediaFile))
                    {
                        var mediaPath = mediaFile.Contains('/') ? mediaFile : Resolve(pageResPath, mediaFile);
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
                        AbbreviatedData = node.Elements()
                            .FirstOrDefault(x => x.Name.LocalName == "AbbreviatedData")?.Value
                            ?? string.Empty,
                        Transform = ParseMatrix(node.Attribute("CTM")?.Value),
                        LineWidthMillimeters = ParseDouble(node.Attribute("LineWidth")?.Value, 0.353d),
                        Stroke = ParseBoolean(node.Attribute("Stroke")?.Value, true),
                        Fill = ParseBoolean(node.Attribute("Fill")?.Value, false),
                        StrokeColor = strokeColor ?? OfdColor.Black,
                        FillColor = fillColor,
                        SourceXml = node.ToString(SaveOptions.DisableFormatting)
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
                    Xml = node.ToString(SaveOptions.DisableFormatting),
                    XMillimeters = rawBoundary.x,
                    YMillimeters = rawBoundary.y,
                    WidthMillimeters = rawBoundary.w,
                    HeightMillimeters = rawBoundary.h
                };
            }
        }
    }

    private static void ReadKeyValueCustomTags(
        OfdPackageArchive archive,
        string detailPath,
        IDictionary<string, string> destination)
    {
        if (!archive.Contains(detailPath))
        {
            return;
        }

        var tagsXml = XDocument.Parse(archive.ReadUtf8Text(detailPath));
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

    private static string Resolve(string basePath, string relativePath) => OfdPackagePath.Resolve(basePath, relativePath);

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

        var parts = value!.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
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
