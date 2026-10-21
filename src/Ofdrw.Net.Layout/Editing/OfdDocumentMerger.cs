using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using System.Threading;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

public sealed class OfdDocumentMergeOptions
{
    public bool IncludeTemplates { get; set; } = true;

    public bool IncludeAttachments { get; set; } = true;

    public bool SkipUnsupportedRawElements { get; set; }

    // Mix requires complete attribute closure; Merge keeps its legacy benign
    // unqualified metadata compatibility while rejecting foreign references.
    internal bool RequireKnownAttributes { get; set; }
}

/// <summary>
/// Merges pages into a self-contained document model. Template content is
/// flattened into each destination page so source package paths and IDs cannot
/// become dangling references.
/// </summary>
public static class OfdDocumentMerger
{
    public static OfdDocumentPackage Merge(
        IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options = null)
    {
        return MergeWithResult(sources, options).Package;
    }

    /// <summary>Merges documents with cooperative cancellation between resource and object copies.</summary>
    public static OfdDocumentPackage Merge(IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options, CancellationToken cancellationToken) => MergeCore(sources, options, cancellationToken).Package;

    /// <summary>Merges documents and reports objects dropped under the explicit skip option.</summary>
    public static OfdDocumentMergeResult MergeWithResult(
        IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options = null) => MergeCore(sources, options, default);

    /// <summary>Merges documents with cooperative cancellation and reports objects dropped under the explicit skip option.</summary>
    /// <param name="sources">The documents to merge in enumeration order.</param>
    /// <param name="options">Optional merge behavior.</param>
    /// <param name="cancellationToken">Cancellation checked before source access and between resource and object copies.</param>
    /// <returns>The merged document and diagnostics for skipped objects.</returns>
    /// <exception cref="OperationCanceledException">The cancellation token was canceled.</exception>
    public static OfdDocumentMergeResult MergeWithResult(
        IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options,
        CancellationToken cancellationToken) => MergeCore(sources, options, cancellationToken);

    internal static OfdDocumentPackage MergeForSplit(OfdDocumentPackage source, OfdModelMaterialization material, CancellationToken token)
        => MergeCore(new[] { source }, null, token, material, preserveAttachmentIdsForSplit: true).Package;

    internal static OfdDocumentPackage MergeForMix(IEnumerable<OfdDocumentPackage> sources, OfdDocumentMergeOptions options, OfdModelMaterialization material, CancellationToken token)
        => MergeCore(sources, options, token, material).Package;

    private static OfdDocumentMergeResult MergeCore(IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options, CancellationToken cancellationToken, OfdModelMaterialization? material = null,
        bool preserveAttachmentIdsForSplit = false)
    {
        cancellationToken.ThrowIfCancellationRequested();
        if (sources is null)
        {
            throw new ArgumentNullException(nameof(sources));
        }

        options ??= new OfdDocumentMergeOptions();
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        var sourceList = new List<OfdDocumentPackage>();
        using (var enumerator = sources.GetEnumerator())
        {
            while (true)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var hasSource = enumerator.MoveNext();
                cancellationToken.ThrowIfCancellationRequested();
                if (!hasSource) break;
                var source = enumerator.Current;
                cancellationToken.ThrowIfCancellationRequested(); material.Structure.Node(1); material.Structure.Reference();
                if (source is null) throw new ArgumentException("Source documents cannot contain null.", nameof(sources));
                material.InspectStorage(source, cancellationToken); material.GuardPreservedXml(source, cancellationToken);
                if (source.CompositeResources is not null || OfdCompositeGraph.Definitions(source).Any(e => e.Clips is not null))
                    throw new NotSupportedException("Merge cannot remap typed composites or clips.");
                OfdAnnotationContract.Validate(source, material, cancellationToken);
                if (OfdAnnotationContract.HasTyped(source)) throw new NotSupportedException("Merger cannot remap typed annotation ownership.");
                if (OfdNavigationXml.HasTypedNavigation(source, material, cancellationToken)) throw new NotSupportedException("Merge cannot remap navigation source ownership.");
                sourceList.Add(source);
            }
        }
        if (sourceList.Count == 0)
        {
            throw new ArgumentException("At least one source document is required.", nameof(sources));
        }

        if (sourceList.Any(source => source is null)) throw new ArgumentException("Source documents cannot contain null.", nameof(sources));
        var diagnostics = new List<string>();
        var knownFonts = new Dictionary<string, OfdFontResource>(StringComparer.Ordinal);
        var first = sourceList[0];
        var destination = new OfdDocumentPackage
        {
            Options = CloneOptions(first.Options)
        };
        destination.Options.DocumentId = "Doc_0";
        if (first.Options.Metadata.Cover is string cover)
        {
            var path = material.ResolvePath("OFD.xml", cover, cancellationToken);
            if (!first.PreservedEntries.TryGetValue(path, out var bytes) ||
                path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase) ||
                OfdXmlContentProbe.LooksLikeXml(bytes, cancellationToken, count => material.ChargeWork(count, cancellationToken)))
                throw new NotSupportedException("Merge cannot retain a Cover without a distinct existing package entry.");
            material.ChargeCopy(bytes.LongLength, cancellationToken);
            destination.PreservedEntries.Add(path, bytes.ToArray());
        }

        foreach (var source in sourceList)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var fonts = CopyFonts(source, destination, knownFonts, cancellationToken, material);
            if (options.IncludeAttachments)
            {
                CopyAttachments(source, destination, cancellationToken, material, preserveIds: preserveAttachmentIdsForSplit);
            }

            foreach (var sourcePage in source.Pages.OrderBy(page => page.Index))
            {
                cancellationToken.ThrowIfCancellationRequested();
                var page = new OfdPage
                {
                    Index = destination.Pages.Count,
                    XMillimeters = sourcePage.XMillimeters,
                    YMillimeters = sourcePage.YMillimeters,
                    WidthMillimeters = sourcePage.WidthMillimeters,
                    HeightMillimeters = sourcePage.HeightMillimeters
                };

                string? lastLayer = null;
                var layerSequence = 0;
                foreach (var element in EnumerateElements(sourcePage, options.IncludeTemplates))
                {
                    try
                    {
                        cancellationToken.ThrowIfCancellationRequested();
                        ValidateStandaloneElement(element, options.RequireKnownAttributes);
                        var clone = OfdModelCloner.CloneElement(element, destination.Options.Namespace, material: material, cancellationToken: cancellationToken);
                        clone.ObjectId = null;
                        // Standalone validation proved the owned primitive root;
                        // its old document-unit ID is explicitly reconstructed.
                        var sourceXml = clone switch { OfdTextElement sourceText => sourceText.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, _ => null };
                        if (sourceXml is not null)
                        {
                            var root = OfdXmlEnvelope.Parse(sourceXml); root.SetAttributeValue("ID", null);
                            var rebuilt = OfdXmlEnvelope.Capture(root, material, cancellationToken);
                            if (clone is OfdTextElement t) t.SourceXml = rebuilt;
                            if (clone is OfdImageElement image) image.SourceXml = rebuilt;
                            if (clone is OfdPathElement path) path.SourceXml = rebuilt;
                        }
                        var layerKey = element.LayerId + "\u001f" + element.LayerType;
                        if (layerKey != lastLayer) layerSequence++;
                        lastLayer = layerKey;
                        clone.LayerId = $"merged-layer-{layerSequence}";
                        if (clone is OfdTextElement text)
                        {
                            var originalFont = OfdFontSelection.Resolve(source.Fonts, text);
                            if (originalFont is not null)
                            {
                                text.FontResourceId = fonts[originalFont].Id;
                                text.FontName = fonts[originalFont].FontName;
                            }
                            else text.FontResourceId = null;
                        }
                        page.Elements.Add(clone);
                    }
                    catch (NotSupportedException exception) when (options.SkipUnsupportedRawElements)
                    {
                        diagnostics.Add($"Page {sourcePage.Index + 1}, object {element.ObjectId}: {exception.Message}");
                    }
                }

                destination.Pages.Add(page);
            }
        }

        return new OfdDocumentMergeResult(destination, diagnostics.AsReadOnly());
    }

    private static IEnumerable<OfdElement> EnumerateElements(OfdPage page, bool includeTemplates)
    {
        if (includeTemplates)
        {
            foreach (var template in page.Templates.Where(template => string.Equals(
                template.ZOrder,
                "Background",
                StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var element in template.Elements)
                {
                    yield return element;
                }
            }
        }

        foreach (var element in page.Elements)
        {
            yield return element;
        }

        if (includeTemplates)
        {
            foreach (var template in page.Templates.Where(template => !string.Equals(
                template.ZOrder,
                "Background",
                StringComparison.OrdinalIgnoreCase)))
            {
                foreach (var element in template.Elements)
                {
                    yield return element;
                }
            }
        }
        foreach (var element in page.AnnotationAppearances) yield return element;
    }

    private static void ValidateStandaloneElement(OfdElement element, bool requireKnownAttributes)
    {
        if (element is OfdRawElement)
            throw new NotSupportedException("Unsupported raw page objects cannot be safely remapped during merge.");
        var xml = element switch
        {
            OfdTextElement text => text.SourceXml,
            OfdImageElement image => image.SourceXml,
            OfdPathElement path => path.SourceXml,
            _ => null
        };
        if (!string.IsNullOrWhiteSpace(element.ClippingXml))
        {
            var clips = XElement.Parse(element.ClippingXml!);
            if (!OfdGraphicXmlContract.HasKnownChildren(clips) || clips.DescendantsAndSelf().Attributes().Any(attribute => !OfdGraphicXmlContract.IsKnownAttribute(attribute)))
                throw new NotSupportedException("Unknown clip structure or attributes cannot be safely remapped.");
            if (clips.Descendants().Any(node => node.Name.Namespace != clips.Name.Namespace))
                throw new NotSupportedException("Unknown clip extension elements cannot be safely remapped.");
            if (clips.DescendantsAndSelf().Attributes().Any(attribute => attribute.Name.LocalName is "Font" or "ResourceID" or "DrawParam" or "ColorSpace" or "RefID" or "ObjectRef"))
                throw new NotSupportedException("Unmodeled clip resource references cannot be safely remapped.");
        }
        if (string.IsNullOrWhiteSpace(xml)) return;
        var root = XElement.Parse(xml!, LoadOptions.PreserveWhitespace);
        if (root.Descendants().Any(node => node.Name.Namespace != root.Name.Namespace))
            throw new NotSupportedException("Unknown object extension elements cannot be safely remapped during merge.");
        foreach (var attribute in root.DescendantsAndSelf().Attributes())
        {
            var name = attribute.Name.LocalName;
            if (attribute.Parent == root && attribute.Name.Namespace == XNamespace.None &&
                ((name == "Font" && element is OfdTextElement) ||
                 (name == "ResourceID" && element is OfdImageElement))) continue;
            if (name is "Font" or "ResourceID" or "DrawParam" or "ColorSpace" or "RefID" or
                "TemplateID" or "PageID" or "ObjectRef")
                throw new NotSupportedException($"Unmodeled resource reference '{name}' cannot be safely remapped during merge.");
        }
        if (OfdGraphicXmlContract.HasUnsupportedReferences(root))
            throw new NotSupportedException("Unmodeled drawing resource references (including Substitution) cannot be safely remapped.");
        if (requireKnownAttributes && root.DescendantsAndSelf().Attributes("Boundary").Any(attribute => !OfdBoxParser.TryParse(attribute.Value, out _)))
            throw new NotSupportedException("Mix cannot flatten inexact drawing Boundary values.");
        if (!OfdGraphicXmlContract.HasKnownChildren(root) || root.DescendantsAndSelf().Attributes().Any(attribute =>
            !OfdGraphicXmlContract.IsKnownAttribute(attribute) && (requireKnownAttributes || attribute.Name.Namespace != XNamespace.None)))
            throw new NotSupportedException("Unknown object structure or attributes cannot be safely remapped during merge.");
        if (root.Descendants().Any(node => node.Name.LocalName is "Actions" or "Action"))
            throw new NotSupportedException("Page-object actions cannot be safely remapped during merge.");
    }

    private static Dictionary<OfdFontResource, OfdFontResource> CopyFonts(
        OfdDocumentPackage source,
        OfdDocumentPackage destination,
        IDictionary<string, OfdFontResource> knownFonts, CancellationToken cancellationToken, OfdModelMaterialization material)
    {
        var mapping = new Dictionary<OfdFontResource, OfdFontResource>();
        foreach (var font in source.Fonts)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var suffixLength = 3L + (font.Bold ? 4 : 5) + (font.Italic ? 4 : 5) + (font.Charset?.Length ?? 0);
            material.ChargeCopy(2L * (font.Data.Length == 0 ? 3L * font.FontName.Length + 14 + 2 * suffixLength : 128 + 2 * suffixLength), cancellationToken);
            var identity = (font.Data.Length == 0 ? "system:" + font.FontName.ToUpperInvariant() : BinaryIdentity.Hash(font.Data))
                + $":{font.CollectionFaceIndex}:{font.Bold}:{font.Italic}:{font.Charset}";
            if (!knownFonts.TryGetValue(identity, out var target))
            {
                material.ChargeCopy(font.Data.LongLength, cancellationToken);
                target = new OfdFontResource
                {
                    Id = $"merged-font-{destination.Fonts.Count + 1}",
                    FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
                    Bold = font.Bold, Italic = font.Italic, FileName = font.FileName, Data = font.Data.ToArray(), CollectionFaceIndex = font.CollectionFaceIndex
                };
                destination.Fonts.Add(target);
                knownFonts.Add(identity, target);
            }
            mapping.Add(font, target);
        }
        return mapping;
    }

    internal static void CopyAttachments(OfdDocumentPackage source, OfdDocumentPackage destination, CancellationToken cancellationToken = default,
        OfdModelMaterialization? material = null, Dictionary<object, object>? retainedSourceMapping = null, bool preserveIds = false)
    {
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        OfdAttachmentSource.AssertCopyable(source, material, cancellationToken);
        if (retainedSourceMapping is not null && source.DocumentEntryPath != destination.DocumentEntryPath)
            throw new NotSupportedException("Retained attachment source requires the original document path.");
        foreach (var attachment in source.Attachments)
        {
            material.ChargeCopy(256L + attachment.Data.LongLength, cancellationToken);
            cancellationToken.ThrowIfCancellationRequested();
            var clone = new OfdAttachment
            {
                Id = retainedSourceMapping is not null || preserveIds ? attachment.Id : null,
                Name = attachment.Name,
                MediaType = attachment.MediaType,
                Format = attachment.Format,
                CreationDate = attachment.CreationDate,
                ModificationDate = attachment.ModificationDate,
                SizeKilobytes = attachment.SizeKilobytes,
                Visible = attachment.Visible,
                Usage = attachment.Usage,
                IsExternal = attachment.IsExternal,
                ExternalPath = attachment.ExternalPath,
                Data = attachment.Data.ToArray()
            };
            destination.Attachments.Add(clone);
            if (retainedSourceMapping is not null)
            {
                material.ChargeCopy(128, cancellationToken);
                if (retainedSourceMapping.ContainsKey(attachment)) throw new InvalidDataException("Ambiguous retained attachment ownership.");
                retainedSourceMapping.Add(attachment, clone);
            }
        }
        // Only the retained-path Split opts in. Merge/Mix still reconstruct
        // closed typed attachments without importing another document's owner.
        if (retainedSourceMapping is not null)
        {
            destination.AttachmentSource = source.AttachmentSource?.Clone(source, destination, retainedSourceMapping, material, cancellationToken);
            foreach (var removal in source.PhysicalAttachmentRemovals)
            {
                cancellationToken.ThrowIfCancellationRequested(); material.ChargeCopy(128, cancellationToken);
                destination.PhysicalAttachmentRemovals.Add(new OfdAttachmentPhysicalRemoval(removal.Ordinal, removal.PayloadPath));
            }
        }
    }

    internal static OfdDocumentOptions CloneOptions(OfdDocumentOptions source)
    {
        return new OfdDocumentOptions
        {
            DocType = source.DocType,
            Namespace = source.Namespace,
            EnableDeflateCompression = source.EnableDeflateCompression,
            FontEmbedding = new OfdFontEmbeddingOptions
            {
                Mode = source.FontEmbedding.Mode,
                MaximumFontBytes = source.FontEmbedding.MaximumFontBytes,
                MaximumCollectionBytes = source.FontEmbedding.MaximumCollectionBytes,
                MaximumUsedScalars = source.FontEmbedding.MaximumUsedScalars,
                MaximumCmapParseWork = source.FontEmbedding.MaximumCmapParseWork,
                MaximumGlyphClosureOperations = source.FontEmbedding.MaximumGlyphClosureOperations
            },
            DefaultPageWidthMillimeters = source.DefaultPageWidthMillimeters,
            DefaultPageHeightMillimeters = source.DefaultPageHeightMillimeters,
            Metadata = new OfdMetadata
            {
                Title = source.Metadata.Title,
                Author = source.Metadata.Author,
                Subject = source.Metadata.Subject,
                Keywords = source.Metadata.Keywords,
                Creator = source.Metadata.Creator,
                CreationDate = source.Metadata.CreationDate,
                ModificationDate = source.Metadata.ModificationDate,
                DocUsage = source.Metadata.DocUsage,
                Cover = source.Metadata.Cover,
                CreatorVersion = source.Metadata.CreatorVersion
            }
        };
    }

}
