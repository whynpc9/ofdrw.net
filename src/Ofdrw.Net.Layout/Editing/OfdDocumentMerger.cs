using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

public sealed class OfdDocumentMergeOptions
{
    public bool IncludeTemplates { get; set; } = true;

    public bool IncludeAttachments { get; set; } = true;

    public bool SkipUnsupportedRawElements { get; set; }
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

    /// <summary>Merges documents and reports objects dropped under the explicit skip option.</summary>
    public static OfdDocumentMergeResult MergeWithResult(
        IEnumerable<OfdDocumentPackage> sources,
        OfdDocumentMergeOptions? options = null)
    {
        if (sources is null)
        {
            throw new ArgumentNullException(nameof(sources));
        }

        options ??= new OfdDocumentMergeOptions();
        var sourceList = sources.ToList();
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

        foreach (var source in sourceList)
        {
            var fonts = CopyFonts(source, destination, knownFonts);
            if (options.IncludeAttachments)
            {
                CopyAttachments(source, destination);
            }

            foreach (var sourcePage in source.Pages.OrderBy(page => page.Index))
            {
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
                        ValidateStandaloneElement(element);
                        var clone = OfdModelCloner.CloneElement(element, destination.Options.Namespace);
                        clone.ObjectId = null;
                        var layerKey = element.LayerId + "\u001f" + element.LayerType;
                        if (layerKey != lastLayer) layerSequence++;
                        lastLayer = layerKey;
                        clone.LayerId = $"merged-layer-{layerSequence}";
                        if (clone is OfdTextElement text)
                        {
                            var originalFont = source.Fonts.FirstOrDefault(font => font.Id == text.FontResourceId)
                                ?? source.Fonts.FirstOrDefault(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase) && !font.Bold && !font.Italic)
                                ?? source.Fonts.FirstOrDefault(font => string.Equals(font.FontName, text.FontName, StringComparison.OrdinalIgnoreCase));
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
    }

    private static void ValidateStandaloneElement(OfdElement element)
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
        if (string.IsNullOrWhiteSpace(xml)) return;
        var root = XElement.Parse(xml!, LoadOptions.PreserveWhitespace);
        foreach (var attribute in root.DescendantsAndSelf().Attributes())
        {
            var name = attribute.Name.LocalName;
            if (attribute.Parent == root &&
                ((name == "Font" && element is OfdTextElement) ||
                 (name == "ResourceID" && element is OfdImageElement))) continue;
            if (name is "Font" or "ResourceID" or "DrawParam" or "ColorSpace" or "RefID" or
                "TemplateID" or "PageID" or "ObjectRef")
                throw new NotSupportedException($"Unmodeled resource reference '{name}' cannot be safely remapped during merge.");
        }
        if (root.Descendants().Any(node => node.Name.LocalName is "Actions" or "Action"))
            throw new NotSupportedException("Page-object actions cannot be safely remapped during merge.");
    }

    private static Dictionary<OfdFontResource, OfdFontResource> CopyFonts(
        OfdDocumentPackage source,
        OfdDocumentPackage destination,
        IDictionary<string, OfdFontResource> knownFonts)
    {
        var mapping = new Dictionary<OfdFontResource, OfdFontResource>();
        foreach (var font in source.Fonts)
        {
            var identity = (font.Data.Length == 0 ? "system:" + font.FontName.ToUpperInvariant() : BinaryIdentity.Hash(font.Data))
                + $":{font.Bold}:{font.Italic}:{font.Charset}";
            if (!knownFonts.TryGetValue(identity, out var target))
            {
                target = new OfdFontResource
                {
                    Id = $"merged-font-{destination.Fonts.Count + 1}",
                    FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
                    Bold = font.Bold, Italic = font.Italic, FileName = font.FileName, Data = font.Data.ToArray()
                };
                destination.Fonts.Add(target);
                knownFonts.Add(identity, target);
            }
            mapping.Add(font, target);
        }
        return mapping;
    }

    private static void CopyAttachments(OfdDocumentPackage source, OfdDocumentPackage destination)
    {
        foreach (var attachment in source.Attachments)
        {
            destination.Attachments.Add(new OfdAttachment
            {
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
            });
        }
    }

    private static OfdDocumentOptions CloneOptions(OfdDocumentOptions source)
    {
        return new OfdDocumentOptions
        {
            DocType = source.DocType,
            Namespace = source.Namespace,
            EnableDeflateCompression = source.EnableDeflateCompression,
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
                ModificationDate = source.Metadata.ModificationDate
            }
        };
    }

}
