using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using System.Xml;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Selects pages into an independent package. Save with OfdPackageWriter to prune deleted private resources and invalidate signatures.</summary>
public static class OfdDocumentSplitter
{
    /// <summary>Copies selected zero-based pages in list order. Empty/duplicate/out-of-range selections and multi-document packages are rejected.</summary>
    public static OfdDocumentPackage Split(OfdDocumentPackage source, IReadOnlyList<int> pages, CancellationToken cancellationToken = default)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(source, cancellationToken); material.GuardPreservedXml(source, cancellationToken);
        OfdCompositeGraph.Validate(source, material, cancellationToken);
        ValidateSingleDocument(source, material, cancellationToken);
        pages = ValidatePages(source, pages, material, cancellationToken);
        var selectedPages = pages.Select(index => source.Pages[index]).ToArray();
        OfdAnnotationContract.ValidatePageSelection(source, selectedPages, material, cancellationToken);
        OfdNavigationXml.Validate(source, selectedPages, cancellationToken, material);
        source.SourceEnvelope?.ValidatePageOrder(source, selectedPages, material, cancellationToken);
        foreach (var index in pages) OfdPageXmlContract.ValidateWritableDimensions(source.Pages[index].WidthMillimeters, source.Pages[index].HeightMillimeters);
        OfdPageXmlContract.ValidateDocumentArea(source, cancellationToken, material);
        if (!source.PreservedEntries.ContainsKey("OFD.xml") && source.CompositeResources is null)
        {
            if (source.PreservedEntries.Count > 0 || source.PreservedCommonDataElements.Count > 0 || source.PreservedDocumentElements.Count > 0 || source.PreservedDocBodyElements.Count > 0 ||
                pages.Any(index => source.Pages[index].PreservedPageElements.Count > 0))
                throw new NotSupportedException("Persist packages with preserved extensions before splitting.");
            var selected = new OfdDocumentPackage { Options = source.Options };
            selected.Pages.AddRange(pages.Select(index => source.Pages[index]));
            // Without an archive baseline, flatten typed template/annotation content
            // and copy only the fonts actually used by these pages.
            var texts = selected.Pages.SelectMany(page => page.Elements.Concat(page.Templates.SelectMany(template => template.Elements)).Concat(page.AnnotationAppearances)).OfType<OfdTextElement>();
            foreach (var text in texts)
            {
                var font = OfdFontSelection.Resolve(source.Fonts, text);
                if (font is not null && !selected.Fonts.Contains(font)) selected.Fonts.Add(font);
            }
            selected.Attachments.AddRange(source.Attachments);
            // Merger sorts by Index, so clone pages into the requested selection order.
            selected.Pages.Clear();
            foreach (var index in pages)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var page = OfdModelCloner.ClonePage(source.Pages[index], material: material, cancellationToken: cancellationToken); page.Index = selected.Pages.Count; selected.Pages.Add(page);
            }
            if (OfdAnnotationContract.HasTyped(source))
            {
                if (selected.Pages.Any(page => page.Templates.Count != 0)) throw new NotSupportedException("Fresh typed annotation split with templates requires persisted source ownership.");
                selected.Options = OfdDocumentMerger.CloneOptions(source.Options);
                selected.Fonts.Clear();
                foreach (var font in source.Fonts)
                {
                    material.ChargeCopy(256L + font.Data.LongLength, cancellationToken);
                    selected.Fonts.Add(new OfdFontResource { Id = font.Id, FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
                        Bold = font.Bold, Italic = font.Italic, FileName = font.FileName, Data = font.Data.ToArray(), CollectionFaceIndex = font.CollectionFaceIndex });
                }
                selected.Attachments.Clear(); OfdDocumentMerger.CopyAttachments(source, selected, cancellationToken, material);
                var typedPageMap = new Dictionary<object, object>();
                for (var position = 0; position < pages.Count; position++) typedPageMap.Add(source.Pages[pages[position]], selected.Pages[position]);
                OfdNavigationCloner.Clone(source, selected, typedPageMap, material, cancellationToken);
                foreach (var tag in source.CustomTags) selected.CustomTags.Add(tag.Key, tag.Value);
                return selected;
            }
            var standalone = OfdDocumentMerger.MergeForSplit(selected, material, cancellationToken);
            var pageMap = new Dictionary<object, object>();
            for (var position = 0; position < pages.Count; position++) pageMap.Add(source.Pages[pages[position]], standalone.Pages[position]);
            OfdNavigationCloner.Clone(source, standalone, pageMap, material, cancellationToken);
            foreach (var tag in source.CustomTags) standalone.CustomTags.Add(tag.Key, tag.Value);
            return standalone;
        }
        foreach (var entry in source.PreservedEntries.Where(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (!OfdXmlContentProbe.LooksLikeXml(entry.Value)) continue;
            try { material.ReadXml(entry.Key, entry.Value, cancellationToken); }
            catch (XmlException exception) { throw new NotSupportedException("Split cannot inspect malformed XML-looking control or extension entries.", exception); }
        }
        foreach (var index in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            OfdPageXmlContract.ValidateForRewrite(source, source.Pages[index], cancellationToken, material);
        }
        var result = new OfdDocumentPackage
        {
            Options = OfdDocumentMerger.CloneOptions(source.Options), DocumentEntryPath = source.DocumentEntryPath,
            PublicResourceLocation = source.PublicResourceLocation, DocumentResourceLocation = source.DocumentResourceLocation
        };
        result.Options.DocumentId = source.Options.DocumentId;
        result.PreservedCommonDataElements.AddRange(source.PreservedCommonDataElements);
        result.PreservedDocumentElements.AddRange(source.PreservedDocumentElements);
        result.PreservedDocBodyElements.AddRange(source.PreservedDocBodyElements);
        var mapping = new Dictionary<object, object>();
        OfdCompositeGraph.CloneResources(source, result, mapping, material, cancellationToken, null);
        foreach (var entry in source.PreservedEntries)
        { material.ChargeCopy(entry.Value.LongLength, cancellationToken); result.PreservedEntries.Add(entry.Key, entry.Value.ToArray()); }
        foreach (var font in source.Fonts)
        {
            material.ChargeCopy(font.Data.LongLength, cancellationToken);
            var clone = new OfdFontResource
            {
                Id = font.Id, FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
                Bold = font.Bold, Italic = font.Italic, FileName = font.FileName, Data = font.Data.ToArray(), CollectionFaceIndex = font.CollectionFaceIndex
            };
            result.Fonts.Add(clone); mapping.Add(font, clone);
        }
        OfdDocumentMerger.CopyAttachments(source, result, cancellationToken, material);
        foreach (var tag in source.CustomTags) result.CustomTags.Add(tag.Key, tag.Value);
        foreach (var index in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = OfdModelCloner.ClonePage(source.Pages[index], keepSourcePath: true, material: material, cancellationToken: cancellationToken, mapping: mapping);
            page.Index = result.Pages.Count;
            result.Pages.Add(page);
        }
        if (OfdAnnotationContract.HasTyped(source))
        {
            foreach (var historical in source.SourceIdentity?.OriginalPages ?? Array.Empty<OfdPage>())
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (!mapping.ContainsKey(historical)) OfdModelCloner.ClonePage(historical, keepSourcePath: true, material: material, cancellationToken: cancellationToken, mapping: mapping);
            }
        }
        OfdNavigationCloner.Clone(source, result, mapping, material, cancellationToken);
        result.AnnotationIndexSource = source.AnnotationIndexSource?.Clone(source, result, material, cancellationToken);
        result.SourceEnvelope = source.SourceEnvelope?.Clone(mapping, material, cancellationToken);
        result.SourceIdentity = source.SourceIdentity?.CloneFor(result, mapping, material, cancellationToken);
        return result;
    }

    internal static void ValidateSingleDocument(OfdDocumentPackage package, OfdModelMaterialization? material = null, CancellationToken cancellationToken = default)
    {
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var bytes)) return;
        if (material is null) { material = new OfdModelMaterialization(new OfdModelLimits()); material.InspectStorage(package, cancellationToken); material.GuardPreservedXml(package, cancellationToken); }
        var root = material.ReadXml("OFD.xml", bytes, cancellationToken).Root!;
        if (root.Elements().Count(node => node.Name == root.Name.Namespace + "DocBody") != 1)
            throw new NotSupportedException("Page editing requires a single DocBody.");
    }

    internal static IReadOnlyList<int> ValidatePages(OfdDocumentPackage package, IReadOnlyList<int> pages,
        OfdModelMaterialization? material = null, CancellationToken token = default)
    {
        if (pages is null) throw new ArgumentNullException(nameof(pages));
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        var count = pages.Count;
        if (count == 0 || count > package.Pages.Count) throw new ArgumentException("Select one or more distinct zero-based page positions within the document.", nameof(pages));
        material.ChargeCopy(sizeof(int) * (long)count, token);
        var snapshot = new List<int>(count); var seen = new HashSet<int>();
        foreach (var index in pages)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (snapshot.Count >= count || index < 0 || index >= package.Pages.Count || !seen.Add(index))
                throw new ArgumentException("Select one or more distinct zero-based page positions within the document.", nameof(pages));
            snapshot.Add(index);
        }
        if (seen.Count != count) throw new ArgumentException("Page selection count is inconsistent.", nameof(pages));
        return snapshot;
    }
}
