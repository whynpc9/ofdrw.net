using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Selects pages into an independent package. Save with OfdPackageWriter to prune deleted private resources and invalidate signatures.</summary>
public static class OfdDocumentSplitter
{
    /// <summary>Copies selected zero-based pages in list order. Empty/duplicate/out-of-range selections and multi-document packages are rejected.</summary>
    public static OfdDocumentPackage Split(OfdDocumentPackage source, IReadOnlyList<int> pages, CancellationToken cancellationToken = default)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        ValidateSingleDocument(source);
        ValidatePages(source, pages);
        if (!source.PreservedEntries.ContainsKey("OFD.xml"))
        {
            if (source.PreservedEntries.Count > 0 || source.PreservedCommonDataElements.Count > 0 || source.PreservedDocumentElements.Count > 0 || source.PreservedDocBodyElements.Count > 0)
                throw new NotSupportedException("Persist packages with preserved extensions before splitting.");
            var selected = new OfdDocumentPackage { Options = source.Options };
            selected.Pages.AddRange(pages.Select(index => source.Pages[index]));
            // Without an archive baseline, flatten typed template/annotation content
            // and copy only the fonts actually used by these pages.
            var texts = selected.Pages.SelectMany(page => page.Elements.Concat(page.Templates.SelectMany(template => template.Elements)).Concat(page.AnnotationAppearances)).OfType<OfdTextElement>();
            foreach (var text in texts)
            {
                var font = source.Fonts.FirstOrDefault(candidate => candidate.Id == text.FontResourceId)
                    ?? source.Fonts.FirstOrDefault(candidate => string.Equals(candidate.FontName, text.FontName, StringComparison.OrdinalIgnoreCase) && !candidate.Bold && !candidate.Italic)
                    ?? source.Fonts.FirstOrDefault(candidate => string.Equals(candidate.FontName, text.FontName, StringComparison.OrdinalIgnoreCase));
                if (font is not null && !selected.Fonts.Contains(font)) selected.Fonts.Add(font);
            }
            selected.Attachments.AddRange(source.Attachments);
            // Merger sorts by Index, so clone pages into the requested selection order.
            selected.Pages.Clear();
            foreach (var index in pages)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var page = OfdModelCloner.ClonePage(source.Pages[index]); page.Index = selected.Pages.Count; selected.Pages.Add(page);
            }
            var standalone = OfdDocumentMerger.Merge(new[] { selected }, null, cancellationToken);
            foreach (var tag in source.CustomTags) standalone.CustomTags.Add(tag.Key, tag.Value);
            return standalone;
        }
        foreach (var entry in source.PreservedEntries.Where(pair => pair.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)))
        {
            cancellationToken.ThrowIfCancellationRequested();
            // Opaque XML prevents proving resource closure. Do not return a deceptively clean split.
            using var stream = new MemoryStream(entry.Value, false);
            var xml = XDocument.Load(stream);
            if (entry.Key == "OFD.xml" && xml.Root!.Elements().Count(node => node.Name.LocalName == "DocBody") != 1)
                throw new NotSupportedException("Page editing requires a single DocBody.");
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
        foreach (var entry in source.PreservedEntries) result.PreservedEntries.Add(entry.Key, entry.Value.ToArray());
        foreach (var font in source.Fonts) result.Fonts.Add(new OfdFontResource
        {
            Id = font.Id, FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
            Bold = font.Bold, Italic = font.Italic, FileName = font.FileName, Data = font.Data.ToArray()
        });
        OfdDocumentMerger.CopyAttachments(source, result);
        foreach (var tag in source.CustomTags) result.CustomTags.Add(tag.Key, tag.Value);
        foreach (var index in pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var page = OfdModelCloner.ClonePage(source.Pages[index], keepSourcePath: true);
            page.Index = result.Pages.Count;
            result.Pages.Add(page);
        }
        return result;
    }

    internal static void ValidateSingleDocument(OfdDocumentPackage package)
    {
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var bytes)) return;
        using var input = new MemoryStream(bytes, false);
        var root = XDocument.Load(input).Root!;
        if (root.Elements().Count(node => node.Name == root.Name.Namespace + "DocBody") != 1)
            throw new NotSupportedException("Page editing requires a single DocBody.");
    }

    internal static void ValidatePages(OfdDocumentPackage package, IReadOnlyList<int> pages)
    {
        if (pages is null) throw new ArgumentNullException(nameof(pages));
        if (pages.Count == 0 || pages.Distinct().Count() != pages.Count || pages.Any(index => index < 0 || index >= package.Pages.Count))
            throw new ArgumentException("Select one or more distinct zero-based page positions within the document.", nameof(pages));
    }
}
