using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.IO;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>A source document and zero-based page position for mixing.</summary>
public sealed class OfdMixSource
{
    /// <summary>Creates a page selection; the source remains unchanged.</summary>
    public OfdMixSource(OfdDocumentPackage package, int pageIndex) { Package = package ?? throw new ArgumentNullException(nameof(package)); PageIndex = pageIndex; }
    /// <summary>Source document.</summary>
    public OfdDocumentPackage Package { get; }
    /// <summary>Zero-based page position.</summary>
    public int PageIndex { get; }
}

/// <summary>Overlays pages in input order without scaling, using the first source PhysicalBox.</summary>
public static class OfdDocumentMixer
{
    /// <summary>Flattens each source's background templates, body, foreground templates and visible annotation appearances, then overlays the next source.
    /// Unsupported raw/referencing content fails. Source packages remain unchanged; existing signatures are not copied.</summary>
    public static OfdDocumentPackage Mix(IEnumerable<OfdMixSource> sources, int maxSourceCount = 1_000,
        long maxExpandedBytes = 512L * 1024 * 1024, int maxObjectCount = 100_000, CancellationToken cancellationToken = default)
    {
        if (sources is null) throw new ArgumentNullException(nameof(sources));
        if (maxSourceCount <= 0 || maxExpandedBytes <= 0 || maxObjectCount <= 0) throw new ArgumentOutOfRangeException(nameof(maxSourceCount));
        var selected = new List<OfdDocumentPackage>();
        long bytes = 0;
        int objects = 0;
        foreach (var item in sources)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (item is null || selected.Count >= maxSourceCount) throw new ArgumentException("Mix source limit exceeded or null source.");
            OfdDocumentSplitter.ValidateSingleDocument(item.Package);
            OfdDocumentSplitter.ValidatePages(item.Package, new[] { item.PageIndex });
            var documentNamespace = EntryNamespace(item.Package, item.Package.DocumentEntryPath);
            if (item.Package.PreservedDocumentElements.Any(xml => XElement.Parse(xml).Name != documentNamespace + "Annotations") ||
                item.Package.PreservedCommonDataElements.Any(xml => XElement.Parse(xml).Name != documentNamespace + "TemplatePage"))
                throw new NotSupportedException("Mix cannot safely remap document extensions or shared drawing resources.");
            var page = item.Package.Pages[item.PageIndex];
            var pageNamespace = EntryNamespace(item.Package, page.SourceEntryPath);
            foreach (var reference in page.PreservedPageElements.Select(XElement.Parse).Where(node => node.Name == pageNamespace + "Template"))
                if (!page.Templates.Any(template => template.TemplateId == reference.Attribute("TemplateID")?.Value))
                    throw new NotSupportedException("Mix cannot flatten an unresolved template reference.");
            objects = checked(objects + page.Elements.Count + page.Templates.Sum(template => template.Elements.Count) + page.AnnotationAppearances.Count);
            bytes = checked(bytes + item.Package.PreservedEntries.Values.Sum(data => (long)data.Length) + item.Package.Fonts.Sum(font => (long)font.Data.Length) + item.Package.Attachments.Sum(attachment => (long)attachment.Data.Length)
                + page.Elements.Concat(page.Templates.SelectMany(template => template.Elements)).Concat(page.AnnotationAppearances)
                    .OfType<OfdImageElement>().Sum(image => (long)image.Data.Length));
            if (bytes > maxExpandedBytes || objects > maxObjectCount) throw new ArgumentException("Mix expanded-byte/object budget exceeded.");
            if (page.PreservedPageElements.Any(xml => XElement.Parse(xml).Name != pageNamespace + "Template"))
                throw new NotSupportedException("Mix cannot safely remap page extensions/actions.");
            var package = new OfdDocumentPackage { Options = item.Package.Options };
            package.Pages.Add(page); package.Fonts.AddRange(item.Package.Fonts);
            package.Attachments.AddRange(item.Package.Attachments);
            selected.Add(package);
        }
        if (selected.Count == 0) throw new ArgumentException("At least one source page is required.", nameof(sources));
        var merged = OfdDocumentMerger.Merge(selected, null, cancellationToken);
        var first = merged.Pages[0];
        var target = new OfdPage { WidthMillimeters = first.WidthMillimeters, HeightMillimeters = first.HeightMillimeters,
            XMillimeters = first.XMillimeters, YMillimeters = first.YMillimeters };
        foreach (var page in merged.Pages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            foreach (var element in page.Elements)
            {
                // Per-source layer identity preserves overlay order despite writer grouping.
                element.LayerId = $"mix-{page.Index}-{element.LayerId}";
                element.LayerType = "Body";
                target.Elements.Add(element);
            }
        }
        merged.Pages.Clear(); merged.Pages.Add(target);
        return merged;
    }
    private static XNamespace EntryNamespace(OfdDocumentPackage package, string? path)
    {
        if (path is not null && package.PreservedEntries.TryGetValue(path, out var bytes))
        {
            using var input = new MemoryStream(bytes, false);
            return XDocument.Load(input).Root!.Name.Namespace;
        }
        return XNamespace.Get(package.Options.Namespace);
    }

}
