using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

internal static class OfdAnnotationContract
{
    internal static IEnumerable<OfdElement> Definitions(OfdPage page)
    {
        foreach (var element in page.Elements) yield return element;
        foreach (var annotation in (page.Annotations ?? page.AnnotationOwnership.Original)?.Items ?? Enumerable.Empty<OfdAnnotation>())
            foreach (var element in annotation.Appearance.Elements) yield return element;
    }
    internal static bool HasTyped(OfdDocumentPackage package) => package.Pages.Any(p => p.Annotations is not null || p.AnnotationOwnership.Original is not null);
    internal static void Validate(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        var typed = HasTyped(package);
        if (typed) package.AnnotationIndexSource?.Validate(package, material, token);
        var owning = new HashSet<object>(OfdReferenceIdentity.Instance);
        foreach (var page in package.Pages)
        {
            page.AnnotationOwnership.Validate(page, material, token);
            foreach (var entry in typed ? page.AnnotationOwnership.Entries : Enumerable.Empty<KeyValuePair<string, string>>())
            {
                token.ThrowIfCancellationRequested();
                if (!package.PreservedEntries.TryGetValue(entry.Key, out var bytes) ||
                    OfdXmlEnvelope.Capture(material.ReadXml(entry.Key, bytes, token).Root!, material, token) != entry.Value)
                    throw new InvalidDataException("Annotation source entry changed behind its reader-owned slots.");
            }
            foreach (var annotation in (page.Annotations ?? page.AnnotationOwnership.Original)?.Items ?? Enumerable.Empty<OfdAnnotation>())
            {
                material.ChargeCopy(128, token);
                if (!owning.Add(annotation) || !owning.Add(annotation.Appearance)) throw new InvalidDataException("Repeated annotation/Appearance ownership.");
                foreach (var element in annotation.Appearance.Elements)
                {
                    token.ThrowIfCancellationRequested(); material.ChargeCopy(64, token);
                    if (!owning.Add(element)) throw new InvalidDataException("Repeated local annotation primitive ownership.");
                    if (new[] { element.XMillimeters, element.YMillimeters, element.WidthMillimeters, element.HeightMillimeters }.Any(value => double.IsNaN(value) || double.IsInfinity(value)) ||
                        !OfdBoxParser.IsWritablePositiveSide(element.WidthMillimeters) || !OfdBoxParser.IsWritablePositiveSide(element.HeightMillimeters))
                        throw new InvalidDataException("Local annotation primitives require finite explicit writable geometry.");
                    if (element is OfdTextElement localText && (!(localText.FontSizeMillimeters > 0) || double.IsInfinity(localText.FontSizeMillimeters)))
                        throw new InvalidDataException("Local annotation text requires a finite positive size.");
                    if (element is not (OfdTextElement or OfdImageElement or OfdPathElement)) throw new NotSupportedException("B supports direct Text/Image/Path appearance definitions.");
                    if (annotation.Source is null && (element is OfdTextElement { SourceXml: not null } || element is OfdImageElement { SourceXml: not null } || element is OfdPathElement { SourceXml: not null }))
                        throw new NotSupportedException("Fresh local primitives must be constructed without imported source XML.");
                }
            }
        }
    }
    internal static void ValidatePageSelection(OfdDocumentPackage package, IReadOnlyCollection<OfdPage> retained,
        OfdModelMaterialization material, CancellationToken token)
    {
        Validate(package, material, token);
        var removed = (package.SourceIdentity?.OriginalPages ?? package.Pages.ToArray()).Where(p => !retained.Contains(p)).ToArray();
        if (removed.Length == 0) return;
        var ids = new HashSet<string>(removed.Select(p => p.Id).OfType<string>());
        var docPath = package.DocumentEntryPath;
        if (docPath is null || !package.PreservedEntries.TryGetValue(docPath, out var docBytes)) return;
        var doc = material.ReadXml(docPath, docBytes, token); var ns = doc.Root!.Name.Namespace;
        foreach (var declaration in doc.Root.Elements().Where(n => n.Name.LocalName == "Annotations"))
        {
            token.ThrowIfCancellationRequested();
            if (declaration.Name != ns + "Annotations" || declaration.Attributes().Any(a => !a.IsNamespaceDeclaration) || declaration.HasElements || declaration.Nodes().Any(n => n is not XText and not XComment))
                throw new NotSupportedException("Opaque annotation declarations prevent page removal.");
            var path = material.ResolvePath(docPath, declaration.Value, token);
            if (!package.PreservedEntries.TryGetValue(path, out var bytes)) throw new NotSupportedException("Missing annotation index prevents page removal.");
            var index = material.ReadXml(path, bytes, token); var root = index.Root!;
            if (root.Name.LocalName != "Annotations" || root.Name.Namespace != ns || root.Attributes().Any(a => !a.IsNamespaceDeclaration) || root.Nodes().Any(n => n is XProcessingInstruction || n is XText text && !string.IsNullOrWhiteSpace(text.Value))) throw new NotSupportedException("Opaque annotation index prevents page removal.");
            foreach (var row in root.Elements())
            {
                if (row.Name != ns + "Page") throw new NotSupportedException("Opaque annotation row prevents page removal.");
                if (row.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "PageID") || row.Elements(ns + "FileLoc").Count() != 1 ||
                    row.Elements().Any(n => n.Name != ns + "FileLoc") || row.Nodes().Any(n => n is not XElement && (n is not XText t || !string.IsNullOrWhiteSpace(t.Value))))
                    throw new NotSupportedException("Cannot remove an annotation row with opaque content.");
                var file = row.Element(ns + "FileLoc")!;
                if (file.HasElements || file.Attributes().Any(a => !a.IsNamespaceDeclaration) || file.Nodes().Any(n => n is not XText))
                    throw new NotSupportedException("Opaque annotation FileLoc prevents page removal.");
                if (row.Attribute("PageID") is null) throw new NotSupportedException("Missing annotation PageID prevents page removal.");
                if (!ids.Contains(row.Attribute("PageID")!.Value)) continue;
                var filePath = material.ResolvePath(path, file.Value, token);
                if (!package.PreservedEntries.ContainsKey(filePath) || !removed.Any(p => p.AnnotationOwnership.Entries.ContainsKey(filePath)))
                    throw new NotSupportedException("Missing or unmodeled PageAnnot prevents page removal.");
                if (retained.Any(p => p.AnnotationOwnership.Entries.ContainsKey(filePath))) throw new NotSupportedException("Shared PageAnnot prevents page removal.");
                if (removed.Any(p => p.AnnotationOwnership.UnsafeDeletionEntries.Contains(filePath) || p.AnnotationOwnership.HasUnclaimedSource && p.AnnotationOwnership.Entries.ContainsKey(filePath)))
                    throw new NotSupportedException("Unclaimed annotation source prevents destructive page removal.");
            }
        }
    }
}
