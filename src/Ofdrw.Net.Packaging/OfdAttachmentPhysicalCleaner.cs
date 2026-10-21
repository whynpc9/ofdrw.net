using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

// A physical request is deliberately narrower than ordinary attachment editing.
// Prove the complete source identity and both original and final file-reference
// closure before removing any entry from the output plan.
internal static class OfdAttachmentPhysicalCleaner
{
    private static readonly HashSet<string> ControlNames = new(StringComparer.OrdinalIgnoreCase)
    { "Document.xml", "PublicRes.xml", "DocumentRes.xml", "PageRes.xml", "Attachments.xml", "Annotations.xml", "Signatures.xml", "Signature.xml" };
    // A closed set for this cleanup profile. Unknown standard-looking names may
    // carry implicit references, so they cannot be admitted by string matching.
    private static readonly HashSet<string> KnownElements = new(StringComparer.Ordinal)
    {
        "OFD", "DocBody", "DocInfo", "DocID", "Title", "Author", "Subject", "Keywords", "Creator", "CreationDate", "ModDate", "DocRoot", "Versions", "Version", "Signatures", "Signature", "SignedInfo", "SignedValue", "References", "Reference", "FileRef", "DigestMethod", "DigestValue", "Provider", "SignatureMethod", "SignatureDateTime", "StampAnnot", "Seal",
        "Document", "CommonData", "MaxUnitID", "PageArea", "PhysicalBox", "ApplicationBox", "ContentBox", "BleedBox", "PublicRes", "DocumentRes", "TemplatePage", "DefaultCS", "Pages", "Page", "Area", "Content", "Layer", "PageBlock", "Outlines", "OutlineElem", "Bookmarks", "Bookmark", "Actions", "Action", "Permissions", "VPreferences", "Annotations", "PageAnnot", "Annot", "Appearance", "CustomTags", "CustomTag", "Extensions", "Extension",
        "Attachments", "Attachment", "FileLoc", "Res", "Fonts", "Font", "FontFile", "MultiMedias", "MultiMedia", "MediaFile", "ColorSpaces", "ColorSpace", "DrawParams", "DrawParam", "CompositeGraphicUnits", "CompositeGraphicUnit", "TextObject", "TextCode", "ImageObject", "PathObject", "CompositeObject", "AbbreviatedData", "FillColor", "StrokeColor", "Clips", "Clip", "Area", "CGTransform", "Glyphs", "Substitution", "Segment", "Template", "ImageMask"
    };
    private static readonly HashSet<string> KnownAttributes = new(StringComparer.Ordinal)
    {
        "Version", "DocType", "ID", "Name", "BaseLoc", "FileLoc", "Format", "MediaType", "External", "CreationDate", "ModDate", "Size", "Visible", "Usage", "Type", "Boundary", "CTM", "ResourceID", "Font", "Size", "Weight", "Italic", "Stroke", "Fill", "HScale", "DrawParam", "Value", "Index", "ColorSpace", "Alpha", "LineWidth", "MiterLimit", "DashPattern", "DashOffset", "LineCap", "LineJoin", "X", "Y", "W", "H", "Width", "Height", "CharCode", "CodePosition", "CodeCount", "GlyphCount", "DeltaX", "DeltaY", "Direction", "ZOrder", "PageRef", "PageID", "FontName", "FamilyName", "Charset", "Bold", "URI", "RefID", "Clipped", "LayerID", "PageIndex", "ImageMask", "BitsPerPixel", "Color"
    };

    internal static void Prune(OfdDocumentPackage package, IDictionary<string, byte[]> entries, OfdPackageWriteResult result,
        OfdModelMaterialization material, CancellationToken token, OfdOwnedPruneDispositions? dispositions)
    {
        if (package.PhysicalAttachmentRemovals.Count == 0) return;
        var source = package.AttachmentSource ?? throw new NotSupportedException("Physical attachment removal has no authenticated reader source.");
        source.Validate(package, material, token);
        if (!package.PreservedEntries.TryGetValue("OFD.xml", out var rootBytes))
            throw new NotSupportedException("Physical attachment removal requires an authenticated OFD root.");
        var originalRoot = material.ReadXml("OFD.xml", rootBytes, token).Root!;
        if (originalRoot.Elements(originalRoot.Name.Namespace + "DocBody").Take(2).Count() != 1 ||
            originalRoot.Elements().Any(node => node.Name != originalRoot.Name.Namespace + "DocBody"))
            throw new NotSupportedException("Physical attachment removal cannot prove multiple or opaque document ownership.");
        material.ChargeCopy(512L + 256L * package.PhysicalAttachmentRemovals.Count, token);
        var requests = new HashSet<int>();
        var payloads = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var authenticatedOriginalPayloads = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var request in package.PhysicalAttachmentRemovals)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(128, token);
            if (!requests.Add(request.Ordinal) || !source.TryOriginal(request.Ordinal, out var attachment, out var original) ||
                !ReferenceEquals(attachment.SourceOwner, package) || package.Attachments.Contains(attachment) || original.External ||
                original.PayloadPath is null || !string.Equals(original.PayloadPath, request.PayloadPath, StringComparison.OrdinalIgnoreCase) ||
                !package.PreservedEntries.TryGetValue(request.PayloadPath, out var originalBytes) ||
                !entries.TryGetValue(request.PayloadPath, out var plannedBytes) || !ReferenceEquals(originalBytes, plannedBytes) ||
                string.Equals(request.PayloadPath, source.DescriptorPath, StringComparison.OrdinalIgnoreCase) ||
                OfdXmlContentProbe.LooksLikeXml(originalBytes, token, count => material.ChargeWork(count, token)))
                throw new NotSupportedException("Physical attachment removal lost its reader-owned payload or detached slot.");
            material.ChargeCopy(128L + 4L * request.PayloadPath.Length, token);
            payloads.Add(request.PayloadPath);
        }
        foreach (var slot in source.Slots.Slots)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
            if (slot.Value is not OfdAttachment slotAttachment || !source.TryOriginal(slotAttachment, out var original))
                throw new InvalidDataException("Attachment source slot lacks an authenticated attachment object.");
            if (original.External || original.PayloadPath is null) continue;
            authenticatedOriginalPayloads.Add(original.PayloadPath);
            if (!payloads.Contains(original.PayloadPath)) continue;
            if (!requests.Contains(original.Ordinal))
                throw new NotSupportedException("An attachment payload is shared with an attachment outside the physical request.");
        }
        // A payload cannot also be an OFD control/resource table, even when an
        // implicit PageRes location has no literal reference in XML.
        foreach (var path in payloads)
        {
            var name = path.Substring(path.LastIndexOf('/') + 1);
            if (string.Equals(path, "OFD.xml", StringComparison.OrdinalIgnoreCase) ||
                ControlNames.Contains(name))
                throw new NotSupportedException("Attachment payload aliases an OFD control or resource table.");
        }
        var deleteDescriptor = package.Attachments.Count == 0;
        if (deleteDescriptor)
        {
            if (requests.Count != source.OriginalCount ||
                !package.PreservedEntries.TryGetValue(source.DescriptorPath, out var descriptorBytes) ||
                !entries.TryGetValue(source.DescriptorPath, out var plannedDescriptor) ||
                !ReferenceEquals(descriptorBytes, plannedDescriptor))
                throw new NotSupportedException("Physical removal cannot discard an attachment descriptor with other owned slots.");
            AssertEmptyableDescriptor(package, source, material, token);
        }
        // Original references catch signature, annotation and resource aliases
        // that an earlier prune pass might have removed from the final plan.
        AssertNoReferences(package.PreservedEntries, payloads, source.DescriptorPath, authenticatedOriginalPayloads, entries, material, token);
        var candidates = new HashSet<string>(payloads, StringComparer.OrdinalIgnoreCase);
        if (deleteDescriptor) candidates.Add(source.DescriptorPath);
        AssertNoReferences(entries, candidates, deleteDescriptor ? source.DescriptorPath : null, null, null, material, token);
        foreach (var path in candidates)
        {
            token.ThrowIfCancellationRequested();
            if (!entries.TryGetValue(path, out var bytes)) throw new InvalidDataException("Authenticated attachment removal candidate disappeared.");
            dispositions?.RecordRemoval(path, bytes, path == source.DescriptorPath ? OfdOwnedPruneReason.RemovedAttachmentDescriptor : OfdOwnedPruneReason.RemovedAttachmentPayload, token);
            if (!entries.Remove(path)) throw new InvalidDataException("Authenticated attachment removal failed.");
            result.Removed.Add(path);
        }
    }

    private static void AssertEmptyableDescriptor(OfdDocumentPackage package, OfdAttachmentSource source,
        OfdModelMaterialization material, CancellationToken token)
    {
        var root = material.ReadXml(source.DescriptorPath, package.PreservedEntries[source.DescriptorPath], token).Root!;
        var declaration = OfdXmlEnvelope.Parse(source.Declaration);
        if (root.Name != source.RootName || root.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
            root.Nodes().Any(n => n is XElement element ? element.Name != root.Name.Namespace + "Attachment" : n is not XText text || !string.IsNullOrWhiteSpace(text.Value)) ||
            declaration.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
            declaration.Nodes().Any(n => n is not XText))
            throw new NotSupportedException("Physical removal cannot discard opaque attachment descriptor or declaration content.");
    }

    private static void AssertNoReferences(IEnumerable<KeyValuePair<string, byte[]>> entries, ISet<string> candidates,
        string? excludedDescriptor, ISet<string>? authenticatedPayloads, IDictionary<string, byte[]>? finalEntries,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(512L + 128L * (authenticatedPayloads?.Count ?? 0), token);
        var knownBinary = new HashSet<string>((IEnumerable<string>?)authenticatedPayloads ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var opaqueBinary = new List<string>();
        foreach (var entry in entries)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
            if (candidates.Contains(entry.Key) || string.Equals(entry.Key, excludedDescriptor, StringComparison.OrdinalIgnoreCase)) continue;
            var looksXml = OfdXmlContentProbe.LooksLikeXml(entry.Value, token, count => material.ChargeWork(count, token));
            if (!looksXml)
            {
                if (entry.Key.EndsWith(".xml", StringComparison.OrdinalIgnoreCase))
                    throw new NotSupportedException("Physical attachment removal cannot inspect an XML control entry: " + entry.Key);
                material.ChargeCopy(128L + 4L * entry.Key.Length, token);
                opaqueBinary.Add(entry.Key);
                continue;
            }
            XDocument xml;
            try { xml = material.ReadXml(entry.Key, entry.Value, token); }
            catch (XmlException exception) { throw new NotSupportedException("Physical attachment removal cannot inspect XML-looking content: " + entry.Key, exception); }
            if (xml.Root is null) throw new NotSupportedException("Physical attachment removal cannot inspect empty XML content.");
            var root = xml.Root;
            if (root.Name.NamespaceName != OfdConstants.Namespace && root.Name.NamespaceName != OfdConstants.StandardNamespace)
                throw new NotSupportedException("Physical attachment removal cannot prove an extension XML reference closure: " + entry.Key);
            foreach (var element in root.DescendantsAndSelf())
            {
                token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
                if (element.Name.Namespace != root.Name.Namespace || element.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                    a.Name.Namespace != XNamespace.None && !(a.Name.Namespace == XNamespace.Xml &&
                        (a.Name.LocalName == "space" || a.Name.LocalName == "lang"))))
                    throw new NotSupportedException("Physical attachment removal cannot prove an extension XML reference closure: " + entry.Key);
                if (!KnownElements.Contains(element.Name.LocalName) || element.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                    a.Name.Namespace == XNamespace.None && !KnownAttributes.Contains(a.Name.LocalName)))
                    throw new NotSupportedException("Physical attachment removal cannot prove an unknown OFD field reference closure: " + entry.Key);
                foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
                    CheckScalar(entry.Key, element, attribute.Value, candidates, knownBinary,
                        attribute.Name.LocalName is "FontFile" or "MediaFile" ||
                        (attribute.Name.LocalName == "FileLoc" && element.Name.LocalName == "Attachment"), material, token);
                if (!element.HasElements && !string.IsNullOrWhiteSpace(element.Value))
                    CheckScalar(entry.Key, element, element.Value, candidates, knownBinary,
                        element.Name.LocalName is "FontFile" or "MediaFile" ||
                        (element.Name.LocalName == "FileLoc" && element.Parent?.Name.LocalName == "Attachment"), material, token);
            }
        }
        foreach (var path in opaqueBinary)
            if (!knownBinary.Contains(path) && (finalEntries is null || finalEntries.ContainsKey(path)))
                throw new NotSupportedException("Physical attachment removal cannot prove opaque binary entry ownership: " + path);
    }

    private static void CheckScalar(string entry, XElement element, string value, ISet<string> candidates,
        ISet<string> knownBinary, bool payloadSlot, OfdModelMaterialization material, CancellationToken token)
    {
        // Scan every scalar conservatively, then interpret all ancestor BaseLoc
        // scopes explicitly. Res/@BaseLoc + FontFile/MediaFile and nested
        // annotation/signature paths are covered even without a literal full path.
        if (string.IsNullOrWhiteSpace(value)) return;
        if (payloadSlot && value.IndexOf('%') >= 0)
            throw new NotSupportedException("Physical attachment removal cannot prove an encoded payload reference.");
        material.ChargeCopy(256L + 4L * entry.Length, token);
        var bases = new List<string> { entry };
        var basis = entry;
        foreach (var owner in element.AncestorsAndSelf().Reverse())
        {
            token.ThrowIfCancellationRequested();
            var location = owner.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(location)) continue;
            try
            {
                basis = material.ResolvePath(basis, location!.TrimEnd('/') + "/_attachment_reference", token);
                material.ChargeCopy(128L + 4L * basis.Length, token); bases.Add(basis);
            }
            catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception))
            { throw new NotSupportedException("Physical attachment removal cannot resolve an ancestor BaseLoc.", exception); }
        }
        foreach (var candidateBasis in bases)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                var target = material.ResolvePath(candidateBasis, value, token);
                if (candidates.Contains(target))
                    throw new NotSupportedException("Attachment payload or descriptor remains referenced from " + entry);
                if (payloadSlot) knownBinary.Add(target);
            }
            catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception))
            { /* Non-path scalar in a closed OFD XML field. */ }
        }
    }
}
