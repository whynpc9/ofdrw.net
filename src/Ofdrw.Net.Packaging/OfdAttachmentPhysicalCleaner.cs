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
        "OFD", "DocBody", "DocInfo", "DocID", "Title", "Author", "Subject", "Keywords", "Creator", "CreationDate", "ModDate", "DocRoot", "Versions", "Version", "Signatures", "MaxSignId", "Signature", "SignedInfo", "SignedValue", "References", "Reference", "FileRef", "CheckValue", "DigestMethod", "DigestValue", "Provider", "SignatureMethod", "SignatureDateTime", "StampAnnot", "Seal", "Parameters", "Parameter",
        "Document", "CommonData", "MaxUnitID", "PageArea", "PhysicalBox", "ApplicationBox", "ContentBox", "BleedBox", "PublicRes", "DocumentRes", "TemplatePage", "DefaultCS", "Pages", "Page", "Area", "Content", "Layer", "PageBlock", "Outlines", "OutlineElem", "Bookmarks", "Bookmark", "Dest", "Actions", "Action", "Goto", "Permissions", "VPreferences", "Annotations", "PageAnnot", "Annot", "Appearance", "Remark", "CustomTags", "CustomTag", "EMRTags", "Tag", "Extensions", "Extension",
        "Attachments", "Attachment", "FileLoc", "Res", "Fonts", "Font", "FontFile", "MultiMedias", "MultiMedia", "MediaFile", "ColorSpaces", "ColorSpace", "DrawParams", "DrawParam", "CompositeGraphicUnits", "CompositeGraphicUnit", "TextObject", "TextCode", "ImageObject", "PathObject", "Path", "CompositeObject", "AbbreviatedData", "FillColor", "StrokeColor", "Clips", "Clip", "Area", "CGTransform", "Glyphs", "Substitution", "Segment", "Template", "ImageMask"
    };
    private static readonly HashSet<string> KnownAttributes = new(StringComparer.Ordinal)
    {
        "Version", "DocType", "ID", "Name", "Title", "Expanded", "Count", "Event", "Left", "Top", "Right", "Bottom", "Zoom", "TypeID", "NameSpace", "Key", "BaseLoc", "FileLoc", "FileRef", "CheckMethod", "ProviderName", "Company", "Relative", "Format", "MediaType", "External", "CreationDate", "ModDate", "LastModDate", "Creator", "Subtype", "Print", "NoZoom", "NoRotate", "ReadOnly", "Size", "Visible", "Usage", "Type", "Boundary", "CTM", "ResourceID", "Font", "Size", "Weight", "Italic", "Stroke", "Fill", "HScale", "ReadDirection", "CharDirection", "DrawParam", "Value", "Index", "ColorSpace", "Alpha", "LineWidth", "MiterLimit", "DashPattern", "DashOffset", "LineCap", "LineJoin", "Cap", "Join", "Rule", "X", "Y", "W", "H", "Width", "Height", "CharCode", "CodePosition", "CodeCount", "GlyphCount", "DeltaX", "DeltaY", "Direction", "ZOrder", "TemplateID", "PageRef", "PageID", "FontName", "FamilyName", "Charset", "Bold", "URI", "RefID", "Clipped", "LayerID", "PageIndex", "ImageMask", "BitsPerPixel", "Color"
    };
    private static readonly HashSet<string> ScalarPathElements = new(StringComparer.Ordinal)
    { "SignedValue", "FileRef", "FileLoc", "FontFile", "MediaFile", "DocRoot", "PublicRes", "DocumentRes", "PageRes" };

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
        AssertNoReferences(package, package.PreservedEntries, payloads, source.DescriptorPath, authenticatedOriginalPayloads, entries, result.SignaturesInvalidated, material, token);
        var candidates = new HashSet<string>(payloads, StringComparer.OrdinalIgnoreCase);
        if (deleteDescriptor) candidates.Add(source.DescriptorPath);
        AssertNoReferences(package, entries, candidates, deleteDescriptor ? source.DescriptorPath : null, null, null, false, material, token);
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

    private static void AssertNoReferences(OfdDocumentPackage package, IDictionary<string, byte[]> entries, ISet<string> candidates,
        string? excludedDescriptor, ISet<string>? authenticatedPayloads, IDictionary<string, byte[]>? finalEntries,
        bool invalidatedSignatures,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(512L + 128L * (authenticatedPayloads?.Count ?? 0), token);
        var knownBinary = new HashSet<string>((IEnumerable<string>?)authenticatedPayloads ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var trustedResources = TrustedResourceTables(package, entries, material, token);
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
                if (element.HasElements && (ScalarPathElements.Contains(element.Name.LocalName) ||
                    element.Nodes().OfType<XText>().Any(text => !string.IsNullOrWhiteSpace(text.Value))))
                    throw new NotSupportedException("Physical attachment removal cannot prove mixed XML path content: " + entry.Key);
                foreach (var attribute in element.Attributes().Where(a => !a.IsNamespaceDeclaration))
                {
                    if (invalidatedSignatures && finalEntries is not null && !finalEntries.ContainsKey(entry.Key) &&
                        root.Name.LocalName == "Signature" && element.Name == root.Name.Namespace + "Reference" &&
                        element.Parent?.Name == root.Name.Namespace + "References" &&
                        element.Parent.Parent?.Name == root.Name.Namespace + "SignedInfo" && element.Parent.Parent.Parent == root &&
                        attribute.Name.LocalName == "FileRef")
                    {
                        if (ContainsAmbiguousPathSyntax(attribute.Value))
                            throw new NotSupportedException("Invalidated signature has an ambiguous digest reference.");
                        // This standard digest reference has no remaining owner
                        // once the entire signature entry is actually pruned.
                        continue;
                    }
                    CheckScalar(entry.Key, element, attribute.Value, IsPathBearingAttribute(attribute.Name.LocalName), candidates, knownBinary,
                        IsTrustedPayloadSlot(package, entry.Key, element, attribute.Name.LocalName, trustedResources), material, token);
                }
                if (!element.HasElements && !string.IsNullOrWhiteSpace(element.Value))
                    CheckScalar(entry.Key, element, element.Value, IsPathBearingText(element), candidates, knownBinary,
                        IsTrustedPayloadSlot(package, entry.Key, element, null, trustedResources), material, token);
            }
        }
        foreach (var path in opaqueBinary)
            if (!knownBinary.Contains(path) && (finalEntries is null || finalEntries.ContainsKey(path)))
                throw new NotSupportedException("Physical attachment removal cannot prove opaque binary entry ownership: " + path);
    }

    private static bool IsPathBearingAttribute(string name) => name is "BaseLoc" or "FileLoc" or "FileRef" or "MediaFile" or "URI";

    private static bool IsPathBearingText(XElement element) => ScalarPathElements.Contains(element.Name.LocalName) ||
        (element.Parent?.Name.LocalName == "DocBody" && element.Name.LocalName == "Signatures") ||
        (element.Parent?.Name.LocalName == "Document" &&
            element.Name.LocalName is ("Attachments" or "Annotations" or "CustomTags"));

    private static void CheckScalar(string entry, XElement element, string value, bool pathBearing, ISet<string> candidates,
        ISet<string> knownBinary, bool payloadSlot, OfdModelMaterialization material, CancellationToken token)
    {
        // Scan every scalar for literal aliases, but reject encoded/backslash
        // syntax only in fields that can carry paths. Ordinary TextCode and
        // metadata values may contain those characters. Resolve every ancestor
        // BaseLoc explicitly for resource and nested annotation/signature paths.
        if (string.IsNullOrWhiteSpace(value)) return;
        if (pathBearing && ContainsAmbiguousPathSyntax(value))
            throw new NotSupportedException("Physical attachment removal cannot prove an encoded or backslash XML reference.");
        material.ChargeCopy(256L + 4L * entry.Length, token);
        var bases = new List<string> { entry };
        var basis = entry;
        foreach (var owner in element.AncestorsAndSelf().Reverse())
        {
            token.ThrowIfCancellationRequested();
            var location = owner.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(location)) continue;
            if (ContainsAmbiguousPathSyntax(location!))
                throw new NotSupportedException("Physical attachment removal cannot prove an encoded or backslash BaseLoc.");
            try
            {
                basis = material.ResolvePath(basis, location!.TrimEnd('/') + "/_attachment_reference", token);
                material.ChargeCopy(128L + 4L * basis.Length, token); bases.Add(basis);
            }
            catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception))
            { throw new NotSupportedException("Physical attachment removal cannot resolve an ancestor BaseLoc.", exception); }
        }
        for (var i = 0; i < bases.Count; i++)
        {
            token.ThrowIfCancellationRequested();
            try
            {
                var target = material.ResolvePath(bases[i], value, token);
                if (candidates.Contains(target))
                    throw new NotSupportedException("Attachment payload or descriptor remains referenced from " + entry);
                // All scopes are checked for aliases, but only the innermost
                // BaseLoc gives a standard resource payload its owned path.
                if (payloadSlot && i == bases.Count - 1) knownBinary.Add(target);
            }
            catch (InvalidDataException exception) when (!OfdTraversalBudget.IsLimit(exception))
            { /* Non-path scalar in a closed OFD XML field. */ }
        }
    }

    private static HashSet<string> TrustedResourceTables(OfdDocumentPackage package, IDictionary<string, byte[]> entries,
        OfdModelMaterialization material, CancellationToken token)
    {
        var paths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var documentPath = package.DocumentEntryPath ?? throw new NotSupportedException("Attachment document path is missing.");
        if (!entries.TryGetValue(documentPath, out var bytes))
            throw new NotSupportedException("Attachment document entry is missing.");
        var root = material.ReadXml(documentPath, bytes, token).Root!;
        if (root.Name.LocalName != "Document") throw new NotSupportedException("Attachment document entry is not a Document.");
        var common = root.Element(root.Name.Namespace + "CommonData");
        foreach (var name in new[] { "PublicRes", "DocumentRes" })
        {
            var value = common?.Element(root.Name.Namespace + name)?.Value;
            if (string.IsNullOrWhiteSpace(value)) continue;
            if (ContainsAmbiguousPathSyntax(value!)) throw new NotSupportedException("Resource table location is ambiguous.");
            paths.Add(material.ResolvePath(documentPath, value!, token));
        }
        var pageLocations = (root.Element(root.Name.Namespace + "Pages")?.Elements(root.Name.Namespace + "Page") ?? Enumerable.Empty<XElement>())
            .Concat(common?.Elements(root.Name.Namespace + "TemplatePage") ?? Enumerable.Empty<XElement>());
        foreach (var page in pageLocations)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
            var value = page.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(value)) continue;
            if (ContainsAmbiguousPathSyntax(value!)) throw new NotSupportedException("Page location is ambiguous.");
            var pagePath = material.ResolvePath(documentPath, value!, token);
            paths.Add(pagePath.Substring(0, pagePath.LastIndexOf('/') + 1) + "PageRes.xml");
        }
        return paths;
    }

    private static bool IsTrustedPayloadSlot(OfdDocumentPackage package, string path, XElement element,
        string? attributeName, ISet<string> trustedResources)
    {
        var root = element.Document?.Root;
        if (root is null) return false;
        var ns = root.Name.Namespace;
        if (root.Name.LocalName == "Res" && trustedResources.Contains(path) && attributeName is null)
            return (element.Name == ns + "FontFile" && element.Parent?.Name == ns + "Font" &&
                    element.Parent.Parent?.Name == ns + "Fonts" && element.Parent.Parent.Parent == root) ||
                   (element.Name == ns + "MediaFile" && element.Parent?.Name == ns + "MultiMedia" &&
                    element.Parent.Parent?.Name == ns + "MultiMedias" && element.Parent.Parent.Parent == root);
        if (package.AttachmentSource?.DescriptorPath != path || root.Name.LocalName != "Attachments") return false;
        return attributeName == "FileLoc" && element.Name == ns + "Attachment" && element.Parent == root ||
               attributeName is null && element.Name == ns + "FileLoc" && element.Parent?.Name == ns + "Attachment" &&
               element.Parent.Parent == root;
    }

    private static bool ContainsAmbiguousPathSyntax(string value)
    {
        if (value.IndexOf('\\') >= 0) return true;
        for (var i = 0; i + 2 < value.Length; i++)
            if (value[i] == '%' && Uri.IsHexDigit(value[i + 1]) && Uri.IsHexDigit(value[i + 2])) return true;
        return false;
    }
}
