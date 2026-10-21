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
    private enum PayloadOwnerScope { None, Entry, ResourceRoot }
    private static readonly HashSet<string> ControlNames = new(StringComparer.OrdinalIgnoreCase)
    { "Document.xml", "PublicRes.xml", "DocumentRes.xml", "PageRes.xml", "Attachments.xml", "Annotations.xml", "Signatures.xml", "Signature.xml" };
    // A closed set for this cleanup profile. Unknown standard-looking names may
    // carry implicit references, so they cannot be admitted by string matching.
    private static readonly HashSet<string> KnownElements = new(StringComparer.Ordinal)
    {
        "OFD", "DocBody", "DocInfo", "DocID", "Title", "Author", "Subject", "Abstract", "Keywords", "Keyword", "Creator", "CreatorVersion", "CreationDate", "ModDate", "DocUsage", "Cover", "CustomDatas", "CustomData", "DocRoot", "Versions", "Version", "DocVersion", "FileList", "File", "Signatures", "MaxSignId", "Signature", "SignedInfo", "SignedValue", "References", "Reference", "FileRef", "CheckValue", "DigestMethod", "DigestValue", "Provider", "SignatureMethod", "SignatureDateTime", "StampAnnot", "Seal", "BaseLoc", "Parameters", "Parameter",
        "Document", "CommonData", "MaxUnitID", "PageArea", "PhysicalBox", "ApplicationBox", "ContentBox", "BleedBox", "PublicRes", "DocumentRes", "TemplatePage", "DefaultCS", "Pages", "Page", "Area", "Content", "Layer", "PageBlock", "Outlines", "OutlineElem", "Bookmarks", "Bookmark", "Dest", "Actions", "Action", "Region", "Move", "Line", "QuadraticBezier", "CubicBezier", "Arc", "Close", "Goto", "GotoA", "URI", "Sound", "Movie", "Permissions", "Edit", "Export", "Signature", "Watermark", "PrintScreen", "Print", "ValidPeriod", "VPreferences", "PageMode", "PageLayout", "TabDisplay", "HideToolbar", "HideMenubar", "HideWindowUI", "ZoomMode", "Zoom", "Annotations", "PageAnnot", "Annot", "Appearance", "Remark", "CustomTags", "CustomTag", "EMRTags", "Tag", "Extensions", "Extension",
        "Attachments", "Attachment", "FileLoc", "Res", "Fonts", "Font", "FontFile", "MultiMedias", "MultiMedia", "MediaFile", "ColorSpaces", "ColorSpace", "DrawParams", "DrawParam", "CompositeGraphicUnits", "CompositeGraphicUnit", "TextObject", "TextCode", "ImageObject", "PathObject", "Path", "CompositeObject", "AbbreviatedData", "FillColor", "StrokeColor", "Clips", "Clip", "Area", "CGTransform", "Glyphs", "Substitution", "Segment", "Template", "ImageMask"
    };
    private static readonly HashSet<string> KnownAttributes = new(StringComparer.Ordinal)
    {
        "Version", "DocType", "ID", "Name", "Title", "Expanded", "Count", "Event", "Left", "Top", "Right", "Bottom", "Zoom", "TypeID", "NameSpace", "Key", "BaseLoc", "FileLoc", "FileRef", "MediaFile", "CheckMethod", "ProviderName", "Company", "Relative", "Format", "MediaType", "External", "CreationDate", "ModDate", "LastModDate", "Creator", "Subtype", "Print", "Printable", "Copies", "StartDate", "EndDate", "NoZoom", "NoRotate", "ReadOnly", "Size", "Visible", "Usage", "Type", "Boundary", "CTM", "ResourceID", "Substitution", "Font", "Size", "Weight", "Italic", "Stroke", "Fill", "HScale", "ReadDirection", "CharDirection", "DrawParam", "Value", "Index", "ColorSpace", "Alpha", "LineWidth", "MiterLimit", "DashPattern", "DashOffset", "LineCap", "LineJoin", "Cap", "Join", "Rule", "X", "Y", "W", "H", "Width", "Height", "CharCode", "CodePosition", "CodeCount", "GlyphCount", "DeltaX", "DeltaY", "Direction", "ZOrder", "Start", "TemplateID", "PageRef", "PageID", "FontName", "FamilyName", "Charset", "Bold", "Serif", "FixedWidth", "Current", "URI", "Base", "Target", "AttachID", "NewWindow", "Volume", "Repeat", "Synchronous", "Operator", "Point1", "Point2", "Point3", "SweepDirection", "LargeArc", "RotationAngle", "EllipseSize", "EndPoint", "RefID", "Clipped", "LayerID", "PageIndex", "ImageMask", "BitsPerPixel", "Color"
    };
    private static readonly HashSet<string> ScalarPathElements = new(StringComparer.Ordinal)
    { "SignedValue", "FileRef", "FileLoc", "FontFile", "MediaFile", "DocRoot", "PublicRes", "DocumentRes", "PageRes", "File", "BaseLoc", "Cover" };

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
        var removedAttachmentIds = new HashSet<string>(StringComparer.Ordinal);
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
            if (original.Values.TryGetValue("ID", out var originalId) && !string.IsNullOrWhiteSpace(originalId))
                removedAttachmentIds.Add(originalId!.Trim());
            if (!string.IsNullOrWhiteSpace(attachment.Id))
                removedAttachmentIds.Add(attachment.Id!.Trim());
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
        AssertNoReferences(package, package.PreservedEntries, payloads, removedAttachmentIds, source.DescriptorPath, authenticatedOriginalPayloads, entries, result.SignaturesInvalidated, material, token);
        var candidates = new HashSet<string>(payloads, StringComparer.OrdinalIgnoreCase);
        if (deleteDescriptor) candidates.Add(source.DescriptorPath);
        AssertNoReferences(package, entries, candidates, removedAttachmentIds, deleteDescriptor ? source.DescriptorPath : null,
            authenticatedOriginalPayloads, null, false, material, token);
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
        ISet<string> removedAttachmentIds,
        string? excludedDescriptor, ISet<string>? authenticatedPayloads, IDictionary<string, byte[]>? finalEntries,
        bool invalidatedSignatures,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(512L + 128L * (authenticatedPayloads?.Count ?? 0), token);
        var knownBinary = new HashSet<string>((IEnumerable<string>?)authenticatedPayloads ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);
        var controlReferences = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var trustedResources = TrustedResourceTables(package, entries, material, token);
        var opaqueBinary = new List<string>();
        foreach (var entry in entries)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
            if (candidates.Contains(entry.Key) || string.Equals(entry.Key, excludedDescriptor, StringComparison.OrdinalIgnoreCase)) continue;
            var looksXml = OfdXmlContentProbe.LooksLikeXml(entry.Value, token, count => material.ChargeWork(count, token));
            if (!looksXml)
            {
                material.ChargeCopy(128L + 4L * entry.Key.Length, token);
                opaqueBinary.Add(entry.Key);
                continue;
            }
            XDocument xml;
            try { xml = material.ReadXml(entry.Key, entry.Value, token); }
            catch (XmlException exception) { throw new NotSupportedException("Physical attachment removal cannot inspect XML-looking content: " + entry.Key, exception); }
            if (xml.Root is null) throw new NotSupportedException("Physical attachment removal cannot inspect empty XML content.");
            if (xml.DescendantNodes().OfType<XProcessingInstruction>().Any())
                throw new NotSupportedException("Physical attachment removal cannot prove a processing instruction reference closure: " + entry.Key);
            var root = xml.Root;
            if (root.Name.NamespaceName != OfdConstants.Namespace && root.Name.NamespaceName != OfdConstants.StandardNamespace)
                throw new NotSupportedException("Physical attachment removal cannot prove an extension XML reference closure: " + entry.Key);
            if (root.Name.LocalName == "DocVersion" && !IsStandardDocVersion(root))
                throw new NotSupportedException("Physical attachment removal cannot prove a malformed DocVersion reference closure: " + entry.Key);
            foreach (var element in root.DescendantsAndSelf())
            {
                token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
                if (element.Name.Namespace != root.Name.Namespace || element.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                    a.Name.Namespace != XNamespace.None && !(a.Name.Namespace == XNamespace.Xml &&
                        (a.Name.LocalName == "space" || a.Name.LocalName == "lang")) &&
                    !IsWriterOwnedHint(element, a)))
                    throw new NotSupportedException("Physical attachment removal cannot prove an extension XML reference closure: " + entry.Key);
                if (!KnownElements.Contains(element.Name.LocalName) || element.Attributes().Any(a => !a.IsNamespaceDeclaration &&
                    a.Name.Namespace == XNamespace.None && !KnownAttributes.Contains(a.Name.LocalName)))
                    throw new NotSupportedException("Physical attachment removal cannot prove an unknown OFD field reference closure: " + entry.Key);
                if (element.Name.LocalName == "Region" && !IsStandardActionRegion(element))
                    throw new NotSupportedException("Physical attachment removal cannot prove an action Region reference closure: " + entry.Key);
                // Actions can be deleted or retargeted by the typed writer.
                // Only the output plan can prove a dangling action remains.
                if (finalEntries is null && element.Name.LocalName == "GotoA" && element.Attribute("AttachID") is XAttribute targetId &&
                    removedAttachmentIds.Contains(targetId.Value.Trim()))
                    throw new NotSupportedException("A retained action targets the removed attachment ID: " + entry.Key);
                if (element.Parent?.Name.LocalName == "DocBody" && element.Name.LocalName == "Versions" &&
                    (element.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
                     !element.HasElements ||
                     element.Elements().Any(child => child.Name != element.Name.Namespace + "Version" ||
                         child.HasElements || !string.IsNullOrWhiteSpace(child.Value) ||
                         string.IsNullOrWhiteSpace(child.Attribute("ID")?.Value) ||
                         string.IsNullOrWhiteSpace(child.Attribute("Index")?.Value) ||
                         string.IsNullOrWhiteSpace(child.Attribute("BaseLoc")?.Value) ||
                         child.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.LocalName is not ("ID" or "Index" or "Current" or "BaseLoc")))))
                    throw new NotSupportedException("Physical attachment removal cannot prove a malformed Versions wrapper: " + entry.Key);
                if (element.HasElements && (IsPathBearingText(element) ||
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
                    CheckScalar(entry.Key, element, attribute.Value, IsPathBearingAttribute(element, attribute.Name.LocalName),
                        IsControlReference(element, attribute.Name.LocalName), candidates, knownBinary, controlReferences,
                        IsTrustedPayloadSlot(package, entry.Key, element, attribute.Name.LocalName, trustedResources), material, token);
                }
                if (!element.HasElements && !string.IsNullOrWhiteSpace(element.Value))
                    CheckScalar(entry.Key, element, element.Value, IsPathBearingText(element), IsControlReference(element, null),
                        candidates, knownBinary, controlReferences,
                        IsTrustedPayloadSlot(package, entry.Key, element, null, trustedResources), material, token);
            }
        }
        foreach (var path in opaqueBinary)
            if ((controlReferences.Contains(path) || !knownBinary.Contains(path)) &&
                (finalEntries is null || finalEntries.ContainsKey(path)))
                throw new NotSupportedException(path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                    ? "Physical attachment removal cannot inspect an XML control entry: " + path
                    : "Physical attachment removal cannot prove opaque binary entry ownership: " + path);
    }

    private static bool IsWriterOwnedHint(XElement element, XAttribute attribute) =>
        element.Name.LocalName == "TextObject" && attribute.Name == OfdTextEmphasis.FauxItalicFactor ||
        element.Name.LocalName == "ImageObject" && attribute.Name == OfdImageRenderingHints.PdfInterpolateV1 &&
        OfdImageRenderingHints.IsCanonicalValue(attribute.Value);

    private static bool IsStandardActionRegion(XElement region)
    {
        var ns = region.Name.Namespace;
        if (region.Parent?.Name != ns + "Action" ||
            !ReferenceEquals(region.Parent.Elements().FirstOrDefault(), region) ||
            region.Parent.Elements(ns + "Region").Count() != 1 ||
            region.Attributes().Any(a => !a.IsNamespaceDeclaration) ||
            region.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value)))
            return false;
        var areas = region.Elements().ToArray();
        if (areas.Length == 0) return false;
        foreach (var area in areas)
        {
            if (area.Name != ns + "Area" || string.IsNullOrWhiteSpace(area.Attribute("Start")?.Value) ||
                area.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name != "Start") ||
                area.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))) return false;
            var commands = area.Elements().ToArray();
            if (commands.Length == 0) return false;
            foreach (var command in commands)
            {
                if (command.HasElements || command.Nodes().OfType<XText>().Any(t => !string.IsNullOrWhiteSpace(t.Value))) return false;
                var names = command.Name == ns + "Move" || command.Name == ns + "Line" ? new[] { "Point1" } :
                    command.Name == ns + "QuadraticBezier" ? new[] { "Point1", "Point2" } :
                    command.Name == ns + "CubicBezier" ? new[] { "Point1", "Point2", "Point3" } :
                    command.Name == ns + "Arc" ? new[] { "SweepDirection", "LargeArc", "RotationAngle", "EllipseSize", "EndPoint" } :
                    command.Name == ns + "Close" ? Array.Empty<string>() : null;
                if (names is null || command.Attributes().Any(a => !a.IsNamespaceDeclaration && !names.Contains(a.Name.LocalName))) return false;
                if (command.Name == ns + "Move" || command.Name == ns + "Line" || command.Name == ns + "QuadraticBezier" || command.Name == ns + "Arc")
                    if (names.Any(name => string.IsNullOrWhiteSpace(command.Attribute(name)?.Value))) return false;
                if (command.Name == ns + "CubicBezier" && string.IsNullOrWhiteSpace(command.Attribute("Point3")?.Value)) return false;
            }
        }
        return true;
    }

    private static bool IsStandardDocVersion(XElement root)
    {
        var ns = root.Name.Namespace;
        var children = root.Elements().ToArray();
        if (children.Length != 2 || children[0].Name != ns + "FileList" || children[1].Name != ns + "DocRoot" ||
            string.IsNullOrWhiteSpace(root.Attribute("ID")?.Value) ||
            root.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.LocalName is not ("ID" or "Version" or "Name" or "CreationDate")))
            return false;
        var files = children[0].Elements().ToArray();
        return children[0].Attributes().All(a => a.IsNamespaceDeclaration) &&
               children[1].Attributes().All(a => a.IsNamespaceDeclaration) &&
               files.Length > 0 && files.All(file => file.Name == ns + "File" && !file.HasElements &&
                   !string.IsNullOrWhiteSpace(file.Value) && !string.IsNullOrWhiteSpace(file.Attribute("ID")?.Value) &&
                   file.Attributes().All(a => a.IsNamespaceDeclaration || a.Name.LocalName == "ID")) &&
               !children[1].HasElements && !string.IsNullOrWhiteSpace(children[1].Value);
    }

    private static bool IsPathBearingAttribute(XElement element, string name) =>
        name is "BaseLoc" or "FileRef" or "MediaFile" ||
        name == "FileLoc" && !IsExternalAttachment(element) ||
        name == "URI" && !(element.Name.LocalName == "URI" && element.Parent?.Name.LocalName == "Action");

    private static bool IsPathBearingText(XElement element) =>
        ScalarPathElements.Contains(element.Name.LocalName) &&
        !(element.Name.LocalName == "FileLoc" && IsExternalAttachment(element.Parent)) ||
        (element.Parent?.Name.LocalName == "DocBody" && element.Name.LocalName == "Signatures") ||
        (element.Document?.Root?.Name.LocalName == "Document" &&
            element.Name.LocalName is ("Attachments" or "Annotations" or "CustomTags" or "Extensions"));

    private static bool IsExternalAttachment(XElement? element) =>
        element?.Name.LocalName == "Attachment" &&
        bool.TryParse(element.Attribute("External")?.Value, out var external) && external;

    private static bool IsControlReference(XElement element, string? attributeName) =>
        attributeName == "BaseLoc" && element.Name.LocalName is ("Page" or "TemplatePage" or "Version" or "Signature") ||
        attributeName == "FileLoc" && element.Name.LocalName == "PageAnnot" ||
        attributeName is null && element.Name.LocalName is
            ("DocRoot" or "PublicRes" or "DocumentRes" or "PageRes" or "Signatures" or
             "Attachments" or "Annotations" or "CustomTags" or "Extensions") ||
        attributeName is null && element.Name.LocalName == "FileLoc" &&
            element.Parent?.Name == element.Name.Namespace + "Page" &&
            element.Parent.Parent?.Name == element.Name.Namespace + "Annotations" &&
            element.Document?.Root == element.Parent.Parent;

    private static void CheckScalar(string entry, XElement element, string value, bool pathBearing, bool controlReference,
        ISet<string> candidates, ISet<string> knownBinary, ISet<string> controlReferences, PayloadOwnerScope payloadSlot,
        OfdModelMaterialization material, CancellationToken token)
    {
        // Scan every scalar for literal aliases, but reject encoded/backslash
        // syntax only in fields that can carry paths. Ordinary TextCode and
        // metadata values may contain those characters. Resolve every ancestor
        // BaseLoc explicitly for resource and nested annotation/signature paths.
        if (string.IsNullOrWhiteSpace(value)) return;
        if (pathBearing && !string.Equals(value, value.Trim(), StringComparison.Ordinal))
            throw new NotSupportedException("Physical attachment removal cannot prove a path with surrounding whitespace.");
        if (pathBearing && ContainsAmbiguousPathSyntax(value))
            throw new NotSupportedException("Physical attachment removal cannot prove an encoded or backslash XML reference.");
        material.ChargeCopy(256L + 4L * entry.Length, token);
        var bases = new List<string> { entry };
        var resourceRootIndex = 0;
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
                if (ReferenceEquals(owner, element.Document?.Root)) resourceRootIndex = bases.Count - 1;
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
                if (controlReference) controlReferences.Add(target);
                // Every scope is checked for aliases. Only the BaseLoc used by
                // the actual consumer can authenticate an opaque payload.
                if (payloadSlot == PayloadOwnerScope.Entry && i == 0 ||
                    payloadSlot == PayloadOwnerScope.ResourceRoot && i == resourceRootIndex)
                    knownBinary.Add(target);
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
            if (page.Name.LocalName == "TemplatePage" && string.IsNullOrWhiteSpace(page.Attribute("ID")?.Value))
                continue;
            var value = page.Attribute("BaseLoc")?.Value;
            if (string.IsNullOrWhiteSpace(value)) continue;
            if (ContainsAmbiguousPathSyntax(value!)) throw new NotSupportedException("Page location is ambiguous.");
            var pagePath = material.ResolvePath(documentPath, value!, token);
            paths.Add(pagePath.Substring(0, pagePath.LastIndexOf('/') + 1) + "PageRes.xml");
        }
        return paths;
    }

    private static PayloadOwnerScope IsTrustedPayloadSlot(OfdDocumentPackage package, string path, XElement element,
        string? attributeName, ISet<string> trustedResources)
    {
        var root = element.Document?.Root;
        if (root is null) return PayloadOwnerScope.None;
        var ns = root.Name.Namespace;
        if (attributeName is null && string.Equals(path, "OFD.xml", StringComparison.OrdinalIgnoreCase) &&
            root.Name == ns + "OFD" && element.Name == ns + "Cover" &&
            element.Parent?.Name == ns + "DocInfo" && element.Parent.Parent?.Name == ns + "DocBody" &&
            element.Parent.Parent.Parent == root &&
            ReferenceEquals(element.Parent.Parent, root.Element(ns + "DocBody")) &&
            ReferenceEquals(element.Parent, element.Parent.Parent.Element(ns + "DocInfo")) &&
            ReferenceEquals(element, element.Parent.Element(ns + "Cover")))
            return PayloadOwnerScope.Entry;
        if (root.Name.LocalName == "Res" && trustedResources.Contains(path))
        {
            if (attributeName is null && element.Name == ns + "FontFile" && element.Parent?.Name == ns + "Font" &&
                element.Parent.Parent?.Name == ns + "Fonts" && element.Parent.Parent.Parent == root)
            {
                var font = element.Parent;
                return !string.IsNullOrWhiteSpace(font.Attribute("ID")?.Value) &&
                       !string.IsNullOrWhiteSpace(font.Attribute("FontName")?.Value ?? font.Attribute("FamilyName")?.Value) &&
                       ReferenceEquals(font.Element(ns + "FontFile"), element)
                    ? PayloadOwnerScope.ResourceRoot : PayloadOwnerScope.None;
            }
            var media = attributeName == "MediaFile" ? element : element.Parent;
            if (media?.Name == ns + "MultiMedia" && media.Parent?.Name == ns + "MultiMedias" &&
                media.Parent.Parent == root && !string.IsNullOrWhiteSpace(media.Attribute("ID")?.Value))
                return attributeName == "MediaFile" && element == media ||
                       attributeName is null && element.Name == ns + "MediaFile" &&
                       media.Attribute("MediaFile") is null &&
                       ReferenceEquals(media.Element(ns + "MediaFile"), element)
                    ? PayloadOwnerScope.ResourceRoot : PayloadOwnerScope.None;
            return PayloadOwnerScope.None;
        }
        if (package.AttachmentSource?.DescriptorPath != path || root.Name.LocalName != "Attachments") return PayloadOwnerScope.None;
        var attachment = attributeName == "FileLoc" ? element : element.Parent;
        // Reader treats External=true as a navigation target, not an archive
        // payload. Fresh internal attachments are emitted by the writer here.
        if (attachment?.Name != ns + "Attachment" ||
            bool.TryParse(attachment.Attribute("External")?.Value, out var external) && external)
            return PayloadOwnerScope.None;
        return attributeName == "FileLoc" && element == attachment && attachment.Parent == root &&
               attachment.Element(ns + "FileLoc") is null ||
               attributeName is null && element.Name == ns + "FileLoc" && element.Parent == attachment &&
               attachment.Parent == root && ReferenceEquals(attachment.Element(ns + "FileLoc"), element)
            ? PayloadOwnerScope.Entry : PayloadOwnerScope.None;
    }

    private static bool ContainsAmbiguousPathSyntax(string value)
    {
        if (value.IndexOf('\\') >= 0) return true;
        for (var i = 0; i + 2 < value.Length; i++)
            if (value[i] == '%' && Uri.IsHexDigit(value[i + 1]) && Uri.IsHexDigit(value[i + 2])) return true;
        return false;
    }
}
