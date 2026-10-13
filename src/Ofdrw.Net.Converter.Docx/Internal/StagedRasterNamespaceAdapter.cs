using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Converter.Docx.Internal;

// A staged raster document becomes a fresh document when its namespace changes.
// This bridge accepts only a closed image-only source graph; it never discards
// an opaque extension or transfers ownership across expanded-name changes.
internal static class StagedRasterNamespaceAdapter
{
    internal static OfdDocumentPackage Rebuild(OfdDocumentPackage source, CancellationToken token, string? targetNamespace = null)
    {
        var material = new OfdModelMaterialization(new OfdModelLimits());
        material.InspectStorage(source, token); material.GuardPreservedXml(source, token);
        if (source.SourceIdentity is null || source.Options.Namespace is not (OfdConstants.Namespace or OfdConstants.StandardNamespace) ||
            source.Fonts.Count != 0 || source.Attachments.Count != 0 || source.CustomTags.Count != 0 ||
            source.PreservedDocBodyElements.Count != 0 || source.PreservedCommonDataElements.Count != 0 || source.PreservedDocumentElements.Count != 0 ||
            OfdNavigationXml.HasTypedNavigation(source, material, token)) throw Unsupported();
        source.SourceIdentity!.Reconcile(source, material, token);
        var ns = XNamespace.Get(source.Options.Namespace);
        var allowedEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var entryRoles = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var root = Read("OFD.xml", "OFD");
        var bodies = root.Elements(ns + "DocBody").ToArray(); if (bodies.Length != 1) throw Unsupported();
        var creationDate = Date("CreationDate"); var modificationDate = Date("ModDate");
        var documentPath = source.DocumentEntryPath!;
        if (material.ResolvePath("OFD.xml", (bodies[0].Element(ns + "DocRoot") ?? throw Unsupported()).Value, token) != documentPath) throw Unsupported();
        var document = Read(documentPath, "Document");
        var declarations = (document.Element(ns + "Pages") ?? throw Unsupported()).Elements(ns + "Page").ToArray();
        if (declarations.Length != source.Pages.Count) throw Unsupported();
        var media = new Dictionary<string, (string Entry, byte[] Data, string MediaType)>(StringComparer.Ordinal);
        foreach (var declaration in (document.Element(ns + "CommonData") ?? throw Unsupported()).Elements().Where(node => node.Name == ns + "PublicRes" || node.Name == ns + "DocumentRes"))
        {
            var path = material.ResolvePath(documentPath, declaration.Value, token); var res = Read(path, "Res");
            foreach (var image in res.Elements(ns + "MultiMedias").Elements(ns + "MultiMedia"))
            {
                token.ThrowIfCancellationRequested();
                if ((string?)image.Attribute("Type") != "Image") throw Unsupported();
                var formatText = (string?)image.Attribute("Format");
                if (formatText is not null) material.ChargeCopy(4L * formatText.Length, token);
                var format = formatText?.Trim().ToUpperInvariant();
                if (format is not ("PNG" or "JPG" or "JPEG" or "BMP" or "TIFF")) throw Unsupported();
                var file = (image.Element(ns + "MediaFile") ?? throw Unsupported()).Value; var basis = (string?)res.Attribute("BaseLoc");
                var relative = !string.IsNullOrEmpty(basis) && !file.StartsWith("/", StringComparison.Ordinal);
                if (relative) material.ChargeCopy(2L * (2L * basis!.Length + file.Length + 1), token);
                var reference = relative ? basis!.TrimEnd('/') + "/" + file : file;
                var payloadPath = material.ResolvePath(path, reference, token);
                if (!source.PreservedEntries.TryGetValue(payloadPath, out var bytes)) throw Unsupported();
                var id = (string?)image.Attribute("ID");
                if (id is null || media.ContainsKey(id)) throw Unsupported();
                var mediaType = format is "JPG" or "JPEG" ? "image/jpeg" : format == "BMP" ? "image/bmp" : format == "TIFF" ? "image/tiff" : "image/png";
                media.Add(id!, (path, bytes!, mediaType)); Role(payloadPath, "Payload"); allowedEntries.Add(payloadPath);
            }
        }
        OfdPageXmlContract.ValidateDocumentArea(source, token, material);
        var used = new HashSet<string>(StringComparer.Ordinal);
        var compared = new HashSet<(byte[], byte[])>();
        for (var index = 0; index < source.Pages.Count; index++)
        {
            token.ThrowIfCancellationRequested(); var page = source.Pages[index];
            if (page.Templates.Count != 0 || page.AnnotationAppearances.Count != 0 || page.PreservedPageElements.Count != 0 || page.Elements.Any(element => element is not OfdImageElement)) throw Unsupported();
            var path = material.ResolvePath(documentPath, (declarations[index].Attribute("BaseLoc") ?? throw Unsupported()).Value, token);
            if (page.SourceEntryPath != path || page.Id != (string?)declarations[index].Attribute("ID")) throw Unsupported();
            Read(path, "PageRoot"); OfdPageXmlContract.ValidateForRewrite(source, page, token, material);
            foreach (var image in page.Elements.Cast<OfdImageElement>())
            {
                token.ThrowIfCancellationRequested();
                if (image.ClippingXml is not null || !source.SourceIdentity.TryImage(image, out var owner) || image.ResourceId != owner.Id ||
                    !media.TryGetValue(owner.Id, out var payload) || payload.Entry != owner.Entry || image.MediaType != payload.MediaType || !SamePayload(image.Data, payload.Data)) throw Unsupported();
                if (image.SourceXml is null) throw Unsupported();
                material.ChargeCopy(2L * image.SourceXml.Length, token);
                var graphic = OfdXmlEnvelope.Parse(image.SourceXml); Check(graphic, "ImageObject");
                source.SourceIdentity.AssertFragment(image, graphic, true, material, token);
                if ((string?)graphic.Attribute("ResourceID") != owner.Id) throw Unsupported();
                used.Add(owner.Id);
            }
        }
        if (used.Count != media.Count || source.PreservedEntries.Keys.Any(path => !allowedEntries.Contains(path))) throw Unsupported();
        // Rebuild fresh definitions using the same operation budget. Only the
        // two proven image ID fields are reconstructed; no ownership is forged.
        var options = source.Options; var metadata = options.Metadata;
        var rebuilt = new OfdDocumentPackage
        {
            Options = new OfdDocumentOptions
            {
                Namespace = targetNamespace ?? options.Namespace, DocType = options.DocType, DocumentId = options.DocumentId,
                EnableDeflateCompression = options.EnableDeflateCompression,
                DefaultPageWidthMillimeters = options.DefaultPageWidthMillimeters, DefaultPageHeightMillimeters = options.DefaultPageHeightMillimeters,
                Metadata = new OfdMetadata { Title = metadata.Title, Author = metadata.Author, Subject = metadata.Subject,
                    Keywords = metadata.Keywords, Creator = metadata.Creator, CreationDate = creationDate, ModificationDate = modificationDate }
            }
        };
        foreach (var original in source.Pages)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Node(1);
            var page = new OfdPage { Index = rebuilt.Pages.Count, XMillimeters = original.XMillimeters, YMillimeters = original.YMillimeters,
                WidthMillimeters = original.WidthMillimeters, HeightMillimeters = original.HeightMillimeters };
            string? previousId = null; string? previousType = null; var sequence = 0;
            foreach (var originalImage in original.Elements.Cast<OfdImageElement>())
            {
                if (sequence == 0 || previousId != originalImage.LayerId || previousType != originalImage.LayerType) sequence++;
                previousId = originalImage.LayerId; previousType = originalImage.LayerType;
                var image = (OfdImageElement)OfdModelCloner.CloneElement(originalImage, rebuilt.Options.Namespace, material: material, cancellationToken: token);
                material.ChargeCopy(2L * image.SourceXml!.Length, token);
                var graphic = OfdXmlEnvelope.Parse(image.SourceXml); graphic.SetAttributeValue("ID", null); graphic.SetAttributeValue("ResourceID", null);
                image.SourceXml = OfdXmlEnvelope.Capture(graphic, material, token);
                material.ChargeCopy(64, token);
                image.ObjectId = null; image.ResourceId = string.Empty; image.LayerId = "generated-layer-" + sequence.ToString(CultureInfo.InvariantCulture);
                page.Elements.Add(image);
            }
            rebuilt.Pages.Add(page);
        }
        return rebuilt;

        bool SamePayload(byte[] actual, byte[] expected)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (ReferenceEquals(actual, expected) || compared.Contains((actual, expected))) return true;
            if (actual.Length != expected.Length) return false;
            for (var index = 0; index < actual.Length; index++)
            {
                if ((index & 1023) == 0) token.ThrowIfCancellationRequested();
                if (actual[index] != expected[index]) return false;
            }
            compared.Add((actual, expected)); return true;
        }

        DateTimeOffset? Date(string name)
        {
            var value = bodies[0].Element(ns + "DocInfo")?.Element(ns + name)?.Value;
            if (value is null) return null;
            if (!DateTimeOffset.TryParseExact(value, "O", CultureInfo.InvariantCulture, DateTimeStyles.None, out var date)) throw Unsupported();
            return date;
        }
        void Role(string path, string role)
        {
            if (entryRoles.TryGetValue(path, out var previous) && previous != role) throw Unsupported();
            entryRoles[path] = role;
        }

        XElement Read(string path, string role)
        {
            token.ThrowIfCancellationRequested();
            if (!source.PreservedEntries.TryGetValue(path, out var bytes)) throw new NotSupportedException("Staged raster source entry is missing.");
            Role(path, role);
            var xml = material.ReadXml(path, bytes, token);
            if (xml.Nodes().Any(node => node is not XElement && !(node is XText text && string.IsNullOrWhiteSpace(text.Value)))) throw Unsupported();
            var element = xml.Root!; Check(element, role); allowedEntries.Add(path); return element;
        }
        void Check(XElement rootElement, string rootRole)
        {
            var stack = new Stack<(XElement Node, string Role)>(); stack.Push((rootElement, rootRole));
            while (stack.Count > 0)
            {
                token.ThrowIfCancellationRequested(); material.Structure.Reference(); var frame = stack.Pop();
                var node = frame.Node; var rule = Rules[frame.Role];
                var local = frame.Role is "PageRoot" or "PageDeclaration" ? "Page" : frame.Role == "PageArea" && node.Parent?.Name == ns + "Page" ? "Area" : frame.Role;
                if (node.Name != ns + local || node.Attributes().Any(a => !a.IsNamespaceDeclaration && (a.Name.Namespace != XNamespace.None || !rule.Attributes.Contains(a.Name.LocalName)))) throw Unsupported();
                if (frame.Role == "Layer" && (string?)node.Attribute("Type") is string layerType && layerType is not ("Background" or "Body" or "Foreground")) throw Unsupported();
                foreach (var child in node.Nodes())
                {
                    token.ThrowIfCancellationRequested(); material.Structure.Reference();
                    if (child is XText text) { if (!rule.Text && !string.IsNullOrWhiteSpace(text.Value)) throw Unsupported(); continue; }
                    if (child is not XElement element || element.Name.Namespace != ns || !rule.Children.Contains(element.Name.LocalName)) throw Unsupported();
                    var childElement = (XElement)child;
                    var childRole = frame.Role == "Pages" ? "PageDeclaration" : childElement.Name.LocalName == "Area" ? "PageArea" : childElement.Name.LocalName;
                    stack.Push((childElement, childRole));
                }
                foreach (var name in rule.Children.Where(name => name is not ("Page" or "Layer" or "ImageObject" or "MultiMedia")))
                    if (node.Elements(ns + name).Count() > 1) throw Unsupported();
            }
        }
    }

    private static NotSupportedException Unsupported() => new NotSupportedException("Namespace conversion of staged OFD requires a closed raster image document; unknown structures or references cannot be remapped.");
    private static readonly Dictionary<string, (string[] Children, string[] Attributes, bool Text)> Rules = CreateRules();
    private static Dictionary<string, (string[], string[], bool)> CreateRules()
    {
        var rules = new Dictionary<string, (string[], string[], bool)>(StringComparer.Ordinal);
        void Add(string name, string children = "", string attributes = "", bool text = false) => rules.Add(name,
            (children.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), attributes.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries), text));
        Add("OFD", "DocBody", "Version DocType"); Add("DocBody", "DocInfo DocRoot");
        Add("DocInfo", "Title Author Subject Keywords Creator CreationDate ModDate");
        foreach (var name in new[] { "Title", "Author", "Subject", "Keywords", "Creator", "CreationDate", "ModDate", "DocRoot", "MaxUnitID", "PhysicalBox", "PublicRes", "DocumentRes", "MediaFile" }) Add(name, text: true);
        Add("Document", "CommonData Pages"); Add("CommonData", "MaxUnitID PageArea PublicRes DocumentRes"); Add("PageArea", "PhysicalBox");
        Add("Pages", "Page"); Add("PageDeclaration", attributes: "ID BaseLoc"); Add("PageRoot", "Area Content");
        Add("Content", "Layer"); Add("Layer", "ImageObject", "ID Type"); Add("ImageObject", attributes: "ID ResourceID Boundary CTM Alpha");
        Add("Res", "MultiMedias", "BaseLoc"); Add("MultiMedias", "MultiMedia"); Add("MultiMedia", "MediaFile", "ID Type Format");
        return rules;
    }
}
