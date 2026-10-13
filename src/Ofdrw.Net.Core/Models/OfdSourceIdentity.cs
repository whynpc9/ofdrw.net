using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Core.Models;

// Reader-owned identity is independent of caller-editable IDs and XML strings.
// A preorder slot identifies the exact source node; unrelated insertions can
// conservatively invalidate it. No source DOM is retained.
internal sealed class OfdSourceIdentity
{
    internal sealed class Owner(string entry, XName kind, int slot, string id, bool activeDefinition)
    {
        internal readonly string Entry = entry;
        internal readonly XName Kind = kind;
        internal readonly int Slot = slot;
        internal readonly string Id = id;
        internal readonly bool ActiveDefinition = activeDefinition;
        internal long Number => ParseId(Id);
    }
    private readonly HashSet<Owner> emptyLayers = new();
    private readonly Dictionary<object, Owner> models = new();
    private readonly Dictionary<OfdElement, Owner> layers = new();
    private sealed class IdSlot(int ordinal, XName node, XName attribute, string value, bool activeFragmentDefinition)
    {
        internal readonly int Ordinal = ordinal;
        internal readonly XName Node = node;
        internal readonly XName Attribute = attribute;
        internal readonly string Value = value;
        internal readonly bool ActiveFragmentDefinition = activeFragmentDefinition;
    }
    private sealed class GraphicIds(List<IdSlot> source, List<IdSlot> clips)
    {
        internal readonly List<IdSlot> Source = source;
        internal readonly List<IdSlot> Clips = clips;
    }
    private readonly Dictionary<OfdElement, GraphicIds> graphicIds = new();
    private readonly Dictionary<OfdImageElement, Owner> images = new();
    private readonly Dictionary<OfdPage, string> pagePaths = new();
    private readonly Dictionary<long, string> templatePaths = new();
    internal string Namespace { get; private set; } = string.Empty;
    internal bool TryImage(OfdImageElement image, out Owner owner) => images.TryGetValue(image, out owner!);
    private readonly Dictionary<string, Owner> resources = new(StringComparer.Ordinal);
    private readonly Dictionary<long, Owner> active = new();
    private readonly Dictionary<string, XName> rootNames = new(StringComparer.OrdinalIgnoreCase);
    private readonly HashSet<string> trustedEntries = new(StringComparer.OrdinalIgnoreCase);
    private readonly Dictionary<string, Dictionary<int, Owner>> entries = new(StringComparer.OrdinalIgnoreCase);
    internal IReadOnlyCollection<Owner> Active => active.Values;
    internal IEnumerable<string> ControlEntries => trustedEntries;
    internal bool TryActive(long id, out Owner owner) => active.TryGetValue(id, out owner!);
    internal OfdPage[] OriginalPages { get; private set; } = [];

    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("identityNamespace", Namespace); sink.Item("originalPages", OriginalPages, true); sink.Item("count", OriginalPages.Length);
        foreach (var page in OriginalPages) sink.Item("originalPage", page, true);
        Map(models); Map(layers); Map(images); Map(resources); Map(active);
        sink.Item("emptyLayers", emptyLayers, true); sink.Item("count", emptyLayers.Count); foreach (var owner in emptyLayers) OwnerValue(owner);
        sink.Item("entrySlots", entries, true); sink.Item("count", entries.Count);
        foreach (var entry in entries) { sink.Item("entry", entry.Key); Map(entry.Value); }
        sink.Item("graphicIds", graphicIds, true); sink.Item("count", graphicIds.Count);
        foreach (var item in graphicIds)
        {
            sink.Item("graphic", item.Key, true); Slots(item.Value.Source); Slots(item.Value.Clips);
        }
        Strings(pagePaths); Strings(templatePaths); Strings(rootNames);
        sink.Item("trusted", trustedEntries, true); sink.Item("count", trustedEntries.Count);
        foreach (var path in trustedEntries) sink.Item("trustedPath", path);
        void Map<T>(Dictionary<T, Owner> values) where T : notnull
        {
            sink.Item("ownerMap", values, true); sink.Item("count", values.Count);
            foreach (var pair in values) { sink.Item("key", pair.Key, pair.Key is not string && pair.Key is not ValueType); OwnerValue(pair.Value); }
        }
        void OwnerValue(Owner owner)
        { sink.Item("ownerEntry", owner.Entry); sink.Item("kind", owner.Kind); sink.Item("slot", owner.Slot); sink.Item("id", owner.Id); sink.Item("active", owner.ActiveDefinition); }
        void Slots(List<IdSlot> slots)
        {
            sink.Item("idSlots", slots, true); sink.Item("count", slots.Count);
            foreach (var slot in slots) { sink.Item("ordinal", slot.Ordinal); sink.Item("node", slot.Node); sink.Item("attribute", slot.Attribute); sink.Item("value", slot.Value); sink.Item("role", slot.ActiveFragmentDefinition); }
        }
        void Strings<T, U>(Dictionary<T, U> values) where T : notnull
        {
            sink.Item("valueMap", values, true); sink.Item("count", values.Count);
            foreach (var pair in values) { sink.Item("key", pair.Key, pair.Key is OfdPage); sink.Item("value", pair.Value); }
        }
    }
    internal bool TryOwner(object model, out Owner owner) => models.TryGetValue(model, out owner!);
    internal bool TryEmptyLayer(OfdPage page, string? id, out Owner owner)
    {
        owner = null!;
        return pagePaths.TryGetValue(page, out var path) && active.TryGetValue(ParseId(id), out owner) &&
            owner.Entry == path && emptyLayers.Contains(owner);
    }
    internal bool TryLayer(OfdElement model, out Owner owner) => layers.TryGetValue(model, out owner!);
    internal bool TryResource(string kind, string? id, out Owner owner)
        => resources.TryGetValue(kind + ":" + ParseId(id).ToString(CultureInfo.InvariantCulture), out owner!);

    internal static long ParseId(string? id) => long.TryParse(id, NumberStyles.None, CultureInfo.InvariantCulture, out var n) && n > 0 ? n : 0;

    internal static OfdSourceIdentity Capture(OfdDocumentPackage package, OfdModelMaterialization budget, CancellationToken token)
    {
        foreach (var page in package.Pages) { token.ThrowIfCancellationRequested(); budget.Structure.Reference(); }
        budget.ChargeCopy((long)package.Pages.Count * IntPtr.Size, token);
        budget.ReserveReadModel(4096L + (package.Pages.Count + package.Fonts.Count) * 512L, token);
        var result = new OfdSourceIdentity { OriginalPages = package.Pages.ToArray() };
        var pagesByIndex = package.Pages.ToDictionary(page => page.Index);
        var fontsById = new Dictionary<long, List<OfdFontResource>>();
        foreach (var font in package.Fonts)
        {
            var id = ParseId(font.Id); if (id == 0) continue;
            if (!fontsById.TryGetValue(id, out var items)) fontsById.Add(id, items = new List<OfdFontResource>());
            items.Add(font);
        }
        var documentPath = package.DocumentEntryPath!;
        var document = Read(documentPath);
        var ns = document.Root!.Name.Namespace; result.Namespace = ns.NamespaceName;
        var documentNodes = Register(documentPath, document, "Document");
        foreach (var item in (document.Root.Element(ns + "Pages")?.Elements(ns + "Page") ?? Enumerable.Empty<XElement>()).Select((node, index) => (Node: node, Index: index)))
        {
            token.ThrowIfCancellationRequested();
            var declaration = item.Node; var path = declaration.Attribute("BaseLoc")?.Value;
            if (path is null) continue;
            path = budget.ResolvePath(documentPath, path, token);
            if (pagesByIndex.TryGetValue(item.Index, out var page)) { Bind(page, declaration, documentNodes); result.pagePaths.Add(page, path); }
        }
        var resourcePaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var name in new[] { "PublicRes", "DocumentRes" })
            foreach (var declaration in document.Root.Element(ns + "CommonData")?.Elements(ns + name) ?? Enumerable.Empty<XElement>())
                resourcePaths.Add(budget.ResolvePath(documentPath, budget.ReadXmlValue(declaration, token), token));
        var contentPaths = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in package.Pages)
        {
            if (page.SourceEntryPath is string path) contentPaths.Add(path);
            foreach (var template in page.Templates)
                if (template.BaseLocation is string location) contentPaths.Add(budget.ResolvePath(documentPath, location, token));
        }
        foreach (var declaration in document.Root.Element(ns + "CommonData")?.Elements(ns + "TemplatePage") ?? Enumerable.Empty<XElement>())
            if (declaration.Attribute("BaseLoc") is XAttribute location && !string.IsNullOrWhiteSpace(location.Value))
            {
                var path = budget.ResolvePath(documentPath, location.Value, token);
                contentPaths.Add(path); var id = ParseId(declaration.Attribute("ID")?.Value);
                if (id != 0) result.templatePaths.Add(id, path);
            }
        foreach (var path in contentPaths) resourcePaths.Add(budget.ReadDirectorySuffix(path, "/PageRes.xml", token));
        foreach (var path in resourcePaths)
        {
            token.ThrowIfCancellationRequested(); if (!package.PreservedEntries.ContainsKey(path)) continue;
            var xml = Read(path); var owners = Register(path, xml, "Res"); var resourceNs = xml.Root!.Name.Namespace;
            foreach (var node in xml.Root.DescendantsAndSelf())
            {
                if (node.Name.Namespace != resourceNs || !owners.TryGetValue(node, out var owner)) continue;
                if (node.Name.LocalName is "Font" or "MultiMedia")
                {
                    budget.ReserveReadModel(512, token);
                    var storage = budget.IsControlledRead ? budget.ReserveReadStorage(64, token) : null;
                    var key = node.Name.LocalName + ":" + owner.Number.ToString(CultureInfo.InvariantCulture);
                    if (storage is not null) budget.AdoptReadString(key, storage, token);
                    result.resources.Add(key, owner);
                }
                if (node.Name.LocalName == "Font" && fontsById.TryGetValue(owner.Number, out var matchedFonts))
                    foreach (var font in matchedFonts) result.models.Add(font, owner);
            }
        }
        foreach (var path in contentPaths)
            if (package.PreservedEntries.ContainsKey(path)) Register(path, Read(path), "Page");
        foreach (var page in package.Pages)
        {
            if (page.SourceEntryPath is string path) BindGraphics(path, page.Elements);
            foreach (var template in page.Templates)
                if (template.BaseLocation is string location) BindGraphics<OfdElement>(budget.ResolvePath(documentPath, location, token), template.Elements);
        }
        // Strict editing admission must close the same authenticated resource scope.
        // Permissively rendered legacy annotations remain unclaimed when it cannot.
        foreach (var page in package.Pages)
        {
            var collection = page.Annotations;
            if (collection is null) continue;
            budget.ChargeCopy(256L + collection.Items.Count * 16L, token); var admitted = new OfdAnnotationCollection();
            foreach (var annotation in collection.Items)
            {
                token.ThrowIfCancellationRequested();
                var closed = true;
                foreach (var graphic in annotation.Appearance.Elements)
                {
                    budget.Structure.Reference();
                    if (graphic is OfdTextElement text && (!result.TryResource("Font", text.FontResourceId, out _) ||
                        !fontsById.ContainsKey(ParseId(text.FontResourceId)))) closed = false;
                    if (graphic is OfdImageElement image && !result.TryResource("MultiMedia", image.ResourceId, out _)) closed = false;
                }
                if (closed) admitted.Items.Add(annotation); else page.AnnotationOwnership.HasUnclaimedSource = true;
            }
            if (admitted.Items.Count != collection.Items.Count)
                page.AnnotationOwnership.Bind(page, admitted.Items.Count == 0 ? null : admitted, budget, token);
        }
        budget.ReserveReadModel(512, token);
        var annotationEntries = new Dictionary<string, (XDocument Xml, Dictionary<XElement, int> Ordinals, XElement[] Slots)>(StringComparer.OrdinalIgnoreCase);
        foreach (var page in package.Pages)
        foreach (var annotation in page.AnnotationOwnership.Original?.Items ?? Enumerable.Empty<OfdAnnotation>())
        {
            var source = annotation.Source!;
            if (!annotationEntries.TryGetValue(source.Entry, out var entry))
            {
                var annotationDocument = Read(source.Entry); budget.ReserveReadModel(512, token);
                var ordinals = new Dictionary<XElement, int>(); var ordinal = 0;
                foreach (var descendant in annotationDocument.Root!.DescendantsAndSelf())
                { token.ThrowIfCancellationRequested(); budget.ChargeCopy(128, token); ordinals.Add(descendant, ordinal++); }
                budget.ChargeCopy(16L * annotationDocument.Root.Elements().Count(), token);
                entry = (annotationDocument, ordinals, annotationDocument.Root.Elements().ToArray()); annotationEntries.Add(source.Entry, entry);
            }
            var root = entry.Xml.Root!;
            if (!result.entries.TryGetValue(source.Entry, out var entrySlots))
            {
                budget.ReserveReadModel(512, token);
                entrySlots = new Dictionary<int, Owner>(); result.entries.Add(source.Entry, entrySlots);
                result.rootNames.Add(source.Entry, root.Name); result.trustedEntries.Add(source.Entry);
            }
            var node = entry.Slots[source.Slot];
            AnnotationBind(annotation, node);
            var appearance = node.Element(node.Name.Namespace + "Appearance")!;
            budget.ReserveReadModel(256L + 32L * appearance.Elements().Count(), token);
            var children = appearance.Elements().ToArray();
            for (var index = 0; index < children.Length; index++)
            {
                var graphic = annotation.Appearance.Elements[index]; var child = children[index];
                AnnotationBind(graphic, child);
                result.graphicIds.Add(graphic, new GraphicIds(CaptureIds(child, budget, token), CaptureIds(child.Element(child.Name.Namespace + "Clips"), budget, token)));
                if (graphic is OfdImageElement image && result.TryResource("MultiMedia", child.Attribute("ResourceID")?.Value, out var media)) result.images.Add(image, media);
            }
            void AnnotationBind(object model, XElement sourceNode)
            {
                token.ThrowIfCancellationRequested(); budget.ChargeCopy(256, token);
                var owner = new Owner(source.Entry, sourceNode.Name, entry.Ordinals[sourceNode], sourceNode.Attribute("ID")!.Value, true);
                if (result.active.ContainsKey(owner.Number)) throw new InvalidDataException("Annotation active ID admission is ambiguous.");
                result.active.Add(owner.Number, owner); entrySlots.Add(owner.Slot, owner); result.models.Add(model, owner);
            }
        }
        return result;

        XDocument Read(string path) => budget.ReadXml(path, package.PreservedEntries[path], token);
        Dictionary<XElement, Owner> Register(string path, XDocument xml, string expectedRoot)
        {
            budget.ReserveReadModel(1024, token);
            var nodes = new Dictionary<XElement, Owner>(); var ordinal = 0;
            var sourceNs = xml.Root!.Name.Namespace;
            var known = result.entries.TryGetValue(path, out var slots);
            if (!known)
            {
                result.entries.Add(path, slots = new Dictionary<int, Owner>()); result.rootNames.Add(path, xml.Root.Name);
                if (xml.Root.Name.LocalName == expectedRoot && (sourceNs == ns || sourceNs.NamespaceName == OfdConstants.Namespace || sourceNs.NamespaceName == OfdConstants.StandardNamespace)) result.trustedEntries.Add(path);
            }
            if (!result.trustedEntries.Contains(path)) return nodes;
            foreach (var node in xml.Root.DescendantsAndSelf())
            {
                token.ThrowIfCancellationRequested(); budget.Structure.Node(1); var slot = ordinal++;
                var id = node.Attribute("ID")?.Value;
                var definitionRole = Accepted(node.Name.Namespace, sourceNs) && IsDefinition(node, xml.Root);
                if (definitionRole && id is not null && ParseId(id) == 0) throw new InvalidDataException("Unsupported non-positive or non-numeric source document-unit ID.");
                var activeDefinition = definitionRole && ParseId(id) != 0;
                var opaqueRoot = node.Parent?.Name.LocalName == "Layer" && node.Parent.Parent?.Name.LocalName == "Content" && node.Parent.Parent.Parent == xml.Root;
                if (!activeDefinition && !opaqueRoot) continue;
                id ??= string.Empty;
                budget.ReserveReadModel(2048, token);
                Owner owner;
                if (known)
                {
                    if (!slots!.TryGetValue(slot, out owner!) || owner.Kind != node.Name || owner.Id != id)
                        throw new InvalidDataException("Inconsistent source identity during read.");
                }
                else
                {
                    budget.StoreString(path); budget.StoreString(node.Name.NamespaceName); budget.StoreString(node.Name.LocalName); budget.StoreString(id);
                    owner = new Owner(path, node.Name, slot, id, activeDefinition);
                    if (activeDefinition)
                    {
                        if (result.active.ContainsKey(owner.Number)) throw new InvalidDataException("Duplicate active document-unit ID.");
                        budget.Structure.Reference(); result.active.Add(owner.Number, owner);
                    }
                    slots!.Add(slot, owner);
                }
                nodes.Add(node, owner);
                if (node.Name.LocalName == "Layer" && !node.HasElements) result.emptyLayers.Add(owner);
            }
            return nodes;
        }
        void Bind(object model, XElement node, Dictionary<XElement, Owner> nodes)
        {
            if (nodes.TryGetValue(node, out var owner)) { budget.Structure.Reference(); result.models.Add(model, owner); }
        }
        void BindGraphics<T>(string path, List<T> elements) where T : OfdElement
        {
            var xml = Read(path); var nodes = Register(path, xml, "Page"); var sourceNs = xml.Root!.Name.Namespace;
            budget.ReserveReadModel(512L + elements.Count * 128L, token);
            var originals = xml.Root.Elements().FirstOrDefault(node => node.Name.LocalName == "Content")?.Elements().Where(node => node.Name.LocalName == "Layer").SelectMany(layer => layer.Elements()).ToList() ?? [];
            if (originals.Count != elements.Count) throw new InvalidDataException("Source graphic slots cannot be authenticated.");
            for (var i = 0; i < originals.Count; i++)
            {
                token.ThrowIfCancellationRequested(); budget.ReserveReadModel(2048, token); var element = elements[i]; var node = originals[i];
                Bind(element, node, nodes);
                result.graphicIds.Add(element, new GraphicIds(CaptureIds(node, budget, token),
                    CaptureIds(node.Element(node.Name.Namespace + "Clips"), budget, token)));
                if (element is OfdImageElement image && result.TryResource("MultiMedia", node.Attribute("ResourceID")?.Value, out var media))
                    result.images.Add(image, media);
                if (node.Parent is XElement layer && nodes.TryGetValue(layer, out var owner)) { budget.Structure.Reference(); result.layers.Add(element, owner); }
            }
        }
    }

    private static bool Accepted(XNamespace value, XNamespace source) => value == source || value == XNamespace.None ||
        value.NamespaceName == OfdConstants.Namespace || value.NamespaceName == OfdConstants.StandardNamespace;

    private static bool IsDefinition(XElement node, XElement root)
    {
        var ns = root.Name.Namespace;
        if (root.Name.LocalName == "Document")
            return node.Name == ns + "Page" && node.Parent?.Name == ns + "Pages" && node.Parent.Parent == root ||
                   node.Name == ns + "TemplatePage" && node.Parent?.Name == ns + "CommonData" && node.Parent.Parent == root;
        if (root.Name.LocalName == "Res")
        {
            var container = node.Name.LocalName switch
            { "Font" => "Fonts", "MultiMedia" => "MultiMedias", "DrawParam" => "DrawParams", "ColorSpace" => "ColorSpaces", "CompositeGraphicUnit" => "CompositeGraphicUnits", _ => null };
            return container is not null && node.Parent?.Name == ns + container && node.Parent.Parent == root;
        }
        if (node.Name.LocalName == "Layer" && node.Parent?.Name != ns + "Content") return false;
        if (root.Name.LocalName != "Page" || node.Name.LocalName is not ("Layer" or "PageBlock" or "TextObject" or "ImageObject" or "PathObject" or "CompositeObject")) return false;
        for (var parent = node.Parent; parent is not null; parent = parent.Parent)
        {
            if (parent.Name == ns + "Content") return parent.Parent == root;
            if (!Accepted(parent.Name.Namespace, ns) || parent.Name.LocalName is not ("Layer" or "PageBlock")) return false;
        }
        return false;
    }

    internal void ValidateResourceSource(string path, OfdDocumentPackage package)
    {
        if (!package.PreservedEntries.ContainsKey(path)) return; // Writer-created table; IDs are planned independently.
        if (!entries.ContainsKey(path) || !trustedEntries.Contains(path) || !rootNames.TryGetValue(path, out var name) || name.LocalName != "Res")
            throw new NotSupportedException("New resource source tables must be read before editing.");
    }

    internal void ValidateTemplateSource(long id, XElement declaration, string documentPath, OfdModelMaterialization material, CancellationToken token)
    {
        if (!templatePaths.TryGetValue(id, out var original) || declaration.Attribute("BaseLoc") is not XAttribute location ||
            !string.Equals(original, material.ResolvePath(documentPath, location.Value, token), StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("New or relocated template sources must be read before editing.");
    }

    internal OfdSourceIdentity CloneFor(OfdDocumentPackage destination, IReadOnlyDictionary<object, object> mapping,
        OfdModelMaterialization budget, CancellationToken token, bool requireClosedMap = false)
    {
        var clone = new OfdSourceIdentity { Namespace = Namespace };
        foreach (var pair in active) { token.ThrowIfCancellationRequested(); budget.ChargeCopy(128, token); clone.active.Add(pair.Key, pair.Value); }
        foreach (var pair in entries)
        {
            token.ThrowIfCancellationRequested(); budget.ChargeCopy(64L * pair.Value.Count, token);
            clone.entries.Add(pair.Key, new Dictionary<int, Owner>(pair.Value));
        }
        foreach (var pair in resources) { budget.ChargeCopy(128, token); clone.resources.Add(pair.Key, pair.Value); }
        foreach (var pair in templatePaths) { budget.ChargeCopy(128, token); clone.templatePaths.Add(pair.Key, pair.Value); }
        foreach (var pair in rootNames) { budget.ChargeCopy(128, token); clone.rootNames.Add(pair.Key, pair.Value); }
        foreach (var path in trustedEntries) { budget.ChargeCopy(128, token); clone.trustedEntries.Add(path); }
        foreach (var owner in emptyLayers) { budget.ChargeCopy(128, token); clone.emptyLayers.Add(owner); }
        foreach (var pair in models) if (mapping.TryGetValue(pair.Key, out var model)) { budget.ChargeCopy(128, token); clone.models.Add(model, pair.Value); }
        foreach (var pair in layers) if (mapping.TryGetValue(pair.Key, out var model)) { budget.ChargeCopy(128, token); clone.layers.Add((OfdElement)model, pair.Value); }
        foreach (var pair in graphicIds) if (mapping.TryGetValue(pair.Key, out var model))
        {
            budget.ChargeCopy((long)(pair.Value.Source.Count + pair.Value.Clips.Count) * IntPtr.Size + 64, token);
            clone.graphicIds.Add((OfdElement)model, new GraphicIds(new List<IdSlot>(pair.Value.Source), new List<IdSlot>(pair.Value.Clips)));
        }
        foreach (var pair in images) if (mapping.TryGetValue(pair.Key, out var model)) { budget.ChargeCopy(128, token); clone.images.Add((OfdImageElement)model, pair.Value); }
        foreach (var pair in pagePaths) if (mapping.TryGetValue(pair.Key, out var model)) { budget.ChargeCopy(128, token); clone.pagePaths.Add((OfdPage)model, pair.Value); }
        budget.ChargeCopy((long)OriginalPages.Length * IntPtr.Size, token);
        clone.OriginalPages = OriginalPages.Select(page => mapping.TryGetValue(page, out var model) ? (OfdPage)model :
            requireClosedMap ? throw new InvalidOperationException("Historical page mapping is incomplete.") : page).ToArray();
        return clone;
    }

    private static List<IdSlot> CaptureIds(XElement? root, OfdModelMaterialization budget, CancellationToken token)
    {
        budget.ReserveReadModel(128, token);
        var result = new List<IdSlot>(); if (root is null) return result; var ordinal = 0;
        foreach (var node in root.DescendantsAndSelf())
        {
            token.ThrowIfCancellationRequested(); budget.Structure.Node(1);
            foreach (var attribute in node.Attributes().Where(attribute => attribute.Name.LocalName == "ID"))
            {
                budget.Structure.Reference(); budget.StoreString(attribute.Value);
                budget.ReserveReadModel(256, token);
                result.Add(new IdSlot(ordinal, node.Name, attribute.Name, attribute.Value, FragmentDefinition(node, root)));
            }
            ordinal++;
        }
        return result;
    }

    private static bool FragmentDefinition(XElement node, XElement root)
    {
        static bool Unit(XElement element, XNamespace ns) => Accepted(element.Name.Namespace, ns) &&
            element.Name.LocalName is "PageBlock" or "TextObject" or "ImageObject" or "PathObject" or "CompositeObject";
        if (!Unit(node, root.Name.Namespace)) return false;
        if (ReferenceEquals(node, root)) return true;
        for (var parent = node.Parent; parent is not null; parent = parent.Parent)
        {
            if (!Accepted(parent.Name.Namespace, root.Name.Namespace) || parent.Name.LocalName != "PageBlock") return false;
            if (ReferenceEquals(parent, root)) return true;
        }
        return ReferenceEquals(node, root);
    }

    internal void AssertFragment(OfdElement model, XElement? current, bool source, OfdModelMaterialization budget, CancellationToken token)
    {
        if (!graphicIds.TryGetValue(model, out var original)) return;
        if (source && current is not null && TryOwner(model, out var rootOwner) && current.Name != rootOwner.Kind)
            throw new InvalidDataException("Imported fragment root kind or namespace changed.");
        var expected = source ? original.Source : original.Clips;
        var actual = CaptureIds(current, budget, token);
        // Typed primitive roots are assigned through their authenticated owner.
        // Removing only that redundant root ID from SourceXml is safe; retained
        // nested IDs and opaque roots still require exact slot/name/value proof.
        bool Root(IdSlot item) => source && model is not OfdRawElement && item.Ordinal == 0 && item.Attribute == "ID";
        var a = 0; var b = 0;
        while (true)
        {
            token.ThrowIfCancellationRequested();
            while (a < expected.Count && Root(expected[a])) a++;
            while (b < actual.Count && Root(actual[b])) b++;
            if (a == expected.Count || b == actual.Count) break;
            var old = expected[a++]; var value = actual[b++];
            if (old.Ordinal != value.Ordinal || old.Node != value.Node || old.Attribute != value.Attribute || old.Value != value.Value || old.ActiveFragmentDefinition != value.ActiveFragmentDefinition)
                throw new InvalidDataException("Imported fragment ID provenance is stale; read back the changed source before editing.");
        }
        if (a != expected.Count || b != actual.Count) throw new InvalidDataException("Imported fragment gained or lost unmodeled ID slots.");
        if (source && current?.Attribute("ID") is XAttribute id && TryOwner(model, out var owner) && owner.ActiveDefinition && id.Value != owner.Id)
            throw new InvalidDataException("Imported fragment root ID disagrees with its authenticated owner.");
    }

    internal void Reconcile(OfdDocumentPackage package, OfdModelMaterialization budget, CancellationToken token)
    {
        if (package.Options.Namespace != Namespace) throw new NotSupportedException("Imported source namespace cannot be remapped safely.");
        foreach (var page in package.Pages)
            if (pagePaths.TryGetValue(page, out var path) && page.SourceEntryPath != path)
                throw new InvalidDataException("Imported page source path was changed without a reference remap.");
        foreach (var entry in entries)
        {
            token.ThrowIfCancellationRequested();
            // Pruning can remove an entire unused page entry. Only claims of
            // retained source owners require its continued presence.
            if (!package.PreservedEntries.TryGetValue(entry.Key, out var bytes)) continue;
            if (budget.OpaqueEntries.Contains(entry.Key)) throw new NotSupportedException("Imported active source was replaced by opaque content.");
            var xml = budget.ReadXml(entry.Key, bytes, token); var ordinal = 0;
            if (xml.Root?.Name != rootNames[entry.Key]) throw new NotSupportedException("Imported source root role changed.");
            if (!trustedEntries.Contains(entry.Key)) continue;
            var seen = new HashSet<int>();
            foreach (var node in xml.Root!.DescendantsAndSelf())
            {
                token.ThrowIfCancellationRequested(); var slot = ordinal++;
                if (!entry.Value.TryGetValue(slot, out var owner))
                {
                    if (Accepted(node.Name.Namespace, xml.Root.Name.Namespace) && IsDefinition(node, xml.Root) && ParseId(node.Attribute("ID")?.Value) != 0)
                        throw new InvalidDataException("Preserved source gained an unauthenticated active definition.");
                    continue;
                }
                var currentActive = Accepted(node.Name.Namespace, xml.Root.Name.Namespace) &&
                    (IsDefinition(node, xml.Root) || xml.Root.Name.LocalName == "PageAnnot" &&
                     (node.Parent == xml.Root && node.Name == xml.Root.Name.Namespace + "Annot" ||
                      node.Parent?.Name == xml.Root.Name.Namespace + "Appearance" && node.Parent.Parent?.Name == xml.Root.Name.Namespace + "Annot" && node.Parent.Parent.Parent == xml.Root &&
                      node.Name.LocalName is "TextObject" or "ImageObject" or "PathObject")) && ParseId(node.Attribute("ID")?.Value) != 0;
                if (currentActive != owner.ActiveDefinition) throw new InvalidDataException("Preserved definition role changed.");
                if (node.Name != owner.Kind || (node.Attribute("ID")?.Value ?? string.Empty) != owner.Id)
                    throw new InvalidDataException("Preserved source identity is stale.");
                seen.Add(slot);
            }
            if (seen.Count != entry.Value.Count) throw new NotSupportedException("Preserved source identity is incomplete.");
        }
    }
}
