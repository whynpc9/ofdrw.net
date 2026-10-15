using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Core.Constants;

namespace Ofdrw.Net.Packaging;

// Reservation, authenticated imported ownership and emitted claims are distinct.
internal sealed class OfdIdPlan
{
    private readonly OfdDocumentPackage package;
    private readonly OfdModelMaterialization budget;
    private readonly CancellationToken token;
    private readonly HashSet<long> originalReserved = new();
    private readonly HashSet<long> fixedFragmentIds = new();
    private readonly HashSet<long> reserved = new();
    private readonly Dictionary<long, object> emitted = new();
    private readonly Dictionary<object, string> assignments = new();
    private readonly HashSet<OfdSourceIdentity.Owner> claimedLayers = new();
    private readonly Dictionary<long, OfdElement> fragmentOwners = new();
    private long cursor = 1;
    private bool exhausted;
    internal OfdModelMaterialization Material => budget;
    internal CancellationToken Token => token;
    internal string Assigned(object model) => assignments.TryGetValue(model, out var id) ? id : throw new InvalidDataException("Unplanned composite definition.");
    internal long MaxId { get; private set; }

    internal OfdIdPlan(OfdDocumentPackage package, OfdModelMaterialization budget, CancellationToken token)
    {
        this.package = package; this.budget = budget; this.token = token;
        package.SourceIdentity?.Reconcile(package, budget, token);
        foreach (var entry in package.PreservedEntries)
        {
            token.ThrowIfCancellationRequested();
            if (budget.OpaqueEntries.Contains(entry.Key)) continue;
            var reservations = budget.Reservations(entry.Value, out var hasHint, token);
            foreach (var id in reservations) Reserve(id, true);
            if (!hasHint) continue;
            var xml = budget.ReadXml(entry.Key, entry.Value, token);
            foreach (var node in xml.Descendants())
            {
                token.ThrowIfCancellationRequested();
                foreach (var attribute in node.Attributes())
                    if (attribute.Name.LocalName is "ID" or "MaxUnitID") Reserve(OfdSourceIdentity.ParseId(attribute.Value), true);
                if (node.Name.LocalName == "MaxUnitID") AdvancePast(OfdSourceIdentity.ParseId(node.Value));
            }
        }
        var commonDefinitions = new HashSet<long>();
        foreach (var xml in package.PreservedCommonDataElements)
        {
            var root = Fixed(xml);
            if (root.Name == XNamespace.Get(package.Options.Namespace) + "TemplatePage")
            {
                var id = OfdSourceIdentity.ParseId(root.Attribute("ID")?.Value);
                if (id == 0) throw new InvalidDataException("Template declaration requires a positive numeric ID.");
                if (!commonDefinitions.Add(id)) throw new InvalidDataException("Duplicate template definition in CommonData fragments.");
                if (package.SourceIdentity is null) throw new NotSupportedException("Template source declarations must be read before editing.");
                package.SourceIdentity.ValidateTemplateSource(id, root, package.DocumentEntryPath!, budget, token);
                if (package.SourceIdentity?.TryActive(id, out var owner) == true && (owner.Kind != root.Name || owner.Entry != package.DocumentEntryPath))
                    throw new InvalidDataException("Template fragment collides with a different active source definition.");
            }
        }
        foreach (var xml in package.PreservedDocumentElements) Fixed(xml);
        foreach (var xml in package.PreservedDocBodyElements) Fixed(xml);
        foreach (var page in package.Pages) foreach (var xml in page.PreservedPageElements) Fixed(xml);
        foreach (var page in package.Pages)
            foreach (var element in OfdAnnotationContract.Definitions(page))
            {
                token.ThrowIfCancellationRequested();
                var xml = element switch { OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, OfdRawElement raw => raw.Xml, _ => null };
                Fragment(element, xml, true); Fragment(element, element.ClippingXml, false);
            }
        // Preferred fresh definitions must be known before any anonymous ID
        // allocation, otherwise an early page/layer can steal a later font ID.
        foreach (var resource in package.CompositeResources?.Items ?? Enumerable.Empty<OfdCompositeResource>())
        {
            Hint(resource.Id);
            foreach (var element in resource.Elements)
            {
                Hint(element.ObjectId);
                var source = element switch { OfdCompositeElement composite => composite.SourceXml, OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, OfdRawElement raw => raw.Xml, _ => null };
                Fragment(element, source, true); Fragment(element, element.ClippingXml, false);
            }
        }
        foreach (var font in package.Fonts) Hint(font.Id);
        foreach (var page in package.Pages)
        {
            Hint(page.Id);
            foreach (var annotation in (page.Annotations ?? page.AnnotationOwnership.Original)?.Items ?? Enumerable.Empty<OfdAnnotation>()) Hint(annotation.Id);
            foreach (var element in OfdAnnotationContract.Definitions(page))
            {
                Hint(element.ObjectId); Hint(element.LayerId);
                if (element is OfdImageElement image) Hint(image.ResourceId);
                if (element is OfdTextElement text) Hint(text.FontResourceId);
            }
        }
        XElement Fixed(string xml)
        {
            var root = OfdXmlEnvelope.Parse(xml);
            foreach (var attribute in root.DescendantsAndSelf().Attributes().Where(attribute => attribute.Name.LocalName == "ID"))
            {
                token.ThrowIfCancellationRequested(); var id = OfdSourceIdentity.ParseId(attribute.Value);
                Reserve(id, false); if (id != 0) fixedFragmentIds.Add(id);
            }
            return root;
        }
        void Hint(string? value) { token.ThrowIfCancellationRequested(); budget.Structure.Reference(); AdvancePast(OfdSourceIdentity.ParseId(value)); }
    }

    private void Reserve(long id, bool original)
    {
        if (id == 0) return;
        budget.Structure.Reference(); reserved.Add(id); if (original) originalReserved.Add(id); AdvancePast(id);
    }
    private void Fragment(OfdElement model, string? source, bool rootOwner)
    {
        var root = string.IsNullOrWhiteSpace(source) ? null : OfdXmlEnvelope.Parse(source!);
        package.SourceIdentity?.AssertFragment(model, root, rootOwner, budget, token);
        if (root is null) return;
        var imported = package.SourceIdentity?.TryOwner(model, out _) == true;
        var activeIds = new HashSet<long>();
        foreach (var node in root.DescendantsAndSelf())
        {
            token.ThrowIfCancellationRequested(); budget.Structure.Node(1);
            foreach (var attribute in node.Attributes().Where(a => a.Name.LocalName == "ID"))
            {
                var id = OfdSourceIdentity.ParseId(attribute.Value);
                if (attribute.Name == "ID" && IsFragmentUnit(node, root))
                {
                    if (id == 0) throw new InvalidDataException("Unsupported non-numeric source document-unit ID.");
                    if (!activeIds.Add(id)) throw new InvalidDataException("Duplicate active IDs inside a source fragment.");
                }
                if (id == 0) continue;
                Reserve(id, false);
                if (imported) continue; // Current preserved source is reconciled separately.
                if (originalReserved.Contains(id) || fixedFragmentIds.Contains(id)) throw new InvalidDataException($"New {model.GetType().Name} source XML collides with preserved numeric ID {id}.");
                if (IsClipPathId(node)) continue; // CT_Path's legacy ID extension is not a document-unit definition.
                if (fragmentOwners.TryGetValue(id, out var other) && !ReferenceEquals(other, model))
                    throw new InvalidDataException("Two new source fragments claim the same numeric ID.");
                fragmentOwners[id] = model;
                if (ReferenceEquals(node, root) && rootOwner && attribute.Name == "ID" &&
                    OfdSourceIdentity.ParseId(model.ObjectId) is var preferred && preferred != 0 && preferred != id)
                    throw new InvalidDataException("Source root ID disagrees with model ObjectId.");
            }
        }
    }

    private bool IsFragmentUnit(XElement node, XElement root)
    {
        if (!UnitRoot(node)) return false;
        if (ReferenceEquals(node, root)) return true;
        for (var parent = node.Parent; parent is not null; parent = parent.Parent)
        {
            if (parent.Name.LocalName != "PageBlock" || !UnitRoot(parent)) return false;
            if (ReferenceEquals(parent, root)) return true;
        }
        return false;
    }
    private bool UnitRoot(XElement root) =>
        (root.Name.Namespace == package.Options.Namespace || root.Name.Namespace == XNamespace.None || root.Name.NamespaceName == OfdConstants.Namespace || root.Name.NamespaceName == OfdConstants.StandardNamespace) &&
        root.Name.LocalName is "PageBlock" or "TextObject" or "ImageObject" or "PathObject" or "CompositeObject";

    private static bool IsClipPathId(XElement node)
    {
        var ns = node.Name.Namespace;
        return node.Name == ns + "Path" && node.Parent?.Name == ns + "Area" && node.Parent.Parent?.Name == ns + "Clip" && node.Parent.Parent.Parent?.Name == ns + "Clips";
    }

    internal string Claim(object model, string? preferred)
    {
        token.ThrowIfCancellationRequested();
        if (assignments.ContainsKey(model)) throw new InvalidDataException("A model definition occurs more than once in the output.");
        if (package.SourceIdentity?.TryOwner(model, out var owner) == true)
        {
            if (owner.Number == 0 && model is not OfdRawElement) return ClaimNew(model, preferred);
            return ClaimImported(owner, model, preferred);
        }
        if (model is OfdRawElement raw)
        {
            var root = OfdXmlEnvelope.Parse(raw.Xml);
            if (!UnitRoot(root))
            {
                var literal = root.Attribute("ID")?.Value;
                if (preferred is not null && preferred != literal) throw new InvalidDataException("Opaque raw root ID disagrees with the model.");
                assignments.Add(model, literal ?? string.Empty); return literal ?? string.Empty;
            }
        }
        return ClaimNew(model, preferred);
    }
    internal string ClaimLayer(OfdPage page, IEnumerable<OfdElement> members, string? preferred)
    {
        OfdSourceIdentity.Owner? owner = null;
        foreach (var member in members)
        {
            token.ThrowIfCancellationRequested(); budget.Structure.Reference();
            if (package.SourceIdentity?.TryLayer(member, out var candidate) != true) continue;
            if (owner is not null && !ReferenceEquals(owner, candidate)) throw new InvalidDataException("One output layer combines different source layer owners.");
            owner = candidate;
        }
        if (owner is not null)
        {
            if (!claimedLayers.Add(owner)) throw new InvalidDataException("One source layer was split into multiple output definitions.");
            return ClaimImported(owner, owner, preferred);
        }
        if (package.SourceIdentity?.TryEmptyLayer(page, preferred, out _) == true) return ClaimEmptyLayer(page, preferred);
        return ClaimNew(new object(), preferred);
    }
    internal string ClaimEmptyLayer(OfdPage page, string? preferred)
    {
        if (package.SourceIdentity?.TryEmptyLayer(page, preferred, out var owner) == true)
        {
            if (!claimedLayers.Add(owner)) throw new InvalidDataException("Source layer has multiple output definitions.");
            return ClaimImported(owner, owner, preferred);
        }
        return ClaimNew(new object(), preferred);
    }
    internal bool TryImageOwner(OfdImageElement image, out OfdSourceIdentity.Owner owner)
    {
        owner = null!;
        if (package.SourceIdentity is null) return false;
        return package.SourceIdentity.TryImage(image, out owner);
    }
    internal string ClaimImage(OfdImageElement image)
    {
        if (TryImageOwner(image, out var owner)) return ClaimImported(owner, owner, image.ResourceId);
        return ClaimNew(new object(), image.ResourceId);
    }
    private string ClaimImported(OfdSourceIdentity.Owner owner, object outputOwner, string? requested)
    {
        if ((requested ?? string.Empty) != owner.Id) throw new InvalidDataException("Imported ID remapping requires a complete reference remap.");
        if (!package.PreservedEntries.ContainsKey(owner.Entry)) throw new InvalidDataException("Imported source owner entry is missing.");
        if (!owner.ActiveDefinition) { assignments[outputOwner] = owner.Id; return owner.Id; }
        if (emitted.TryGetValue(owner.Number, out var existing))
        {
            if (!ReferenceEquals(existing, owner)) throw new InvalidDataException("Two output definitions claim one imported ID.");
        }
        else emitted.Add(owner.Number, owner);
        assignments[outputOwner] = owner.Id; return owner.Id;
    }
    private string ClaimNew(object model, string? preferred)
    {
        var id = OfdSourceIdentity.ParseId(preferred);
        if (id == 0)
        {
            if (model is OfdElement element)
            {
                var source = element switch { OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml, OfdPathElement path => path.SourceXml, OfdRawElement raw => raw.Xml, OfdCompositeElement composite => composite.SourceXml, _ => null };
                if (!string.IsNullOrWhiteSpace(source)) id = OfdSourceIdentity.ParseId(OfdXmlEnvelope.Parse(source!).Attribute("ID")?.Value);
            }
            if (id == 0) return Allocate(model);
        }
        if (emitted.ContainsKey(id) || fixedFragmentIds.Contains(id) || reserved.Contains(id) && (!fragmentOwners.TryGetValue(id, out var self) || !ReferenceEquals(self, model)))
            throw new InvalidDataException("New preferred numeric ID collides with a reserved or assigned ID.");
        emitted.Add(id, model); reserved.Add(id); AdvancePast(id);
        var result = preferred is not null && OfdSourceIdentity.ParseId(preferred) == id ? preferred : id.ToString(CultureInfo.InvariantCulture);
        assignments.Add(model, result); return result;
    }
    internal string Allocate() => Allocate(new object());
    private string Allocate(object model)
    {
        token.ThrowIfCancellationRequested();
        if (exhausted) throw new InvalidDataException("Document-unit ID space is exhausted.");
        var value = cursor; AdvancePast(value); emitted.Add(value, model); reserved.Add(value);
        var result = value.ToString(CultureInfo.InvariantCulture); assignments.Add(model, result); return result;
    }
    private void AdvancePast(long maximum)
    {
        if (maximum <= 0) return;
        MaxId = Math.Max(MaxId, maximum);
        if (maximum == long.MaxValue) { exhausted = true; return; }
        cursor = Math.Max(cursor, maximum + 1);
    }
    internal void ApplyRawRootId(OfdRawElement model, XElement root, string assigned)
    {
        if (!UnitRoot(root) || package.SourceIdentity?.TryOwner(model, out var owner) == true && !owner.ActiveDefinition) return;
        root.SetAttributeValue("ID", assigned);
    }

    internal void AssertNestedIds(OfdElement model, XElement root)
    {
        foreach (var attribute in root.Descendants().Attributes().Where(attribute => attribute.Name.LocalName == "ID"))
        {
            token.ThrowIfCancellationRequested(); var id = OfdSourceIdentity.ParseId(attribute.Value);
            if (id != 0 && emitted.ContainsKey(id) && !(originalReserved.Contains(id) && package.SourceIdentity?.TryOwner(model, out _) == true))
                throw new InvalidDataException("A retained nested ID collides with an emitted definition.");
            // An authenticated unchanged opaque source slot may already shadow
            // an active original number (for example a vendor-local Node ID).
            // AssertFragment certified every ID slot before serialization; new
            // definitions cannot acquire any of these reservations.
            // Fixed literal IDs were reserved before allocation. Never rewrite
            // unknown ID/reference pairs to repair an allocator collision.
            if (id != 0 && !reserved.Contains(id)) throw new InvalidDataException("Unplanned nested numeric ID.");
        }
    }
}
