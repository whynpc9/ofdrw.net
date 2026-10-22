using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Definition walks count each owner once; placements are reference edges.
internal static class OfdCompositeGraph
{
    internal static IEnumerable<OfdElement> Definitions(OfdDocumentPackage package)
    {
        foreach (var page in package.Pages) foreach (var element in OfdAnnotationContract.Definitions(page)) yield return element;
        foreach (var resource in package.CompositeResources?.Items ?? Enumerable.Empty<OfdCompositeResource>())
            foreach (var element in resource.Elements) yield return element;
    }

    internal static void Validate(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token, bool rejectRepeatedOwners = false)
    {
        material.ChargeCopy(512, token);
        var resources = new HashSet<OfdCompositeResource>();
        var owners = new HashSet<OfdElement>();
        var resourceOwners = new HashSet<OfdElement>();
        foreach (var resource in package.CompositeResources?.Items ?? Enumerable.Empty<OfdCompositeResource>())
        {
            token.ThrowIfCancellationRequested(); material.ChargeCopy(256, token); material.Structure.Node(1);
            if (!resources.Add(resource)) throw new InvalidDataException("Repeated composite resource owner.");
            if (!OfdBoxParser.IsWritablePositiveSide(resource.WidthMillimeters) || !OfdBoxParser.IsWritablePositiveSide(resource.HeightMillimeters) || double.IsInfinity(resource.WidthMillimeters) || double.IsInfinity(resource.HeightMillimeters))
                throw new InvalidDataException("Composite dimensions must be finite and positive.");
            foreach (var element in resource.Elements)
            {
                token.ThrowIfCancellationRequested(); material.ChargeCopy(64, token); material.ChargeWork(64, token);
                if (!resourceOwners.Add(element)) throw new InvalidDataException("Repeated composite child owner.");
            }
        }
        foreach (var element in Definitions(package))
        {
            token.ThrowIfCancellationRequested(); material.ChargeCopy(64, token); material.ChargeWork(64, token);
            if (!owners.Add(element) && (rejectRepeatedOwners || resourceOwners.Contains(element))) throw new InvalidDataException("Repeated graphic owner in composite graph.");
            if (element is OfdCompositeElement placement)
            {
                material.Structure.Reference(); material.StoreMatrix(placement.Transform);
                if (placement.Resource is null || !resources.Contains(placement.Resource)) throw new InvalidDataException("Composite placement has a foreign or missing resource.");
                if (new[] { placement.XMillimeters, placement.YMillimeters, placement.WidthMillimeters, placement.HeightMillimeters }.Any(v => double.IsNaN(v) || double.IsInfinity(v)) || !OfdBoxParser.IsWritablePositiveSide(placement.WidthMillimeters) || !OfdBoxParser.IsWritablePositiveSide(placement.HeightMillimeters)) throw new InvalidDataException("Composite placement requires a positive Boundary.");
            }
        }
        var done = new HashSet<OfdCompositeResource>(); var active = new HashSet<OfdCompositeResource>();
        var depths = new Dictionary<OfdCompositeResource, int>();
        foreach (var root in resources)
        {
            material.ChargeCopy(128, token);
            var stack = new Stack<(OfdCompositeResource Resource, int Next)>(); stack.Push((root, -1));
            while (stack.Count > 0)
            {
                token.ThrowIfCancellationRequested(); material.ChargeWork(64, token); material.Structure.Depth(stack.Count);
                var frame = stack.Pop();
                if (frame.Next < 0)
                {
                    if (done.Contains(frame.Resource)) continue;
                    material.ChargeCopy(128, token);
                    if (!active.Add(frame.Resource)) throw new InvalidDataException("Cyclic composite resource graph.");
                    frame.Next = 0;
                }
                if (frame.Next == frame.Resource.Elements.Count)
                {
                    var depth = 1;
                    foreach (var nested in frame.Resource.Elements.OfType<OfdCompositeElement>())
                    { token.ThrowIfCancellationRequested(); material.ChargeWork(64, token); depth = Math.Max(depth, checked(1 + depths[nested.Resource])); }
                    material.Structure.Depth(depth); material.ChargeCopy(128, token);
                    depths.Add(frame.Resource, depth); active.Remove(frame.Resource); done.Add(frame.Resource); continue;
                }
                material.ChargeCopy(128, token); material.Structure.Reference();
                stack.Push((frame.Resource, frame.Next + 1));
                if (frame.Resource.Elements[frame.Next] is OfdCompositeElement child) stack.Push((child.Resource, -1));
            }
        }
    }

    internal static void CloneResources(OfdDocumentPackage source, OfdDocumentPackage destination,
        IDictionary<object, object> map, OfdModelMaterialization material, CancellationToken token, OfdPayloadCloneSession? payloads)
    {
        var collection = source.CompositeResources ?? source.OriginalCompositeResources;
        if (collection is null) return;
        source.ValidateCompositeOwnership(material, token);
        material.ChargeCopy(512L + collection.Items.Count * 256L, token);
        var copiedCollection = new OfdCompositeResourceCollection();
        destination.CompositeValue = source.CompositeResources is null ? null : copiedCollection;
        destination.OriginalCompositeResources = source.OriginalCompositeResources is null ? null : copiedCollection;
        map.Add(collection, copiedCollection); map.Add(collection.Items, copiedCollection.Items);
        foreach (var resource in collection.Items)
        {
            token.ThrowIfCancellationRequested(); material.ChargeCopy(512L + 32L * resource.Elements.Count, token);
            var clone = new OfdCompositeResource { Id = resource.Id, WidthMillimeters = resource.WidthMillimeters,
                HeightMillimeters = resource.HeightMillimeters, SourceXml = resource.SourceXml, SourceEntryPath = resource.SourceEntryPath,
                SourceSlot = resource.SourceSlot, SourceContainerSupported = resource.SourceContainerSupported };
            map.Add(resource, clone); map.Add(resource.Elements, clone.Elements); copiedCollection.Items.Add(clone);
        }
        foreach (var resource in collection.Items)
        {
            var clone = (OfdCompositeResource)map[resource];
            foreach (var element in resource.Elements) clone.Elements.Add(OfdModelCloner.CloneElement(element, material: material, cancellationToken: token, mapping: map, payloads: payloads));
        }
        material.ChargeCopy(32L * source.OriginalCompositeOrder.Length, token);
        destination.OriginalCompositeOrder = source.OriginalCompositeOrder.Select(r => (OfdCompositeResource)map[r]).ToArray();
        destination.CompositeBaseline = source.CompositeBaseline?.Remap(map, material, token);
    }
}
