using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal static class OfdAnnotationCloner
{
    internal static void Clone(OfdPage source, OfdPage destination, IDictionary<object, object> mapping,
        bool payloads, OfdModelMaterialization material, CancellationToken token, OfdPayloadCloneSession? payloadSession = null)
    {
        source.AnnotationOwnership.Validate(source, material, token);
        var original = Collection(source.AnnotationOwnership.Original);
        var baselineCollection = Collection(source.AnnotationOwnership.BaselineCollection);
        var value = Collection(source.Annotations);
        var sourceOwner = source.AnnotationOwnership; var owner = destination.AnnotationOwnership;
        material.ChargeCopy(256L + sourceOwner.Entries.Count * 128L, token);
        foreach (var pair in sourceOwner.Entries) owner.Entries.Add(pair.Key, pair.Value);
        foreach (var path in sourceOwner.UnsafeDeletionEntries) { material.ChargeCopy(64, token); owner.UnsafeDeletionEntries.Add(path); }
        owner.Restore(value, original, sourceOwner.Baseline?.Remap(mapping, material, token),
            sourceOwner.ProjectionBaseline?.Remap(mapping, material, token), sourceOwner.HasUnclaimedSource, baselineCollection);
        mapping.Add(sourceOwner, owner);
        OfdAnnotationCollection? Collection(OfdAnnotationCollection? collection)
        {
            if (collection is null) return null;
            if (mapping.TryGetValue(collection, out var existing)) return (OfdAnnotationCollection)existing;
            material.ChargeCopy(128L + 16L * collection.Items.Count, token);
            var result = new OfdAnnotationCollection(); mapping.Add(collection, result); mapping.Add(collection.Items, result.Items);
            foreach (var annotation in collection.Items)
            {
                token.ThrowIfCancellationRequested();
                if (mapping.TryGetValue(annotation, out var mapped)) { result.Items.Add((OfdAnnotation)mapped); continue; }
                material.ChargeCopy(512L + 16L * annotation.Appearance.Elements.Count, token);
                var appearance = annotation.Appearance;
                var copy = new OfdAnnotation
                {
                    Id = annotation.Id, Type = annotation.Type, Creator = annotation.Creator, LastModifiedDate = annotation.LastModifiedDate,
                    Subtype = annotation.Subtype, Visible = annotation.Visible, Print = annotation.Print, NoZoom = annotation.NoZoom,
                    NoRotate = annotation.NoRotate, ReadOnly = annotation.ReadOnly, Remark = annotation.Remark,
                    AttachmentId = annotation.AttachmentId, SourceXml = annotation.SourceXml,
                    Appearance = new OfdAnnotationAppearance { XMillimeters = appearance.XMillimeters, YMillimeters = appearance.YMillimeters,
                        WidthMillimeters = appearance.WidthMillimeters, HeightMillimeters = appearance.HeightMillimeters, SourceXml = appearance.SourceXml }
                };
                if (appearance.Transform is { } transform) { material.ChargeCopy(transform.LongLength * sizeof(double), token); copy.Appearance.Transform = transform.ToArray(); }
                mapping.Add(annotation, copy); mapping.Add(appearance, copy.Appearance); mapping.Add(appearance.Elements, copy.Appearance.Elements);
                foreach (var element in appearance.Elements) copy.Appearance.Elements.Add(OfdModelCloner.CloneElement(element, clonePayloads: payloads, material: material, cancellationToken: token, mapping: mapping, payloads: payloadSession));
                if (annotation.Source is { } sourceSlot)
                {
                    material.ChargeCopy(128, token);
                    copy.Source = new OfdAnnotationSource(sourceSlot.Entry, sourceSlot.Slot, sourceSlot.ProjectionStart, sourceSlot.ProjectionCount, sourceSlot.Primitives.Remap(mapping, material, token), sourceSlot.Namespace);
                }
                result.Items.Add(copy);
            }
            return result;
        }
    }
}
