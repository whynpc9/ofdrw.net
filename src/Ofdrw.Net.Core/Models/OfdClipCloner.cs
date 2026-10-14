using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal static class OfdClipCloner
{
    internal static void Clone(OfdElement source, OfdElement target, IDictionary<object, object>? mapping,
        OfdModelMaterialization material, CancellationToken token, bool clonePayloads, OfdPayloadCloneSession? payloads)
    {
        var map = mapping ?? new Dictionary<object, object>(OfdReferenceIdentity.Instance);
        var owner = source.ClipOwnership; var copy = target.ClipOwnership;
        owner.Validate(source.ClippingXml, material, token);
        material.ChargeCopy(256, token); map.Add(owner, copy);
        copy.Value = Collection(owner.Value); copy.Original = Collection(owner.Original); copy.Raw = owner.Raw;
        copy.Baseline = owner.Baseline?.Remap(map, material, token);
        OfdClipCollection? Collection(OfdClipCollection? value)
        {
            if (value is null) return null;
            if (map.TryGetValue(value, out var prior)) return (OfdClipCollection)prior;
            material.ChargeCopy(256L + 32L * value.Items.Count, token);
            var result = new OfdClipCollection { SourceXml = value.SourceXml }; map.Add(value, result); map.Add(value.Items, result.Items);
            foreach (var clip in value.Items)
            {
                token.ThrowIfCancellationRequested(); material.ChargeCopy(256L + 32L * clip.Areas.Count, token);
                var c = new OfdClip { SourceXml = clip.SourceXml, SourceSlot = clip.SourceSlot }; map.Add(clip, c); map.Add(clip.Areas, c.Areas);
                foreach (var area in clip.Areas)
                {
                    material.ChargeCopy(256L + 48, token);
                    var a = new OfdClipArea { SourceXml = area.SourceXml, SourceSlot = area.SourceSlot, Transform = area.Transform?.ToArray(),
                        Path = (OfdPathElement)OfdModelCloner.CloneElement(area.Path, clonePayloads: clonePayloads, material: material, cancellationToken: token, mapping: map, payloads: payloads) };
                    a.OriginalPathXml = area.OriginalPathXml;
                    if (area.OriginalPath is not null)
                        a.OriginalPath = map.TryGetValue(area.OriginalPath, out var originalPath) ? (OfdPathElement)originalPath :
                            (OfdPathElement)OfdModelCloner.CloneElement(area.OriginalPath, clonePayloads: clonePayloads, material: material, cancellationToken: token, mapping: map, payloads: payloads);
                    map.Add(area, a); c.Areas.Add(a);
                }
                result.Items.Add(c);
            }
            return result;
        }
    }
}
