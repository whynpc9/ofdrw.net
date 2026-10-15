using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Reader-owned baseline. All containers are private and copied on capture/clone.
internal sealed class OfdCustomTagSource
{
    internal const string EmrNamespace = "urn:ofdrw-net:custom-tags:emr";
    private readonly OfdDocumentPackage owner;
    private readonly string[] declarations, indexes, details;
    private readonly Dictionary<string, string> baseline;
    private readonly Dictionary<string, string?> digests;

    internal OfdCustomTagSource(OfdDocumentPackage owner, IReadOnlyList<string> declarations, IReadOnlyList<string> indexes,
        IReadOnlyList<string> details, OfdModelMaterialization material, CancellationToken token)
    {
        material.ReserveReadModel(512L + (declarations.Count + indexes.Count + details.Count) * 32L +
            (indexes.Count + details.Count + owner.CustomTags.Count) * 256L, token);
        this.owner = owner;
        this.declarations = declarations.ToArray(); this.indexes = indexes.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        this.details = details.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        baseline = new Dictionary<string, string>(owner.CustomTags);
        digests = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);
        foreach (var path in this.indexes.Concat(this.details))
        {
            token.ThrowIfCancellationRequested();
            if (!digests.ContainsKey(path)) digests.Add(path, owner.PreservedEntries.TryGetValue(path, out var bytes) ? Digest(bytes, material, token) : null);
        }
    }

    private OfdCustomTagSource(OfdCustomTagSource source, OfdDocumentPackage owner)
    {
        this.owner = owner; declarations = (string[])source.declarations.Clone(); indexes = (string[])source.indexes.Clone();
        details = (string[])source.details.Clone(); baseline = new Dictionary<string, string>(source.baseline);
        digests = new Dictionary<string, string?>(source.digests, StringComparer.OrdinalIgnoreCase);
    }

    internal IEnumerable<string> Declarations => declarations.Select(x => x);
    internal IEnumerable<string> IndexPaths => indexes.Select(x => x);
    internal IEnumerable<string> DetailPaths => details.Select(x => x);
    internal bool TryOriginalValue(string key, out string value) => baseline.TryGetValue(key, out value!);
    internal bool IsUnchanged(OfdDocumentPackage package) => baseline.Count == package.CustomTags.Count &&
        baseline.All(pair => package.CustomTags.TryGetValue(pair.Key, out var value) && pair.Value == value);

    internal void Validate(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        if (!ReferenceEquals(owner, package)) throw new InvalidDataException("Foreign custom tag source owner.");
        if (declarations.Length > 0)
        {
            var path = package.DocumentEntryPath;
            if (path is null || !package.PreservedEntries.TryGetValue(path, out var document)) throw new InvalidDataException("Custom tag document source missing.");
            var xml = material.ReadXml(path, document, token);
            var actual = xml.Root!.Elements(xml.Root.Name.Namespace + "CustomTags").Select(node => OfdXmlEnvelope.Capture(node, material, token));
            if (!declarations.SequenceEqual(actual)) throw new InvalidDataException("Custom tag declarations changed behind their reader-owned slots.");
        }
        foreach (var entry in digests)
        {
            token.ThrowIfCancellationRequested();
            var current = package.PreservedEntries.TryGetValue(entry.Key, out var bytes) ? Digest(bytes, material, token) : null;
            if (current != entry.Value) throw new InvalidDataException("Custom tag source changed behind its reader-owned slots.");
        }
    }

    internal OfdCustomTagSource Clone(OfdDocumentPackage source, OfdDocumentPackage destination, OfdModelMaterialization material, CancellationToken token)
    {
        if (!ReferenceEquals(owner, source)) throw new InvalidDataException("Foreign custom tag source owner.");
        material.ChargeCopy(512L + (declarations.Length + indexes.Length + details.Length) * 16L + (baseline.Count + digests.Count) * 128L, token);
        return new OfdCustomTagSource(this, destination);
    }

    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("customTagOwner", owner, true);
        sink.Item("customTagDeclarations", declarations, true); sink.Item("count", declarations.Length);
        foreach (var declaration in declarations) { sink.Item("declaration", declaration); sink.Xml(declaration); }
        sink.Item("customTagIndexes", indexes, true); sink.Item("count", indexes.Length);
        foreach (var path in indexes) { sink.Item("index", path); sink.Path(path); }
        sink.Item("customTagDetails", details, true); sink.Item("count", details.Length);
        foreach (var path in details) { sink.Item("detail", path); sink.Path(path); }
        sink.Item("customTagBaseline", baseline, true); sink.Item("count", baseline.Count);
        foreach (var pair in baseline) { sink.Item("key", pair.Key); sink.Item("value", pair.Value); }
        sink.Item("customTagDigests", digests, true); sink.Item("count", digests.Count);
        foreach (var pair in digests) { sink.Item("path", pair.Key); sink.Path(pair.Key); sink.Item("digest", pair.Value); }
    }

    private static string Digest(byte[] bytes, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(512, token);
        var hash = material.HashBytes(bytes, token);
        var storage = material.IsControlledRead ? material.ReserveReadStorage(190, token) : null;
        var result = BitConverter.ToString(hash);
        if (storage is not null) material.AdoptReadString(result, storage, token);
        return result;
    }
}
