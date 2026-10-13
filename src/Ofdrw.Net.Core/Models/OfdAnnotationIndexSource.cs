using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Reader-owned immutable scalar baselines. Clones own their containers and owner identity.
internal sealed class OfdAnnotationIndexSource
{
    internal object? DocumentOwner;
    internal readonly List<string> Declarations = new();
    internal readonly Dictionary<string, string?> IndexDigests = new(StringComparer.OrdinalIgnoreCase);
    internal void Capture(string path, byte[]? bytes, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(256, token); IndexDigests.Add(path, bytes is null ? null : Digest(bytes, material, token));
    }
    internal void Validate(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        if (!ReferenceEquals(DocumentOwner, package)) throw new InvalidDataException("Foreign annotation index source owner.");
        foreach (var entry in IndexDigests)
        {
            token.ThrowIfCancellationRequested();
            package.PreservedEntries.TryGetValue(entry.Key, out var bytes);
            if ((bytes is null ? null : Digest(bytes, material, token)) != entry.Value)
                throw new InvalidDataException("Annotation index changed behind its reader-owned source slots.");
        }
        if (Declarations.Count == 0) return;
        var path = package.DocumentEntryPath;
        if (path is null || !package.PreservedEntries.TryGetValue(path, out var docBytes)) throw new InvalidDataException("Annotation document source is missing.");
        var xml = material.ReadXml(path, docBytes, token);
        var current = xml.Root!.Elements().Where(n => n.Name.LocalName == "Annotations").Select(n => OfdXmlEnvelope.Capture(n, material, token));
        if (!Declarations.SequenceEqual(current)) throw new InvalidDataException("Annotation declaration source changed.");
        var fragments = new List<string>();
        foreach (var fragment in package.PreservedDocumentElements)
        {
            token.ThrowIfCancellationRequested(); material.GuardFragmentXml(fragment, token);
            material.ChargeCopy(2L * fragment.Length + 128, token); material.ChargeWork(2L * fragment.Length, token);
            var node = OfdXmlEnvelope.Parse(fragment);
            if (node.Name.LocalName == "Annotations") fragments.Add(OfdXmlEnvelope.Capture(node, material, token));
        }
        if (!Declarations.SequenceEqual(fragments)) throw new InvalidDataException("Annotation declaration projection changed.");
    }
    internal OfdAnnotationIndexSource Clone(OfdDocumentPackage source, OfdDocumentPackage destination, OfdModelMaterialization material, CancellationToken token)
    {
        if (!ReferenceEquals(DocumentOwner, source)) throw new InvalidDataException("Foreign annotation index source owner.");
        material.ChargeCopy(256L + Declarations.Count * 16L + IndexDigests.Count * 128L, token);
        var clone = new OfdAnnotationIndexSource { DocumentOwner = destination }; clone.Declarations.AddRange(Declarations);
        foreach (var entry in IndexDigests) { token.ThrowIfCancellationRequested(); clone.IndexDigests.Add(entry.Key, entry.Value); }
        return clone;
    }
    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("documentOwner", DocumentOwner, true); sink.Item("declarations", Declarations, true); sink.Item("count", Declarations.Count);
        foreach (var declaration in Declarations) { sink.Item("declaration", declaration); sink.Xml(declaration); }
        sink.Item("indexDigests", IndexDigests, true); sink.Item("count", IndexDigests.Count);
        foreach (var entry in IndexDigests) { sink.Item("indexPath", entry.Key); sink.Path(entry.Key); sink.Item("indexDigest", entry.Value); }
    }
    private static string Digest(byte[] value, OfdModelMaterialization material, CancellationToken token)
    { material.ChargeCopy(256, token); return BitConverter.ToString(material.HashBytes(value, token)); }
}
