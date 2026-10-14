using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal enum OfdOwnedPruneReason
{
    DeletedPage, DeletedPageResource, UnusedTemplate, DeletedAnnotation,
    UnreachableResourcePayload, InvalidatedSignature
}

// Recorded at the actual removal edge, never inferred from a missing final key.
internal sealed class OfdOwnedPruneDispositions
{
    private readonly OfdControlledEditContext context;
    private readonly Dictionary<string, (OfdOwnedPruneReason Reason, byte[] Bytes)> removed = new(StringComparer.OrdinalIgnoreCase);
    private readonly List<XElement> removedSignatures = new();
    private OfdOwnedPruneDispositions(OfdControlledEditContext context) { this.context = context; }
    internal static OfdOwnedPruneDispositions Begin(OfdControlledEditContext context, CancellationToken token)
    { context.Material.ChargeCopy(1024, token); return new(context); }
    internal void RecordRemoval(string path, byte[] actualRemoved, OfdOwnedPruneReason reason, CancellationToken token)
    {
        context.AssertActive(); context.Material.ChargeCopy(checked(512L + 4L * path.Length), token);
        context.Material.ChargeWork(checked(16L + 2L * path.Length), token);
        if (removed.ContainsKey(path)) throw new InvalidDataException("Repeated actual prune disposition.");
        removed.Add(path, (reason, actualRemoved));
    }
    internal void RecordActualSignatureDeclarations(List<XElement> actual, CancellationToken token)
    {
        context.AssertActive(); context.Material.ChargeCopy(checked(512L + 256L * actual.Count), token);
        foreach (var node in actual)
        {
            context.Material.ChargeWork(16, token);
            if (node.Parent?.Name != node.Name.Namespace + "DocBody" || node.Name != node.Name.Namespace + "Signatures")
                throw new InvalidDataException("Signature disposition lacks its actual declaration owner.");
            removedSignatures.Add(node);
        }
    }
    internal bool MatchesActualRemovedSignature(XElement source, CancellationToken token)
    {
        if (source.Name != source.Name.Namespace + "Signatures" || source.Parent?.Name != source.Name.Namespace + "DocBody") return false;
        foreach (var actual in removedSignatures)
        {
            ChargeTree(source); ChargeTree(actual);
            if (XNode.DeepEquals(source, actual)) return true;
        }
        return false;
        void ChargeTree(XElement root)
        {
            foreach (var node in root.DescendantsAndSelf())
            {
                context.Material.ChargeWork(64, token);
                foreach (var attribute in node.Attributes()) context.Material.ChargeWork(checked(32L + 2L * attribute.Value.Length), token);
                foreach (var text in node.Nodes().OfType<XText>()) context.Material.ChargeWork(checked(16L + 2L * text.Value.Length), token);
            }
        }
    }
    internal void RequireActualRemoval(string path, byte[] actualOriginal, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (!removed.TryGetValue(path, out var record) || !ReferenceEquals(record.Bytes, actualOriginal))
            throw new InvalidDataException("Entry lacks an exact actual pruner disposition.");
    }
}
