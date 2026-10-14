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
    private sealed class SignatureRemoval(XElement node, XName[] names, int[] positions)
    {
        internal readonly XElement Node = node;
        internal readonly XName[] Names = names;
        internal readonly int[] Positions = positions;
    }
    private readonly List<SignatureRemoval> removedSignatures = new();
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
            var location = SignatureLocation(node, token);
            removedSignatures.Add(new SignatureRemoval(node, location.Names, location.Positions));
        }
    }
    private (XName[] Names, int[] Positions) SignatureLocation(XElement node, CancellationToken token)
    {
        var depth = 0;
        foreach (var ancestor in node.AncestorsAndSelf()) { context.Material.ChargeWork(16, token); depth++; }
        context.Material.ChargeCopy(checked(512L + 256L * depth), token);
        var chain = node.AncestorsAndSelf().Reverse().ToArray(); var names = new XName[depth]; var positions = new int[depth];
        for (var i = 0; i < depth; i++)
        {
            names[i] = chain[i].Name;
            if (chain[i].Parent is null) continue;
            foreach (var sibling in chain[i].Parent!.Elements())
            {
                context.Material.ChargeWork(16, token);
                if (ReferenceEquals(sibling, chain[i])) break;
                positions[i]++;
            }
        }
        return (names, positions);
    }
    internal void ApplyActualSignatureRemovals(string entry, XDocument expected, CancellationToken token)
    {
        if (entry != "OFD.xml" || removedSignatures.Count == 0) return;
        var root = expected.Root ?? throw new InvalidDataException("Signature source root is absent."); var ns = root.Name.Namespace;
        context.Material.ChargeCopy(checked(512L + 256L * removedSignatures.Count), token);
        // Collect first: removing one declaration must not change the positions
        // against which the subsequent actual prune dispositions are checked.
        var matches = new List<XElement>();
        foreach (var source in root.Elements(ns + "DocBody").Elements(ns + "Signatures"))
        {
            var location = SignatureLocation(source, token); SignatureRemoval? match = null;
            foreach (var actual in removedSignatures)
            {
                context.Material.ChargeWork(checked(32L + 16L * location.Names.Length), token);
                if (!actual.Names.SequenceEqual(location.Names) || !actual.Positions.SequenceEqual(location.Positions)) continue;
                ChargeTree(source); ChargeTree(actual.Node);
                if (!XNode.DeepEquals(source, actual.Node)) throw new InvalidDataException("Actual signature prune changed its source declaration.");
                match = actual; break;
            }
            if (match is not null) matches.Add(source);
        }
        if (matches.Count != removedSignatures.Count) throw new InvalidDataException("Actual signature dispositions do not bind exact source ancestor positions.");
        foreach (var source in matches) { context.Material.ChargeWork(16, token); source.Remove(); }
        void ChargeTree(XElement element)
        {
            foreach (var node in element.DescendantsAndSelf())
            {
                context.Material.ChargeWork(64, token);
                foreach (var attribute in node.Attributes()) context.Material.ChargeWork(checked(32L + 2L * attribute.Value.Length), token);
                foreach (var child in node.Nodes()) context.Material.ChargeWork(child switch
                {
                    XText text => checked(16L + 2L * text.Value.Length),
                    XComment comment => checked(16L + 2L * comment.Value.Length),
                    XProcessingInstruction pi => checked(32L + 2L * (pi.Target.Length + pi.Data.Length)),
                    _ => 16
                }, token);
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
