using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>The custom-tag shapes this library writes and therefore recognizes.</summary>
internal enum KnownCustomTagShape
{
    None,
    /// <summary>Neutral channel: TypeID=ofdrw-kv, payload root in the declared namespace.</summary>
    KeyValue,
    /// <summary>
    /// Legacy EMR channel as historically written by this library: the declared NameSpace differs from the payload
    /// root, which lives in the OFD namespace. Recognized as a library-written shape, not as a verified standard tag.
    /// </summary>
    LegacyEmr
}

/// <summary>One located result of following the known custom-tag chain.</summary>
internal sealed class KnownCustomTagFinding
{
    internal KnownCustomTagFinding(PreflightOutcome outcome, string reason, string entry, XElement? node, string observed,
        KnownCustomTagShape shape = KnownCustomTagShape.None, string? target = null)
    { Outcome = outcome; Reason = reason; Entry = entry; Node = node; Observed = observed; Shape = shape; Target = target; }
    internal PreflightOutcome Outcome { get; }
    internal string Reason { get; }
    /// <summary>The XML entry that contains <see cref="Node"/>.</summary>
    internal string Entry { get; }
    internal XElement? Node { get; }
    internal string Observed { get; }
    internal KnownCustomTagShape Shape { get; }
    internal string? Target { get; }
    /// <summary>A known reference whose target is absent from, or outside, the package.</summary>
    internal bool TargetUnavailable => Outcome == PreflightOutcome.Fail;
}

/// <summary>
/// Follows Document → CustomTags index → CustomTag/FileLoc → payload for the shapes this library writes. Every edge is
/// read by XML field semantics (element text or the FileLoc attribute) and resolved with package path rules; tag values
/// are plain text and are never followed. Anything outside the known shapes is reported Unknown, never whitelisted.
/// Neither shape has been verified against GB/T 33190's custom-tag requirements.
/// </summary>
internal static class KnownCustomTagScope
{
    private static readonly XName TypeId = "TypeID", NameSpace = "NameSpace", FileLoc = "FileLoc", Key = "Key", Value = "Value";

    /// <summary>A Document child CustomTags declaration in the document namespace whose only content is its location text.</summary>
    internal static bool IsDeclaration(XElement node, XElement documentRoot) =>
        node.Parent == documentRoot && node.Name == documentRoot.Name.Namespace + "CustomTags" && !node.HasElements &&
        !node.Attributes().Any(x => !x.IsNamespaceDeclaration) && node.Nodes().All(x => x is XText);

    internal static KnownCustomTagShape Classify(XElement record) =>
        OfdCustomTagChannels.Declared(record) switch
        {
            OfdCustomTagChannel.KeyValue => KnownCustomTagShape.KeyValue,
            OfdCustomTagChannel.Emr => KnownCustomTagShape.LegacyEmr,
            _ => KnownCustomTagShape.None
        };

    /// <summary>
    /// Traverses every CustomTags declaration of one document. <paramref name="read"/> returns the parsed entry or null
    /// for unreadable XML and enforces the caller's budgets (it may throw the caller's budget exception);
    /// <paramref name="visit"/> charges one reference; <paramref name="reserved"/> reports entries already owning another
    /// role (OFD root, pages, resources), which a tag edge must not reuse.
    /// </summary>
    internal static IReadOnlyList<KnownCustomTagFinding> Traverse(OfdPackageArchive archive, string documentEntry,
        XElement documentRoot, Func<string, XDocument?> read, Action visit, Func<string, bool> reserved, CancellationToken token)
    {
        var findings = new List<KnownCustomTagFinding>();
        var ns = documentRoot.Name.Namespace;
        var indexes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var payloads = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        bool Taken(string entry) => string.Equals(entry, documentEntry, StringComparison.OrdinalIgnoreCase) ||
            indexes.Contains(entry) || payloads.Contains(entry) || reserved(entry);
        void Unknown(string reason, string entry, XElement? node, string observed = "") =>
            findings.Add(new KnownCustomTagFinding(PreflightOutcome.Unknown, reason, entry, node, observed));

        foreach (var declaration in documentRoot.Elements().Where(x => x.Name.LocalName == "CustomTags"))
        {
            token.ThrowIfCancellationRequested();
            if (!IsDeclaration(declaration, documentRoot))
            { Unknown("UnrecognizedCustomTagDeclaration", documentEntry, declaration); continue; }
            var index = Reference(archive, documentEntry, declaration, declaration.Value, findings, visit, KnownCustomTagShape.None);
            if (index == null) continue;
            if (indexes.Contains(index)) { Unknown("DuplicateCustomTagTarget", documentEntry, declaration, "Target=" + index); continue; }
            if (Taken(index)) { Unknown("CustomTagRoleConflict", documentEntry, declaration, "Target=" + index); continue; }
            indexes.Add(index);
            var root = read(index)?.Root;
            if (root == null) { Unknown("UnreadableCustomTagIndex", documentEntry, declaration, "Target=" + index); continue; }
            if (root.Name != ns + "CustomTags" || root.Attributes().Any(x => !x.IsNamespaceDeclaration) || Opaque(root))
            { Unknown("UnrecognizedCustomTagIndex", index, root); continue; }
            foreach (var record in root.Elements())
            {
                token.ThrowIfCancellationRequested();
                Record(archive, index, record, ns, findings, read, visit, Taken, payloads, token);
            }
        }
        return findings.AsReadOnly();
    }

    private static void Record(OfdPackageArchive archive, string index, XElement record, XNamespace ns,
        List<KnownCustomTagFinding> findings, Func<string, XDocument?> read, Action visit, Func<string, bool> taken,
        HashSet<string> payloads, CancellationToken token)
    {
        void Unknown(string reason, string entry, XElement node, string observed = "") =>
            findings.Add(new KnownCustomTagFinding(PreflightOutcome.Unknown, reason, entry, node, observed));
        if (record.Name != ns + "CustomTag") { Unknown("UnrecognizedCustomTagRecord", index, record); return; }
        var shape = Classify(record);
        if (shape == KnownCustomTagShape.None)
        {
            Unknown("UnknownCustomTagSchema", index, record, "TypeID=" + (string?)record.Attribute(TypeId) +
                "; NameSpace=" + (string?)record.Attribute(NameSpace));
            return;
        }
        // SchemaLoc and any other field are references or semantics outside the known shape.
        if (record.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != TypeId && x.Name != NameSpace && x.Name != FileLoc) ||
            record.Elements().Any(x => x.Name != ns + "FileLoc") || Opaque(record))
        { Unknown("UnrecognizedCustomTagRecord", index, record); return; }
        var elements = record.Elements(ns + "FileLoc").ToList();
        var attribute = record.Attribute(FileLoc);
        if (elements.Count + (attribute == null ? 0 : 1) > 1) { Unknown("AmbiguousFileLoc", index, record); return; }
        var location = elements.FirstOrDefault();
        if (location != null && (location.HasElements || location.Attributes().Any(x => !x.IsNamespaceDeclaration)))
        { Unknown("UnsupportedReferenceShape", index, location); return; }
        var node = location ?? record;
        var payload = Reference(archive, index, node, location?.Value ?? attribute?.Value, findings, visit, shape);
        if (payload == null) return;
        if (payloads.Contains(payload)) { Unknown("DuplicateCustomTagTarget", index, node, "Target=" + payload); return; }
        if (taken(payload)) { Unknown("CustomTagRoleConflict", index, node, "Target=" + payload); return; }
        payloads.Add(payload);
        var root = read(payload)?.Root;
        if (root == null) { Unknown("UnreadableCustomTagPayload", index, node, "Target=" + payload); return; }
        var tags = PayloadTags(root, shape, ns, token);
        if (tags < 0) { Unknown("UnrecognizedCustomTagPayload", index, node, "Target=" + payload + "; Shape=" + shape); return; }
        findings.Add(new KnownCustomTagFinding(PreflightOutcome.Pass, "KnownCustomTagPayload", index, node,
            "Shape=" + shape + "; Target=" + payload + "; tags=" + tags, shape, payload));
    }

    // Plain Key/Value tags only: no other elements, attributes, text or nested content. Returns -1 when unrecognized.
    private static int PayloadTags(XElement root, KnownCustomTagShape shape, XNamespace ns, CancellationToken token)
    {
        var channel = shape == KnownCustomTagShape.KeyValue ? OfdCustomTagChannel.KeyValue : OfdCustomTagChannel.Emr;
        if (OfdCustomTagChannels.Payload(root, ns) != channel || root.Attributes().Any(x => !x.IsNamespaceDeclaration) || Opaque(root))
            return -1;
        var tag = OfdCustomTagChannels.TagName(channel, ns);
        var count = 0;
        foreach (var element in root.Elements())
        {
            token.ThrowIfCancellationRequested();
            if (element.Name != tag || element.Nodes().Any() || string.IsNullOrWhiteSpace((string?)element.Attribute(Key)) ||
                element.Attribute(Value) == null ||
                element.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != Key && x.Name != Value)) return -1;
            count++;
        }
        return count;
    }

    private static bool Opaque(XElement element) =>
        element.Nodes().Any(x => x is XText text ? !string.IsNullOrWhiteSpace(text.Value) : x is not XElement && x is not XComment);

    private static string? Reference(OfdPackageArchive archive, string containing, XElement node, string? value,
        List<KnownCustomTagFinding> findings, Action visit, KnownCustomTagShape shape)
    {
        visit();
        void Add(PreflightOutcome outcome, string reason, string observed) =>
            findings.Add(new KnownCustomTagFinding(outcome, reason, containing, node, observed, shape));
        value = value?.Trim() ?? string.Empty;
        if (value.Length == 0) { Add(PreflightOutcome.Fail, "EmptyReference", "Reference is empty"); return null; }
        if (value.Contains(":") || value.StartsWith("\\", StringComparison.Ordinal) || value.StartsWith("//", StringComparison.Ordinal))
        { Add(PreflightOutcome.Fail, "ExternalReference", "External location (redacted)"); return null; }
        if (value.IndexOfAny(new[] { '?', '#', '%', '\\' }) >= 0 || value.Any(char.IsControl))
        { Add(PreflightOutcome.Unknown, "AmbiguousReference", "Ambiguous location (redacted)"); return null; }
        string target;
        try { target = OfdPackagePath.Resolve(containing, value); }
        catch (InvalidDataException) { Add(PreflightOutcome.Fail, "RootEscape", "Reference escapes the package root"); return null; }
        if (target.Length == 0 || !archive.Contains(target))
        { Add(PreflightOutcome.Fail, "MissingTarget", "Target=" + target); return null; }
        if (shape == KnownCustomTagShape.None)
            findings.Add(new KnownCustomTagFinding(PreflightOutcome.Pass, "KnownCustomTagIndex", containing, node, "Target=" + target, shape, target));
        return target;
    }
}
