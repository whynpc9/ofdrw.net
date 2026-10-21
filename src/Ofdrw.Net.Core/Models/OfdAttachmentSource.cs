using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// The descriptor remains in PreservedEntries. Reader-owned object slots, not
// editable IDs/names, identify its nodes. No public XML escape hatch is added.
internal sealed class OfdAttachmentSource
{
    internal sealed class Original(int ordinal, string fileLoc, string? payloadPath, bool external, Dictionary<string, string?> values)
    {
        internal readonly int Ordinal = ordinal;
        internal readonly string FileLoc = fileLoc;
        internal readonly string? PayloadPath = payloadPath;
        internal readonly bool External = external;
        internal readonly Dictionary<string, string?> Values = values;
    }

    private readonly OfdDocumentPackage owner;
    internal readonly string DocumentPath, DescriptorPath, Declaration;
    internal readonly XName RootName;
    private readonly byte[] digest;
    internal readonly OfdNavigationSource Slots = new();
    private readonly Dictionary<OfdAttachment, Original> originals = new();
    internal int OriginalCount => originals.Count;

    private OfdAttachmentSource(OfdDocumentPackage owner, string documentPath, string descriptorPath, string declaration, XName rootName, byte[] digest)
    { this.owner = owner; DocumentPath = documentPath; DescriptorPath = descriptorPath; Declaration = declaration; RootName = rootName; this.digest = digest; }

    internal bool TryOriginal(OfdAttachment attachment, out Original original) => originals.TryGetValue(attachment, out original!);

    internal bool TryOriginal(int ordinal, out OfdAttachment attachment, out Original original)
    {
        foreach (var pair in originals)
            if (pair.Value.Ordinal == ordinal) { attachment = pair.Key; original = pair.Value; return true; }
        attachment = null!; original = null!; return false;
    }

    internal static OfdAttachmentSource Capture(OfdDocumentPackage package, XElement declaration, string path, XElement root,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ReserveReadModel(2048, token);
        var digestStorage = material.IsControlledRead ? material.ReserveReadStorage(32, token) : null;
        var digest = material.HashBytes(package.PreservedEntries[path], token);
        if (digestStorage is not null) material.AdoptReadStorage(digest, digest.Length, digestStorage, token);
        var result = new OfdAttachmentSource(package, package.DocumentEntryPath!, path, OfdXmlEnvelope.Capture(declaration, material, token), root.Name, digest);
        var ordinal = 0; var index = 0;
        foreach (var node in root.Elements())
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (node.Name == root.Name.Namespace + "Attachment")
            {
                material.ReserveReadModel(2048, token);
                var attachment = package.Attachments[index++];
                var fileLoc = material.ReadXmlValue(node.Element(root.Name.Namespace + "FileLoc"), token) ?? node.Attribute("FileLoc")?.Value ?? string.Empty;
                var payloadPath = attachment.IsExternal ? null : material.ResolvePath(path, fileLoc, token);
                var original = new Original(ordinal, fileLoc, payloadPath, attachment.IsExternal, Values(attachment, material, token));
                result.originals.Add(attachment, original); result.Slots.Slots.Add(ordinal, attachment); attachment.SourceOwner = package;
            }
            ordinal++;
        }
        if (index != package.Attachments.Count) throw new InvalidDataException("Attachment source projection is inconsistent.");
        return result;
    }

    internal static Dictionary<string, string?> Values(OfdAttachment attachment, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(2048, token);
        // Compare modeled values, leaving unchanged source lexicals intact.
        return new Dictionary<string, string?>
        {
            ["ID"] = attachment.Id, ["Name"] = attachment.Name, ["MediaType"] = attachment.MediaType, ["Format"] = attachment.Format,
            ["CreationDate"] = Date(attachment.CreationDate), ["ModDate"] = Date(attachment.ModificationDate),
            ["Size"] = Number(attachment.SizeKilobytes), ["Visible"] = attachment.Visible.HasValue ? attachment.Visible.Value ? "true" : "false" : null, ["Usage"] = attachment.Usage
        };
        string? Date(DateTimeOffset? value) => value.HasValue ? Format(() => value.Value.ToString("O", CultureInfo.InvariantCulture)) : null;
        string? Number(double? value) => value.HasValue ? Format(() => value.Value.ToString("R", CultureInfo.InvariantCulture)) : null;
        string Format(Func<string> format)
        {
            // These fixed-size scalar formats are bounded before allocation.
            var storage = material.IsControlledRead ? material.ReserveReadStorage(1024, token) : null;
            var value = format();
            if (storage is not null) material.AdoptReadString(value, storage, token);
            return value;
        }
    }

    internal static void ValidateOrigins(OfdDocumentPackage package, CancellationToken token)
    {
        foreach (var attachment in package.Attachments)
        {
            token.ThrowIfCancellationRequested();
            if (attachment.SourceOwner is not null && (!ReferenceEquals(attachment.SourceOwner, package) ||
                package.AttachmentSource?.TryOriginal(attachment, out _) != true))
                throw new NotSupportedException("Attachment source cannot be transferred between documents without proven reference closure.");
        }
    }

    internal void Validate(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        ValidateOrigins(package, token);
        if (!ReferenceEquals(owner, package) || package.DocumentEntryPath != DocumentPath)
            throw new NotSupportedException("Cannot migrate imported attachment document ownership or path.");
        if (RootName != XNamespace.Get(package.Options.Namespace) + "Attachments")
            throw new NotSupportedException("Cannot migrate attachment descriptor namespace.");
        material.ChargeWork(32, token);
        if (!package.PreservedEntries.TryGetValue(DescriptorPath, out var bytes) || !material.HashBytes(bytes, token).SequenceEqual(digest))
            throw new InvalidDataException("Attachment descriptor changed behind its reader-owned slots.");
        if (!package.PreservedEntries.TryGetValue(DocumentPath, out var document)) throw new InvalidDataException("Attachment document source is missing.");
        var root = material.ReadXml(DocumentPath, document, token).Root!;
        material.ChargeCopy(128, token);
        var declarations = root.Elements(root.Name.Namespace + "Attachments").Take(2).ToArray();
        if (declarations.Length != 1 || Declaration != OfdXmlEnvelope.Capture(declarations[0], material, token))
            throw new InvalidDataException("Attachment reference changed behind its reader-owned slot.");
    }

    internal static void AssertCopyable(OfdDocumentPackage package, OfdModelMaterialization material, CancellationToken token)
    {
        ValidateOrigins(package, token);
        var source = package.AttachmentSource;
        if (source is null) return;
        source.Validate(package, material, token);
        // Retain the existing typed-copy behavior only for the closed, plain
        // descriptor grammar. Opaque envelopes require a separate migration.
        var root = material.ReadXml(source.DescriptorPath, package.PreservedEntries[source.DescriptorPath], token).Root!;
        material.ChargeCopy(2L * source.Declaration.Length, token);
        var declaration = OfdXmlEnvelope.Parse(source.Declaration);
        if (Attributes(declaration).Any() || declaration.Nodes().Any(n => n is not XText) || Attributes(root).Any() || Mixed(root)) Refuse();
        foreach (var node in root.Elements())
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (node.Name != root.Name.Namespace + "Attachment" || Mixed(node) || Attributes(node).Any(a => a.Name.Namespace != XNamespace.None ||
                a.Name.LocalName is not ("ID" or "Name" or "Format" or "CreationDate" or "ModDate" or "Size" or "Visible" or "Usage" or "External" or "FileLoc"))) Refuse();
            material.ChargeCopy(128, token);
            var locations = node.Elements().Take(2).ToArray();
            if (locations.Length != 1 || locations[0].Name != node.Name.Namespace + "FileLoc" || Attributes(locations[0]).Any() ||
                locations[0].Nodes().Any(n => n is not XText)) Refuse();
        }
        static IEnumerable<XAttribute> Attributes(XElement node) => node.Attributes().Where(a => !a.IsNamespaceDeclaration);
        static bool Mixed(XElement node) => node.Nodes().Any(n => n is not XElement && (n is not XText text || !string.IsNullOrWhiteSpace(text.Value)));
        static void Refuse() => throw new NotSupportedException("Cannot copy opaque attachment descriptors across documents without proven reference closure.");
    }

    internal OfdAttachmentSource Clone(OfdDocumentPackage source, OfdDocumentPackage destination, IReadOnlyDictionary<object, object> mapping,
        OfdModelMaterialization material, CancellationToken token)
    {
        if (!ReferenceEquals(source, owner)) throw new InvalidDataException("Foreign attachment source owner.");
        material.ChargeCopy(2048L + originals.Count * 2048L, token);
        var copy = new OfdAttachmentSource(destination, DocumentPath, DescriptorPath, Declaration, RootName, (byte[])digest.Clone());
        foreach (var pair in originals)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            // Removed slots retain a private tombstone identity; they cannot
            // become active merely because source XML still contains a node.
            // The detached object's current ID still participates in GotoA
            // closure checks when its physical payload is removed.
            var attachment = mapping.TryGetValue(pair.Key, out var mapped) ? (OfdAttachment)mapped : new OfdAttachment { Id = pair.Key.Id };
            attachment.SourceOwner = destination;
            var original = pair.Value;
            copy.originals.Add(attachment, new Original(original.Ordinal, original.FileLoc, original.PayloadPath, original.External,
                new Dictionary<string, string?>(original.Values)));
            copy.Slots.Slots.Add(original.Ordinal, attachment);
        }
        return copy;
    }

    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("attachmentSourceOwner", owner, true);
        sink.Item("attachmentDocument", DocumentPath); sink.Path(DocumentPath);
        sink.Item("attachmentDescriptor", DescriptorPath); sink.Path(DescriptorPath);
        sink.Item("attachmentDeclaration", Declaration); sink.Xml(Declaration);
        sink.Item("attachmentRoot", RootName); sink.Bytes("attachmentSourceDigest", digest);
        sink.Item("attachmentSlots", Slots, true); sink.Item("attachmentSlotXml", Slots.Xml); sink.Xml(Slots.Xml);
        sink.Item("attachmentSlotOwner", Slots.DocumentOwner, true); sink.Item("attachmentSlotOpaque", Slots.HasOpaqueLocalContent); sink.Item("attachmentSlotName", Slots.OriginalName);
        sink.Item("attachmentSlotMap", Slots.Slots, true); sink.Item("count", Slots.Slots.Count);
        foreach (var slot in Slots.Slots) { sink.Item("ordinal", slot.Key); sink.Item("attachmentIdentity", slot.Value, true); }
        sink.Item("attachmentOriginals", originals, true); sink.Item("count", originals.Count);
        foreach (var pair in originals)
        {
            sink.Item("attachmentOriginalIdentity", pair.Key, true); var value = pair.Value;
            sink.Item("attachmentOriginal", value, true); sink.Item("ordinal", value.Ordinal);
            sink.Item("originalLocation", value.FileLoc); sink.Path(value.FileLoc);
            sink.Item("originalPayload", value.PayloadPath); sink.Path(value.PayloadPath); sink.Item("originalExternal", value.External);
            sink.Item("attachmentBaseline", value.Values, true); sink.Item("count", value.Values.Count);
            foreach (var field in value.Values) { sink.Item("field", field.Key); sink.Item("value", field.Value); }
        }
    }
}
