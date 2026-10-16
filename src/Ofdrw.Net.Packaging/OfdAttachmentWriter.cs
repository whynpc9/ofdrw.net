using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal sealed class OfdAttachmentWritePlan
{
    internal XElement? Declaration;
    internal string? DescriptorPath;
    internal readonly Dictionary<string, byte[]> Entries = new(StringComparer.OrdinalIgnoreCase);
}

internal static class OfdAttachmentWriter
{
    internal static OfdAttachmentWritePlan Prepare(OfdDocumentPackage package, string documentPath, XNamespace ns,
        IEnumerable<string> outputPaths, OfdModelMaterialization material, CancellationToken token)
    {
        OfdAttachmentSource.ValidateOrigins(package, token);
        var source = package.AttachmentSource; source?.Validate(package, material, token);
        material.ChargeCopy(2048L + package.Attachments.Count * 512L, token);
        var plan = new OfdAttachmentWritePlan();
        if (package.Attachments.Count == 0 && (source is null || source.OriginalCount > 0)) return plan;
        material.ChargeCopy(128L * package.PreservedEntries.Count, token);
        var occupied = new HashSet<string>(package.PreservedEntries.Keys, StringComparer.OrdinalIgnoreCase);
        foreach (var path in outputPaths) { material.ChargeCopy(128, token); occupied.Add(path); }
        var directory = OfdPackagePath.GetDirectory(documentPath);
        plan.DescriptorPath = source?.DescriptorPath ?? Fresh(InDirectory(directory, "Attachs/Attachments.xml"));
        occupied.Add(plan.DescriptorPath);
        if (source is not null) material.ChargeCopy(2L * source.Declaration.Length, token);
        plan.Declaration = source is null ? new XElement(ns + "Attachments", material.RelativePath(documentPath, plan.DescriptorPath, token)) : OfdXmlEnvelope.Parse(source.Declaration);
        var xml = source is null ? new XDocument(new XDeclaration("1.0", "UTF-8", null),
            new XElement(ns + "Attachments", new XAttribute(XNamespace.Xmlns + "ofd", ns.NamespaceName))) :
            material.ReadXml(source.DescriptorPath, package.PreservedEntries[source.DescriptorPath], token);
        var root = xml.Root!;
        material.ChargeCopy(32L * root.Elements().Count(), token);
        var originalNodes = root.Elements().ToArray();
        var index = 0;
        OfdXmlEnvelope.Children(root, source?.Slots, package.Attachments, attachment =>
        {
            token.ThrowIfCancellationRequested(); var ordinal = ++index;
            if (source?.TryOriginal(attachment, out var original) == true)
            {
                if (attachment.IsExternal != original.External)
                    throw new NotSupportedException("Cannot migrate imported attachment payload ownership without proven reference closure.");
                // Original materialization already admitted the complete DOM;
                // this second node copy also pays for its serialized envelope.
                var captured = OfdXmlEnvelope.Capture(originalNodes[original.Ordinal], material, token);
                material.ChargeCopy(2L * captured.Length, token);
                var node = OfdXmlEnvelope.Parse(captured);
                var values = OfdAttachmentSource.Values(attachment, material, token);
                foreach (var pair in values)
                {
                    token.ThrowIfCancellationRequested();
                    if (pair.Value != original.Values[pair.Key]) node.SetAttributeValue(pair.Key, pair.Value);
                }
                if (attachment.IsExternal)
                {
                    var location = attachment.ExternalPath ?? string.Empty;
                    if (location != original.FileLoc) SetLocation(node, location, token);
                }
                else Payload(original.PayloadPath!, attachment.Data);
                return node;
            }
            material.ChargeCopy(512L + 8L * attachment.Name.Length, token);
            var fileName = $"Attach_{ordinal}_{Sanitize(attachment.Name)}";
            var path = attachment.IsExternal ? null : Fresh(material.ResolvePath(plan.DescriptorPath!, fileName, token));
            if (path is not null) { occupied.Add(path); Payload(path, attachment.Data); }
            return new XElement(ns + "Attachment",
                new XAttribute("ID", attachment.Id ?? $"Attachment_{ordinal}"), new XAttribute("Name", attachment.Name),
                new XAttribute("Format", attachment.Format ?? attachment.MediaType),
                attachment.CreationDate.HasValue ? new XAttribute("CreationDate", attachment.CreationDate.Value.ToString("O", CultureInfo.InvariantCulture)) : null,
                attachment.ModificationDate.HasValue ? new XAttribute("ModDate", attachment.ModificationDate.Value.ToString("O", CultureInfo.InvariantCulture)) : null,
                attachment.SizeKilobytes.HasValue ? new XAttribute("Size", attachment.SizeKilobytes.Value.ToString("0.###", CultureInfo.InvariantCulture)) : null,
                attachment.Visible.HasValue ? new XAttribute("Visible", attachment.Visible.Value) : null,
                !string.IsNullOrWhiteSpace(attachment.Usage) ? new XAttribute("Usage", attachment.Usage) : null,
                new XElement(ns + "FileLoc", attachment.IsExternal ? attachment.ExternalPath ?? string.Empty : material.RelativePath(plan.DescriptorPath!, path!, token)));
        }, token);
        plan.Entries.Add(plan.DescriptorPath, Encoding.UTF8.GetBytes(xml.ToString(SaveOptions.DisableFormatting)));
        return plan;

        void Payload(string path, byte[] data)
        {
            if (string.Equals(path, plan.DescriptorPath, StringComparison.OrdinalIgnoreCase)) throw new InvalidDataException("Attachment payload aliases its descriptor.");
            if (plan.Entries.TryGetValue(path, out var previous))
            {
                material.ChargeWork(previous.LongLength + data.LongLength, token);
                if (!Equal(previous, data, token)) throw new NotSupportedException("Shared attachment payload has conflicting edits.");
            }
            else { material.ChargeCopy(128, token); plan.Entries.Add(path, data); }
        }
        string Fresh(string preferred)
        {
            var path = preferred; var ordinal = 0;
            while (occupied.Contains(path))
            {
                token.ThrowIfCancellationRequested(); material.ChargeCopy(512L + 8L * preferred.Length, token);
                path = preferred + "." + (++ordinal).ToString(CultureInfo.InvariantCulture);
            }
            return path;
        }
    }

    private static bool Equal(byte[] first, byte[] second, CancellationToken token)
    {
        if (first.Length != second.Length) return false;
        for (var index = 0; index < first.Length; index++) { if ((index & 1023) == 0) token.ThrowIfCancellationRequested(); if (first[index] != second[index]) return false; }
        return true;
    }

    private static void SetLocation(XElement node, string value, CancellationToken token)
    {
        var locations = node.Elements(node.Name.Namespace + "FileLoc").ToArray();
        if (locations.Length > 1) throw new InvalidDataException("Duplicate attachment FileLoc slots.");
        if (locations.Length == 0) { node.SetAttributeValue("FileLoc", value); return; }
        var location = locations[0];
        if (location.HasElements || location.Nodes().Any(child => child is XProcessingInstruction))
            throw new NotSupportedException("Cannot overwrite an opaque attachment FileLoc scalar.");
        if (node.Attribute("FileLoc") is not null) throw new NotSupportedException("Cannot migrate ambiguous legacy attachment FileLoc aliases.");
        token.ThrowIfCancellationRequested(); location.Nodes().OfType<XText>().Remove(); location.AddFirst(new XText(value));
    }

    private static string InDirectory(string directory, string path) => directory.Length == 0 ? path : directory + "/" + path;
    private static string Sanitize(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "Attachment.bin";
        var invalid = Path.GetInvalidFileNameChars();
        return string.Concat(value.Select(character => invalid.Contains(character) || character is '/' or '\\' ? '_' : character));
    }
}
