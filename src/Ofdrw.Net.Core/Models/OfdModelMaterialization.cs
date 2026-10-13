using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Models;

// Three independent operation-wide limits, not a combined managed-heap cap.
// Callers must pass one instance through nested copy/edit operations.
internal sealed class OfdModelMaterialization
{
    private readonly OfdModelLimits limits;
    private readonly HashSet<object> stored = new(Identity.Instance);
    private readonly HashSet<(object Owner, string? MediaType, string? FileName)> outputImages = new();
    internal OfdTraversalBudget Structure { get; }
    internal long StorageBytes { get; private set; }
    internal long CloneCopyBytes { get; private set; }
    internal long OutputUpperBoundBytes { get; private set; }
    private readonly Dictionary<byte[], XmlWork> xmlEntries = new();
    private sealed class XmlWork
    {
        internal long Nodes; internal long References; internal int Depth; internal bool HasMaxHint;
        internal readonly HashSet<long> ReservedIds = new();
    }
    private readonly HashSet<byte[]> opaqueBytes = new();
    internal readonly HashSet<string> OpaqueEntries = new(StringComparer.OrdinalIgnoreCase);

    internal OfdModelMaterialization(OfdModelLimits limits)
    {
        this.limits = limits.Snapshot(); Structure = new OfdTraversalBudget(this.limits);
    }

    internal void InspectStorage(OfdDocumentPackage package, CancellationToken token)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        token.ThrowIfCancellationRequested(); Structure.Node(1);
        if (package.Options is null || package.Options.Metadata is null) throw new InvalidDataException("Document options and metadata are required.");
        Strings(package.DocumentEntryPath, package.PublicResourceLocation, package.DocumentResourceLocation,
            package.Options.DocumentId, package.Options.Namespace, package.Options.DocType);
        Paths(package.DocumentEntryPath, package.PublicResourceLocation, package.DocumentResourceLocation, package.Options.DocumentId);
        var metadata = package.Options.Metadata;
        Strings(metadata.Title, metadata.Author, metadata.Subject, metadata.Keywords, metadata.Creator);
        foreach (var entry in package.PreservedEntries) { Visit(); StoreString(entry.Key); StoreArray(entry.Value); }
        foreach (var font in package.Fonts)
        {
            Visit(); if (font is null) throw new InvalidDataException("Null font resource.");
            Strings(font.Id, font.FontName, font.FamilyName, font.Charset, font.FileName); Paths(font.FileName); StoreArray(font.Data);
        }
        foreach (var attachment in package.Attachments)
        {
            Visit(); if (attachment is null) throw new InvalidDataException("Null attachment.");
            Strings(attachment.Id, attachment.Name, attachment.MediaType, attachment.Format, attachment.Usage, attachment.ExternalPath);
            Paths(attachment.ExternalPath); StoreArray(attachment.Data);
        }
        foreach (var tag in package.CustomTags) { Visit(); Strings(tag.Key, tag.Value); }
        Preserved(package.PreservedCommonDataElements); Preserved(package.PreservedDocumentElements); Preserved(package.PreservedDocBodyElements);
        foreach (var page in package.Pages)
        {
            Visit(); if (page is null) throw new InvalidDataException("Null page."); Strings(page.Id, page.SourceEntryPath); Paths(page.SourceEntryPath);
            Graphics(page.Elements); Graphics(page.AnnotationAppearances); Preserved(page.PreservedPageElements);
            foreach (var template in page.Templates)
            {
                Visit(); if (template is null) throw new InvalidDataException("Null template.");
                Strings(template.TemplateId, template.ZOrder, template.BaseLocation); Paths(template.BaseLocation); Graphics(template.Elements);
            }
        }
        Navigation(package.OutlineOwnership); Navigation(package.BookmarkOwnership); Navigation(package.ActionOwnership);
        token.ThrowIfCancellationRequested();

        void Navigation<T>(OfdNavigationOwnership<T> owner) where T : class
        {
            if (owner.Value is not null) OfdNavigationSnapshot.Capture(owner.Value, this, token).InspectMaterial(this, false, token, guardXml: true);
            StoreString(owner.OriginalXml); owner.BaselineSnapshot?.InspectMaterial(this, false, token);
        }
        void Visit() { token.ThrowIfCancellationRequested(); Structure.Node(1); Structure.Reference(); }
        void Paths(params string?[] values) { foreach (var value in values) if (value is not null) CheckPathWork("OFD.xml", value, token); }
        void Strings(params string?[] values) { foreach (var value in values) { token.ThrowIfCancellationRequested(); StoreString(value); } }
        void Preserved(List<string> values) { foreach (var value in values) { Visit(); StoreString(value); Xml(value); } }
        void Xml(string? value) { if (!string.IsNullOrWhiteSpace(value)) GuardFragmentXml(value!, token); }
        void Graphics(List<OfdElement> elements)
        {
            foreach (var element in elements)
            {
                Visit(); if (element is null) throw new InvalidDataException("Null graphic.");
                Strings(element.ObjectId, element.LayerId, element.LayerType, element.ClippingXml); Xml(element.ClippingXml);
                switch (element)
                {
                    case OfdTextElement text:
                        Strings(text.Text, text.FontName, text.FontResourceId, text.SourceXml); Xml(text.SourceXml); StoreMatrix(text.Transform);
                        foreach (var run in text.Runs) { Visit(); if (run is null) throw new InvalidDataException("Null text run."); Strings(run.Text, run.DeltaX, run.DeltaY); }
                        break;
                    case OfdImageElement image:
                        Strings(image.ResourceId, image.FileName, image.MediaType, image.SourceXml); Xml(image.SourceXml); StoreArray(image.Data); StoreMatrix(image.Transform); break;
                    case OfdPathElement path:
                        Strings(path.SourceXml, path.AbbreviatedData); Xml(path.SourceXml); StoreMatrix(path.Transform); break;
                    case OfdRawElement raw: Strings(raw.Xml, raw.LocalName); Xml(raw.Xml); break;
                    default: throw new NotSupportedException("Unknown graphic materialization contract.");
                }
            }
        }
    }

    internal string ResolvePath(string basis, string reference, CancellationToken token)
    {
        CheckPathWork(basis, reference, token); return OfdPackagePath.Resolve(basis, reference);
    }
    internal string RelativePath(string basis, string target, CancellationToken token)
    {
        CheckPathWork(basis, target, token); return OfdPackagePath.RelativeTo(basis, target);
    }
    private void CheckPathWork(string basis, string reference, CancellationToken token)
    {
        token.ThrowIfCancellationRequested(); Structure.SourcePathCharacters((long)basis.Length + reference.Length + 2);
        Separators(basis); Separators(reference);
        void Separators(string value)
        {
            var next = 0;
            for (var index = 0; index < value.Length; index++)
            {
                OfdModelTraversal.CheckCancellationCheckpoint(index, ref next, token);
                if (value[index] is '/' or '\\') Structure.Reference();
            }
        }
    }

    internal void GuardFragmentXml(string value, CancellationToken token) => OfdModelTraversal.GuardXml(value, Structure, limits, token);

    internal void StoreString(string? value)
    {
        if (value is null) return;
        Store(value, 2L * value.Length);
    }
    internal void StoreArray(byte[] value)
    {
        if (value is null) throw new InvalidDataException("Null payload array.");
        Store(value, value.LongLength);
    }
    internal void StoreMatrix(double[]? value)
    {
        if (value is null) return;
        if (value.Length != 6) throw new ArgumentException("A model CTM must contain six finite values.");
        foreach (var number in value) if (double.IsNaN(number) || double.IsInfinity(number)) throw new ArgumentException("A model CTM must contain six finite values.");
        Store(value, value.LongLength * sizeof(double));
    }
    private void Store(object value, long bytes)
    {
        Structure.Reference();
        if (stored.Contains(value)) return;
        StorageBytes = Add(StorageBytes, bytes, "input storage");
        stored.Add(value);
    }

    internal void ChargeCopy(long bytes, CancellationToken token)
    {
        token.ThrowIfCancellationRequested(); Structure.Reference();
        CloneCopyBytes = Add(CloneCopyBytes, bytes, "clone-copy");
    }
    internal void ChargeOutput(long bytes)
        => OutputUpperBoundBytes = Add(OutputUpperBoundBytes, bytes, "logical output");

    internal void GuardPreservedXml(OfdDocumentPackage package, CancellationToken token)
    {
        foreach (var entry in package.PreservedEntries) GuardEntry(entry.Key, entry.Value, token);
    }

    internal void GuardEntry(string path, byte[] bytes, CancellationToken token)
    {
            token.ThrowIfCancellationRequested(); Structure.Reference();
            if (xmlEntries.ContainsKey(bytes)) { OpaqueEntries.Remove(path); return; }
            if (opaqueBytes.Contains(bytes) || !OfdXmlContentProbe.LooksLikeXml(bytes)) { opaqueBytes.Add(bytes); OpaqueEntries.Add(path); return; }
            Structure.XmlBytes(bytes.LongLength);
            // UTF-16 decoding is at most two bytes per source byte for the
            // supported XML encodings. Reserve each decoder materialization
            // before XmlReader can form large attribute/text strings.
            ChargeCopy(2L * bytes.LongLength, token);
            var work = new XmlWork();
            try
            {
                using var input = new MemoryStream(bytes, false);
                using var reader = XmlReader.Create(input, Settings());
                while (reader.Read())
                {
                    token.ThrowIfCancellationRequested();
                    if (reader.NodeType == XmlNodeType.DocumentType) throw new InvalidDataException("DTD is prohibited in preserved XML.");
                    Structure.Node(reader.Depth + 1); work.Nodes++; work.Depth = Math.Max(work.Depth, reader.Depth + 1);
                    if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "MaxUnitID") work.HasMaxHint = true;
                    for (var a = 0; a < reader.AttributeCount; a++)
                    {
                        token.ThrowIfCancellationRequested(); reader.MoveToAttribute(a);
                        Structure.Reference(); work.References++;
                        if (reader.LocalName is "ID" or "MaxUnitID")
                        {
                            var id = OfdSourceIdentity.ParseId(reader.Value);
                            if (id != 0 && !work.ReservedIds.Contains(id)) { Structure.Reference(); work.ReservedIds.Add(id); }
                            if (reader.LocalName == "MaxUnitID") work.HasMaxHint = true;
                        }
                        // Existing numeric projection uses Split. Bound token work
                        // before its arrays can be created, without interpreting it.
                        if (reader.LocalName is "CTM" or "Boundary" or "Value") Tokens(reader.Value);
                    }
                    reader.MoveToElement();
                    if (reader.NodeType == XmlNodeType.Text) Tokens(reader.Value);
                }
                xmlEntries[bytes] = work;
            }
            catch (XmlException)
            {
                // Only malformed XML is an opaque extension. Cancellation,
                // DTD and quota failures are intentionally not caught.
                opaqueBytes.Add(bytes); OpaqueEntries.Add(path);
            }

            void Tokens(string value)
            {
                var inToken = false;
                for (var i = 0; i < value.Length; i++)
                {
                    if ((i & 1023) == 0) token.ThrowIfCancellationRequested();
                    var next = !char.IsWhiteSpace(value[i]);
                    if (next && !inToken) { Structure.Reference(); work.References++; }
                    inToken = next;
                }
            }
        token.ThrowIfCancellationRequested();
    }

    internal IEnumerable<long> Reservations(byte[] bytes, out bool hasMaxHint)
    {
        if (!xmlEntries.TryGetValue(bytes, out var work)) { hasMaxHint = false; return Array.Empty<long>(); }
        hasMaxHint = work.HasMaxHint; return work.ReservedIds;
    }

    internal XDocument ReadXml(string path, byte[] bytes, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (!xmlEntries.ContainsKey(bytes)) GuardEntry(path, bytes, token);
        if (!xmlEntries.TryGetValue(bytes, out var work)) throw new XmlException("Active entry is not bounded parseable XML: " + path);
        // Every DOM materialization is work, including aliases of the same entry.
        ChargeCopy(2L * bytes.LongLength, token);
        Structure.NodeWork(work.Nodes, work.Depth);
        for (long index = 0; index < work.References; index++) { token.ThrowIfCancellationRequested(); Structure.Reference(); }
        using var input = new MemoryStream(bytes, false);
        using var reader = XmlReader.Create(input, Settings());
        var result = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        token.ThrowIfCancellationRequested(); return result;
    }

    private XmlReaderSettings Settings() => new()
    {
        // Report DocumentType without expanding an entity: no DOM is admitted
        // until the streaming pass has rejected the declaration.
        DtdProcessing = DtdProcessing.Parse, XmlResolver = null,
        MaxCharactersInDocument = limits.MaxSourceXmlBytes, MaxCharactersFromEntities = 1024
    };

    internal void EstimateWriter(OfdDocumentPackage package, CancellationToken token)
    {
        var ns = package.Options.Namespace;
        foreach (var entry in package.PreservedEntries)
        {
            Check(); EmitString(entry.Key);
            if (xmlEntries.TryGetValue(entry.Value, out var work))
                ChargeOutput(6L * entry.Value.LongLength + work.Nodes * (4L * (limits.MaxDepth + 16L) + 4 + 6L * ns.Length));
            else ChargeOutput(entry.Value.LongLength);
        }
        Markup(12, 16, 16); EmitString(ns, 12);
        var metadata = package.Options.Metadata;
        EmitString(metadata.Title); EmitString(metadata.Author); EmitString(metadata.Subject); EmitString(metadata.Keywords); EmitString(metadata.Creator);
        EmitString(package.Options.DocType); EmitString(package.Options.DocumentId); EmitString(package.DocumentEntryPath);
        EmitString(package.PublicResourceLocation); EmitString(package.DocumentResourceLocation);
        foreach (var value in package.PreservedCommonDataElements) EmitXml(value);
        foreach (var value in package.PreservedDocumentElements) EmitXml(value);
        foreach (var value in package.PreservedDocBodyElements) EmitXml(value);
        foreach (var page in package.Pages)
        {
            Check(); Markup(8, 32, 32); EmitString(ns, 8);
            EmitString(package.DocumentEntryPath ?? package.Options.DocumentId, 4);
            EmitString(page.SourceEntryPath, 4); EmitString(page.Id);
            foreach (var xml in page.PreservedPageElements) EmitXml(xml);
            foreach (var element in page.Elements)
            {
                Check(); Markup(8, 32, 32); EmitString(ns, 8);
                EmitString(element.LayerId); EmitString(element.LayerType); EmitString(element.ObjectId); EmitXml(element.ClippingXml);
                switch (element)
                {
                    case OfdTextElement text:
                        EmitXml(text.SourceXml); EmitString(text.Text); EmitString(text.FontName, 2); EmitString(text.FontResourceId);
                        Markup(3, 4, 8); EmitString(ns, 3); // possible fallback Font/FontFile declaration
                        if (text.Runs.Count == 0) { Markup(1, 2, 4); EmitString(ns); }
                        foreach (var run in text.Runs) { Check(); Markup(1, 2, 4); EmitString(ns); EmitString(run.Text); EmitString(run.DeltaX); EmitString(run.DeltaY); }
                        break;
                    case OfdImageElement image:
                        EmitXml(image.SourceXml); EmitString(image.FileName); EmitString(image.MediaType); EmitString(image.ResourceId);
                        // Different imported resource owners remain distinct output
                        // entries even when their current byte arrays are aliased.
                        var imageKey = package.SourceIdentity?.TryImage(image, out var imageOwner) == true
                            ? ((object)imageOwner, (string?)null, (string?)null)
                            : ((object)image.Data, image.MediaType, image.FileName);
                        if (outputImages.Add(imageKey)) ChargeOutput(image.Data.LongLength);
                        break;
                    case OfdPathElement path: EmitXml(path.SourceXml); EmitString(path.AbbreviatedData); break;
                    case OfdRawElement raw: EmitXml(raw.Xml); break;
                }
            }
        }
        foreach (var font in package.Fonts)
        {
            Check(); Markup(3, 4, 8); EmitString(ns, 3); ChargeOutput(font.Data.LongLength);
            EmitString(font.Id); EmitString(font.FontName); EmitString(font.FamilyName); EmitString(font.Charset); EmitString(font.FileName);
        }
        foreach (var attachment in package.Attachments)
        {
            Check(); Markup(2, 8, 16); EmitString(ns, 2);
            EmitString(attachment.Name, 3); EmitString(attachment.Id); EmitString(attachment.MediaType);
            EmitString(attachment.Format); EmitString(attachment.Usage); EmitString(attachment.ExternalPath);
            if (!attachment.IsExternal) ChargeOutput(attachment.Data.LongLength);
        }
        foreach (var tag in package.CustomTags) { Check(); Markup(1, 0, 2); EmitString(ns); EmitString(tag.Key); EmitString(tag.Value); }
        Navigation(package.OutlineOwnership); Navigation(package.BookmarkOwnership); Navigation(package.ActionOwnership);
        Check();

        void Navigation<T>(OfdNavigationOwnership<T> owner) where T : class
        {
            if (owner.Value is not null) OfdNavigationSnapshot.Capture(owner.Value, this, token).InspectMaterial(this, true, token);
            // Replacement still retains the original container's opaque slots.
            EmitXml(owner.OriginalXml);
            if (owner.OriginalXml is not null) ChargeCopy(2L * owner.OriginalXml.Length, token);
        }
        void Check() => token.ThrowIfCancellationRequested();
        void EmitString(string? value, int positions = 1)
        {
            Check(); if (value is not null) ChargeOutput(6L * value.Length * positions);
        }
        void EmitXml(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            EmitString(value);
            // Each lexical XML token accounts for extra indentation and a
            // potential namespace binding at its output placement.
            using var reader = XmlReader.Create(new StringReader(value), new XmlReaderSettings { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = limits.MaxSourceXmlBytes });
            while (reader.Read()) { Check(); Structure.Depth(reader.Depth + 1); Markup(1, 0, 0); EmitString(ns); }
        }
        void Markup(int elements, int numbers, int attributes)
        {
            Structure.NodeWork(elements, 1);
            // Owned emitters have names <=64 UTF8 bytes and <=32 numeric
            // scalars per graphic (including faux-italic's duplicated CTM).
            // Fixed 0.### double formatting is <=326 ASCII chars. Source XML
            // names/values are separately covered by six times lexical length.
            var literal = elements * (2L * (64 + 3) + 4L * (limits.MaxDepth + 16L) + 4);
            ChargeOutput(literal + attributes * (64L + 4) + numbers * 326L);
        }
    }

    internal long ValidateActualEntries(IDictionary<string, byte[]> entries, CancellationToken token)
    {
        long actual = 0;
        foreach (var entry in entries) { token.ThrowIfCancellationRequested(); actual = Add(actual, entry.Value.LongLength, "actual logical output"); }
        if (actual > OutputUpperBoundBytes) throw new InvalidDataException("Writer output exceeded its preflight upper bound.");
        return actual;
    }

    private long Add(long used, long next, string kind)
    {
        if (next < 0 || next > limits.MaxMaterializedBytes - used) throw OfdTraversalBudget.Limit($"OFD {kind} exceeds its materialization budget.");
        return used + next;
    }
    private sealed class Identity : IEqualityComparer<object>
    {
        internal static readonly Identity Instance = new();
        public new bool Equals(object? a, object? b) => ReferenceEquals(a, b);
        public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
    }
}
