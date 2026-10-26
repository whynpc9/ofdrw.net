using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Security.Cryptography;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Models;

// Four independent operation-wide limits, not a combined managed-heap cap.
// Callers must pass one instance through nested copy/edit operations.
internal sealed partial class OfdModelMaterialization
{
    private readonly OfdModelLimits limits;
    private readonly HashSet<object> stored = new(Identity.Instance);
    internal OfdTraversalBudget Structure { get; }
    internal long StorageBytes { get; private set; }
    internal long CloneCopyBytes { get; private set; }
    internal long OutputUpperBoundBytes { get; private set; }
    // Cumulative hashing/comparison work; not allocation, CPU time or RSS.
    internal long ValidationWorkBytes { get; private set; }
    private readonly Dictionary<byte[], XmlWork> xmlEntries = new();
    private readonly Dictionary<byte[], string?> pageBlockImageAdmissions = new();
    internal HashSet<(OfdCompositeResource Resource, bool InPageBlock)> PageBlockAdmittedComposites { get; } = new();
    private sealed class XmlWork
    {
        internal byte[] Digest = Array.Empty<byte>();
        internal long Nodes; internal long References; internal int Depth; internal bool HasMaxHint;
        internal readonly HashSet<long> ReservedIds = new();
    }
    internal readonly HashSet<string> OpaqueEntries = new(StringComparer.OrdinalIgnoreCase);

    internal OfdModelMaterialization(OfdModelLimits limits)
    {
        this.limits = limits.Snapshot(); Structure = new OfdTraversalBudget(this.limits);
    }

    internal void InspectStorage(OfdDocumentPackage package, CancellationToken token)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        OfdSerializedInventory.Walk(package, new StorageSink(this, token), this, token);
    }

    private sealed class StorageSink(OfdModelMaterialization material, CancellationToken token) : OfdSerializedInventory.ISink
    {
        public void Item(string role, object? value, bool identity = false)
        {
            token.ThrowIfCancellationRequested();
            if (value is string text) material.StoreString(text);
            else if (value is XName name) { material.StoreString(name.NamespaceName); material.StoreString(name.LocalName); }
        }
        public void Bytes(string role, byte[] value, object? owner = null) { token.ThrowIfCancellationRequested(); material.StoreArray(value); }
        public void Matrix(double[]? value) { token.ThrowIfCancellationRequested(); material.StoreMatrix(value); }
        public void Xml(string? value) { if (!string.IsNullOrWhiteSpace(value)) material.GuardFragmentXml(value!, token); }
        public void Path(string? value) { if (value is not null) material.CheckPathWork("OFD.xml", value, token); }
    }

    internal string ResolvePath(string basis, string reference, CancellationToken token)
    {
        CheckPathWork(basis, reference, token); ChargeCopy(8L * (basis.Length + (long)reference.Length + 2), token);
        if (!IsControlledRead) return OfdPackagePath.Resolve(basis, reference);
        var storage = ReserveReadStorage(checked(2L * (basis.Length + (long)reference.Length + 2)), token);
        var result = OfdPackagePath.Resolve(basis, reference); AdoptReadString(result, storage, token); return result;
    }
    internal string RelativePath(string basis, string target, CancellationToken token)
    {
        CheckPathWork(basis, target, token); ChargeCopy(12L * (basis.Length + (long)target.Length + 2), token);
        if (!IsControlledRead) return OfdPackagePath.RelativeTo(basis, target);
        var storage = ReserveReadStorage(checked(4L * (basis.Length + (long)target.Length + 2)), token);
        var result = OfdPackagePath.RelativeTo(basis, target); AdoptReadString(result, storage, token); return result;
    }
    internal void ReservePathComposition(string basis, string reference, CancellationToken token)
    { CheckPathWork(basis, reference, token); ChargeCopy(2L * (basis.Length + (long)reference.Length + 2), token); }
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

    internal void GuardFragmentXml(string value, CancellationToken token)
    {
        PrepareReadXmlParser(checked(2L * value.Length), token);
        OfdModelTraversal.GuardXml(value, Structure, limits, token);
    }

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
    internal long OwnedPayloadAllocations { get; private set; }
    internal byte[] AllocateOwnedPayload(int length, CancellationToken token)
    {
        // Pay array/header/copy and retained-storage admission before allocation.
        // Adoption prevents the final whole-model inventory charging it twice.
        ChargeCopy(checked(128L + length), token);
        StorageBytes = Add(StorageBytes, length, "owned payload retained-storage reservation");
        OwnedPayloadAllocations++;
        var value = new byte[length]; stored.Add(value); return value;
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
        token.ThrowIfCancellationRequested(); controlledContext?.AssertActive(); Structure.Reference();
        CloneCopyBytes = Add(CloneCopyBytes, bytes, "clone-copy");
    }
    internal void ChargeOutput(long bytes)
        => OutputUpperBoundBytes = Add(OutputUpperBoundBytes, bytes, "logical output");

    internal void ChargeWork(long bytes, CancellationToken token)
    {
        token.ThrowIfCancellationRequested(); controlledContext?.AssertActive(); Structure.Reference();
        ValidationWorkBytes = Add(ValidationWorkBytes, bytes, "validation work");
    }

    // Only PageBlock references use this shared three-export profile. Charge
    // compressed-byte and decoded-pixel work once per archive payload, including
    // repeated placements, before doing the costly structural walk.
    internal string? AdmitPageBlockImage(byte[] payload, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (pageBlockImageAdmissions.TryGetValue(payload, out var cached)) return cached;
        var mime = EncodedImageAdmission.IdentifyFormat(payload);
        if (mime is not ("image/png" or "image/jpeg" or "image/jbig2"))
        { pageBlockImageAdmissions.Add(payload, null); return null; }
        ChargeWork(checked(payload.LongLength * 16), token);
        try
        {
            if (mime == "image/jbig2")
            {
                var size = RestrictedJbig2Admission.Validate(payload, token);
                ChargeWork(checked((long)size.Width * size.Height * 16), token);
            }
            else EncodedImageAdmission.Validate(payload, mime, token,
                pixels => ChargeWork(checked(pixels * 16), token));
            pageBlockImageAdmissions.Add(payload, mime); return mime;
        }
        catch (Exception error) when (!OfdTraversalBudget.IsLimit(error) &&
            error is InvalidDataException or IOException or NotSupportedException or OverflowException or IndexOutOfRangeException)
        { pageBlockImageAdmissions.Add(payload, null); return null; }
    }

    internal byte[] HashBytes(byte[] bytes, CancellationToken token)
    {
        if (bytes is null) throw new InvalidDataException("Null payload array.");
        ChargeCopy(32, token);
        using var hash = SHA256.Create();
        for (var offset = 0; offset < bytes.Length;)
        {
            var count = Math.Min(64 * 1024, bytes.Length - offset); ChargeWork(count, token);
            hash.TransformBlock(bytes, offset, count, null, 0); offset += count;
        }
        token.ThrowIfCancellationRequested(); hash.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return hash.Hash!;
    }

    internal void GuardPreservedXml(OfdDocumentPackage package, CancellationToken token)
    {
        foreach (var entry in package.PreservedEntries) GuardEntry(entry.Key, entry.Value, token);
    }

    internal void GuardEntry(string path, byte[] bytes, CancellationToken token)
    {
            token.ThrowIfCancellationRequested(); Structure.Reference();
            if (xmlEntries.TryGetValue(bytes, out var cached)) { AssertXmlBytes(bytes, cached, token); OpaqueEntries.Remove(path); return; }
            if (!OfdXmlContentProbe.LooksLikeXml(bytes, token, count => ChargeWork(count, token))) { OpaqueEntries.Add(path); return; }
            Structure.XmlBytes(bytes.LongLength);
            // UTF-16 decoding is at most two bytes per source byte for the
            // supported XML encodings. Reserve each decoder materialization
            // before XmlReader can form large attribute/text strings.
            ChargeCopy(2L * bytes.LongLength, token);
            PrepareReadXmlParser(bytes.LongLength, token);
            var work = new XmlWork { Digest = HashBytes(bytes, token) };
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
                            if (id != 0 && !work.ReservedIds.Contains(id)) { ReserveReadModel(128, token); Structure.Reference(); work.ReservedIds.Add(id); }
                            if (reader.LocalName == "MaxUnitID") work.HasMaxHint = true;
                        }
                        // Existing numeric projection uses Split. Bound token work
                        // before its arrays can be created, without interpreting it.
                        if (reader.LocalName is "CTM" or "Boundary" or "Value") Tokens(reader.Value);
                    }
                    reader.MoveToElement();
                    if (reader.NodeType == XmlNodeType.Text) Tokens(reader.Value);
                }
                xmlEntries[bytes] = work; OpaqueEntries.Remove(path);
            }
            catch (XmlException)
            {
                // Controlled parsing prohibits DTD before its internal grammar
                // compiler can allocate. Reject all XmlException here rather than
                // identifying DTD via a localized runtime exception message.
                // Public loading retains its original malformed-opaque policy.
                if (HasControlledXmlPolicy) throw new InvalidDataException("Controlled operation contains malformed or prohibited XML.");
                OpaqueEntries.Add(path);
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

    private void AssertXmlBytes(byte[] bytes, XmlWork work, CancellationToken token)
    {
        var digest = HashBytes(bytes, token); ChargeWork(32, token);
        for (var index = 0; index < digest.Length; index++)
            if (digest[index] != work.Digest[index]) throw new InvalidDataException("Cached XML bytes changed during the materialization operation.");
    }

    internal IEnumerable<long> Reservations(byte[] bytes, out bool hasMaxHint, CancellationToken token = default)
    {
        if (!xmlEntries.TryGetValue(bytes, out var work)) { hasMaxHint = false; return Array.Empty<long>(); }
        AssertXmlBytes(bytes, work, token);
        hasMaxHint = work.HasMaxHint; return work.ReservedIds;
    }

    internal XDocument ReadXml(string path, byte[] bytes, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        GuardEntry(path, bytes, token);
        if (!xmlEntries.TryGetValue(bytes, out var work)) throw new XmlException("Active entry is not bounded parseable XML: " + path);
        // Every DOM materialization is work, including aliases of the same entry.
        ChargeCopy(2L * bytes.LongLength, token);
        Structure.NodeWork(work.Nodes, work.Depth);
        for (long index = 0; index < work.References; index++) { token.ThrowIfCancellationRequested(); Structure.Reference(); }
        PrepareReadXmlParser(bytes.LongLength, token);
        var readStorage = PrepareReadDom(bytes.LongLength, work.Nodes, work.References, token);
        if (IsControlledRead) ChargeWork(bytes.LongLength, token);
        using var input = new MemoryStream(bytes, false);
        using var reader = XmlReader.Create(input, Settings());
        var result = XDocument.Load(reader, LoadOptions.PreserveWhitespace);
        AdoptReadDom(result, readStorage, token);
        token.ThrowIfCancellationRequested(); return result;
    }

    private XmlReaderSettings Settings() => new()
    {
        // Public path retains its original DocumentType admission. Controlled
        // parsing rejects the declaration before compiling internal DTD models.
        DtdProcessing = HasControlledXmlPolicy ? DtdProcessing.Prohibit : DtdProcessing.Parse, XmlResolver = null,
        MaxCharactersInDocument = limits.MaxSourceXmlBytes, MaxCharactersFromEntities = 1024
    };

    internal void EstimateWriter(OfdDocumentPackage package, CancellationToken token)
    {
        var sink = new WriterEstimateSink(this, package, token);
        OfdSerializedInventory.Walk(package, sink, this, token);
    }

    // The inventory owns the field table. Only output structure/placement
    // constants and resource-owner deduplication live in this emitter policy.
    private sealed class WriterEstimateSink(OfdModelMaterialization material, OfdDocumentPackage package, CancellationToken token) : OfdSerializedInventory.ISink
    {
        private readonly HashSet<(object Owner, string? MediaType, string? FileName)> images = new();
        private string Namespace => package.Options.Namespace;
        public void Item(string role, object? value, bool identity = false)
        {
            token.ThrowIfCancellationRequested(); if (value is string text) EmitString(text);
            switch (role)
            {
                case "package": Markup(12, 16, 16); EmitString(Namespace, 12); break;
                case "page":
                    Markup(8, 32, 32); EmitString(Namespace, 8);
                    EmitString(package.DocumentEntryPath ?? package.Options.DocumentId, 4);
                    EmitString(((OfdPage)value!).SourceEntryPath, 4); break;
                case "graphic":
                    Markup(8, 32, 32); EmitString(Namespace, 8);
                    if (value is OfdTextElement textElement)
                    {
                        Markup(3, 4, 8); EmitString(Namespace, 3); // possible fallback font declaration
                        if (textElement.Runs.Count == 0) { Markup(1, 2, 4); EmitString(Namespace); }
                    }
                    break;
                case "run": Markup(1, 2, 4); EmitString(Namespace); break;
                case "font": Markup(3, 4, 8); EmitString(Namespace, 3); break;
                case "attachment": Markup(2, 8, 16); EmitString(Namespace, 2); break;
                case "annotation": Markup(8, 24, 24); EmitString(Namespace, 8); EmitString(package.DocumentEntryPath ?? package.Options.DocumentId, 8); break;
                case "compositeDefinition": Markup(4, 12, 12); EmitString(Namespace, 4); break;
                case "clipCollection": case "clip": case "area": Markup(4, 12, 12); EmitString(Namespace, 4); break;
                case "appearance": Markup(2, 16, 4); EmitString(Namespace, 2); break;
                case "navigation": Markup(2, 8, 8); EmitString(Namespace, 2); break;
                case "tagKey": Markup(1, 0, 2); EmitString(Namespace); break;
            }
        }
        public void Bytes(string role, byte[] value, object? owner = null)
        {
            token.ThrowIfCancellationRequested();
            if (role == "entryBytes" && material.xmlEntries.TryGetValue(value, out var work))
                material.ChargeOutput(6L * value.LongLength + work.Nodes * (4L * (material.limits.MaxDepth + 16L) + 4 + 6L * Namespace.Length));
            else if (owner is OfdImageElement image)
            {
                var key = package.SourceIdentity?.TryImage(image, out var imported) == true
                    ? ((object)imported, (string?)null, (string?)null) : ((object)value, image.MediaType, image.FileName);
                if (images.Add(key)) material.ChargeOutput(value.LongLength);
            }
            else if (owner is not OfdAttachment attachment || !attachment.IsExternal) material.ChargeOutput(value.LongLength);
        }
        public void Matrix(double[]? value) { }
        public void Path(string? value) { }
        public void Xml(string? value)
        {
            if (string.IsNullOrWhiteSpace(value)) return;
            material.ChargeCopy(2L * value!.Length, token);
            using var reader = XmlReader.Create(new StringReader(value), new XmlReaderSettings
            { DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null, MaxCharactersInDocument = material.limits.MaxSourceXmlBytes });
            while (reader.Read()) { token.ThrowIfCancellationRequested(); material.Structure.Depth(reader.Depth + 1); Markup(1, 0, 0); EmitString(Namespace); }
        }
        private void EmitString(string? value, int positions = 1)
        { token.ThrowIfCancellationRequested(); if (value is not null) material.ChargeOutput(6L * value.Length * positions); }
        private void Markup(int elements, int numbers, int attributes)
        {
            material.Structure.NodeWork(elements, 1);
            // Fixed 0.### formatting is <=326 ASCII chars per double;
            // lexical source names/values are covered by six times input length.
            var literal = elements * (2L * (64 + 3) + 4L * (material.limits.MaxDepth + 16L) + 4);
            material.ChargeOutput(literal + attributes * (64L + 4) + numbers * 326L);
        }
    }

    internal long ValidateActualEntries(IDictionary<string, byte[]> entries, CancellationToken token, long? reservation = null)
    {
        long actual = 0;
        foreach (var entry in entries) { token.ThrowIfCancellationRequested(); actual = Add(actual, entry.Value.LongLength, "actual logical output"); }
        if (actual > (reservation ?? OutputUpperBoundBytes)) throw new InvalidDataException("Writer output exceeded its preflight upper bound.");
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
