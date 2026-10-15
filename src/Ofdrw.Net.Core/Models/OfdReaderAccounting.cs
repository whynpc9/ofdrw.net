using System;
using System.IO;
using System.Collections.Generic;
using Ofdrw.Net.Core.IO;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// Construction-site accounting only. SerializedInventory remains the field table.
internal sealed partial class OfdModelMaterialization
{
    private OfdControlledEditContext? controlledContext;
    private OfdControlledEditContext? readContext;
    // XML admission persists across read, capture, clone, prepare and write.
    // Exiting a Reader construction scope never downgrades this policy.
    internal void BindControlledOperation(OfdControlledEditContext context)
    {
        if (controlledContext is not null || !ReferenceEquals(context.Material, this))
            throw new InvalidOperationException("Materialization already bound or belongs to another controlled operation.");
        context.AssertActive(); controlledContext = context;
    }
    private bool HasControlledXmlPolicy => controlledContext is not null;
    internal bool IsControlledRead => readContext is not null;
    internal long ReadReservedStorageBytes { get; private set; }
    internal long ReadAdoptedStorageBytes { get; private set; }
    internal long ReadMetadataReservations { get; private set; }
    internal long ReadEntryAllocations { get; private set; }
    internal long ReadDomAllocations { get; private set; }
    internal long ReadSourceClones { get; private set; }
    internal long ReadSourceStrings { get; private set; }
    internal long ReadModelReservations { get; private set; }
    internal long ReadEntryBytes { get; private set; }
    internal long ReadWriterAdmissions { get; private set; }
    internal void PrepareReadXmlWriter(CancellationToken token)
    {
        if (!IsControlledRead) return;
        // .NET 10 XmlEncodedRawTextWriter's 6144-char buffer plus wrappers.
        ReserveReadModel(32768, token); ReadWriterAdmissions++;
    }
    internal long ReadParserAdmissions { get; private set; }
    internal long ReadAnnotationClones { get; private set; }
    internal long ReadNumericParses { get; private set; }

    // .NET 10 XmlTextReaderImpl: even a one-byte token position can fund a
    // NodeData, doubled node/sort/name/namespace arrays and decoded/name copies.
    // 512 bytes per input byte plus fixed reader buffers is deliberately loose;
    // this is reserved construction work, not a measurement of managed memory.
    internal void PrepareReadXmlParser(long inputBytes, CancellationToken token)
    {
        if (!HasControlledXmlPolicy) return;
        controlledContext!.AssertActive();
        ChargeCopy(checked(65536L + 512L * inputBytes), token); ReadModelReservations++;
        ChargeWork(inputBytes, token); ReadParserAdmissions++;
    }
    internal void PrepareReadAnnotationClone(IEnumerable<XElement> roots, CancellationToken token)
    {
        if (!IsControlledRead) return;
        long nodes = 0, attributes = 0;
        foreach (var root in roots)
            foreach (var node in root.DescendantNodesAndSelf())
            {
                AssertReadActive(token); nodes++;
                if (node is XElement element)
                    foreach (var attribute in element.Attributes()) { AssertReadActive(token); attributes++; }
            }
        // LINQ to XML clones shells and aliases the immutable values/names.
        ReserveReadModel(checked(2048L + 256L * nodes + 192L * attributes), token);
        ReadAnnotationClones++;
    }
    internal void PrepareReadNumericParse(string? value, CancellationToken token)
    {
        if (!IsControlledRead || value is null) return;
        ReserveReadModel(checked(512L + 32L * value.Length), token); ReadNumericParses++;
    }
    internal string ReadConcat(string left, string right, CancellationToken token)
    {
        if (!IsControlledRead) return left + right;
        var length = checked((long)left.Length + right.Length);
        ReserveReadModel(checked(256L + 8L * length), token);
        var storage = ReserveReadStorage(checked(2L * length), token);
        var result = left + right; AdoptReadString(result, storage, token); return result;
    }
    internal string ReadDirectorySuffix(string path, string suffix, CancellationToken token)
    {
        if (!IsControlledRead) return OfdPackagePath.GetDirectory(path) + suffix;
        var length = checked((long)path.Length + suffix.Length);
        ReserveReadModel(checked(512L + 16L * length), token);
        var storage = ReserveReadStorage(checked(2L * length), token);
        var result = OfdPackagePath.GetDirectory(path) + suffix; AdoptReadString(result, storage, token); return result;
    }

    internal IDisposable BeginControlledRead(OfdControlledEditContext context, CancellationToken token)
    {
        if (readContext is not null) throw new InvalidOperationException("A controlled read is already active on this operation.");
        if (!ReferenceEquals(context.Material, this)) throw new InvalidOperationException("Controlled read material belongs to another context.");
        context.AssertActive(); token.ThrowIfCancellationRequested(); ChargeCopy(128, token);
        readContext = context; return new ReadScope(this);
    }
    private sealed class ReadScope(OfdModelMaterialization owner) : IDisposable
    {
        private bool disposed;
        public void Dispose() { if (disposed) return; disposed = true; owner.readContext = null; }
    }
    internal void AssertReadActive(CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        (readContext ?? throw new InvalidOperationException("Accounted read requires its exact active materialization scope.")).AssertActive();
    }
    internal void ReserveReadModel(long bytes, CancellationToken token)
    {
        if (!IsControlledRead) return;
        AssertReadActive(token); ChargeCopy(bytes, token); ReadModelReservations++;
    }
    internal void AdmitReadMetadata(long allowance, CancellationToken token)
    {
        AssertReadActive(token); ChargeCopy(allowance, token); ReadMetadataReservations++;
    }
    internal byte[] AllocateReadEntry(int length, CancellationToken token)
    {
        AssertReadActive(token); ChargeCopy(128L + length, token);
        var storage = ReserveReadStorage(length, token);
        ReadEntryAllocations++; // after quota admission, immediately before allocation
        var bytes = new byte[length]; AdoptReadStorage(bytes, length, storage, token);
        ReadEntryBytes = checked(ReadEntryBytes + length); return bytes;
    }

    // Private material-bound tickets cannot refund a reservation or reset a cap.
    internal sealed class ReadStorageReservation
    {
        internal readonly OfdModelMaterialization Owner;
        internal long Remaining;
        internal ReadStorageReservation(OfdModelMaterialization owner, long bytes) { Owner = owner; Remaining = bytes; }
    }
    internal ReadStorageReservation ReserveReadStorage(long bytes, CancellationToken token)
    {
        AssertReadActive(token); ChargeCopy(64, token);
        StorageBytes = Add(StorageBytes, bytes, "read retained-storage reservation");
        ReadReservedStorageBytes = checked(ReadReservedStorageBytes + bytes);
        return new ReadStorageReservation(this, bytes);
    }
    internal void AdoptReadStorage(object value, long bytes, ReadStorageReservation reservation, CancellationToken token)
    {
        AssertReadActive(token);
        if (!ReferenceEquals(reservation.Owner, this)) throw new InvalidDataException("Foreign read storage reservation.");
        if (stored.Contains(value)) return;
        if (bytes < 0 || bytes > reservation.Remaining) throw OfdTraversalBudget.Limit("Read construction exceeded its retained-storage reservation.");
        reservation.Remaining -= bytes; stored.Add(value); ReadAdoptedStorageBytes = checked(ReadAdoptedStorageBytes + bytes);
    }
    internal void AdoptReadString(string value, ReadStorageReservation reservation, CancellationToken token)
        => AdoptReadStorage(value, checked(2L * value.Length), reservation, token);

    // One allowance funds one DOM construction; repeated ReadXml calls pay again.
    internal ReadStorageReservation? PrepareReadDom(long byteLength, long nodes, long attributes, CancellationToken token)
    {
        if (!IsControlledRead) return null;
        AssertReadActive(token);
        // Node/attribute shells and table growth are separate from decoded text
        // and StringBuilder/copy positions. These are reserved workload bytes.
        ChargeCopy(checked(32768L + 256L * nodes + 192L * attributes + 6L * byteLength), token);
        var storage = ReserveReadStorage(checked(1024L + 2L * byteLength), token);
        ReadDomAllocations++; return storage;
    }
    internal void AdoptReadDom(XDocument document, ReadStorageReservation? reservation, CancellationToken token)
    {
        if (reservation is null) return;
        foreach (var node in document.DescendantNodes())
        {
            AssertReadActive(token);
            if (node is XElement element)
            {
                Name(element.Name);
                foreach (var attribute in element.Attributes()) { AssertReadActive(token); Name(attribute.Name); AdoptReadString(attribute.Value, reservation, token); }
            }
            else if (node is XText text) AdoptReadString(text.Value, reservation, token);
            else if (node is XComment comment) AdoptReadString(comment.Value, reservation, token);
            else if (node is XProcessingInstruction instruction) { AdoptReadString(instruction.Target, reservation, token); AdoptReadString(instruction.Data, reservation, token); }
        }
        void Name(XName name) { AdoptReadString(name.LocalName, reservation, token); AdoptReadString(name.NamespaceName, reservation, token); }
    }

    internal string ReadXmlValue(XElement source, CancellationToken token)
    {
        if (!IsControlledRead) return source.Value;
        ReserveReadModel(256, token); long length = 0;
        foreach (var node in source.DescendantNodes())
        { AssertReadActive(token); if (node is XText text) length = checked(length + text.Value.Length); }
        ChargeCopy(checked(6L * length), token); ChargeWork(checked(4L * length), token);
        var storage = ReserveReadStorage(checked(2L * length), token);
        var value = source.Value; AdoptReadString(value, storage, token); return value;
    }

    internal void ReadSourceCloneCheckpoint() { if (IsControlledRead) ReadSourceClones++; }
    internal void ReadSourceStringCheckpoint() { if (IsControlledRead) ReadSourceStrings++; }
}
