using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal sealed class OfdObservedPayloadIdentity
{
    internal long ByteLength { get; }
    internal string Sha256Hex { get; }
    private OfdObservedPayloadIdentity(long length, string hash) { ByteLength = length; Sha256Hex = hash; }
    internal static OfdObservedPayloadIdentity FromActual(byte[] bytes, byte[] digest,
        OfdModelMaterialization material, CancellationToken token)
    {
        if (digest is null || digest.Length != 32) throw new InvalidDataException("Actual SHA256 must contain 32 bytes.");
        material.ChargeCopy(512, token);
        return new(bytes.LongLength, BitConverter.ToString(digest).Replace("-", string.Empty).ToLowerInvariant());
    }
}

// Facts from one actual, current inventory walk. They may be retained as byte
// identity targets, but cannot satisfy any later Assert. Mutable arrays/maps,
// the package and context are never returned by this object.
internal sealed class OfdCurrentPayloadEvidence
{
    private readonly OfdDocumentPackage package;
    private readonly OfdControlledEditContext context;
    private readonly Dictionary<byte[], OfdObservedPayloadIdentity> observed;
    private int consumed;
    private OfdCurrentPayloadEvidence(OfdDocumentPackage package, OfdControlledEditContext context,
        Dictionary<byte[], OfdObservedPayloadIdentity> observed)
    { this.package = package; this.context = context; this.observed = observed; }
    internal static OfdCurrentPayloadEvidence Issue(OfdDocumentPackage package, OfdControlledEditContext context,
        Dictionary<byte[], byte[]> actualDigests, CancellationToken token)
    {
        context.AssertActive(); context.Material.ChargeCopy(checked(512L + 256L * actualDigests.Count), token);
        var records = new Dictionary<byte[], OfdObservedPayloadIdentity>();
        foreach (var entry in actualDigests)
        {
            token.ThrowIfCancellationRequested();
            records.Add(entry.Key, OfdObservedPayloadIdentity.FromActual(entry.Key, entry.Value, context.Material, token));
        }
        return new(package, context, records);
    }
    internal OfdObservedPayloadIdentity RequireObserved(OfdDocumentPackage exactPackage,
        OfdControlledEditContext exactContext, byte[] exactObservedArray, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (Volatile.Read(ref consumed) != 0 || !ReferenceEquals(package, exactPackage) || !ReferenceEquals(context, exactContext) ||
            !observed.TryGetValue(exactObservedArray, out var identity) || identity.ByteLength != exactObservedArray.LongLength)
            throw new InvalidDataException("Array was not observed by this exact current inventory pass.");
        return identity;
    }
    internal OfdObservedPayloadIdentity? FindObservedCopySource(OfdDocumentPackage exactPackage,
        OfdControlledEditContext exactContext, byte[] exactArray, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (Volatile.Read(ref consumed) != 0 || !ReferenceEquals(package, exactPackage) || !ReferenceEquals(context, exactContext))
            throw new InvalidOperationException("Copy source lookup is outside its private current inventory pass.");
        if (!observed.TryGetValue(exactArray, out var value)) return null;
        if (value.ByteLength != exactArray.LongLength) throw new InvalidDataException("Observed copy source length differs.");
        return value;
    }
    internal void CompletePrivateIssuance()
    {
        context.AssertActive();
        if (Interlocked.CompareExchange(ref consumed, 1, 0) != 0)
            throw new InvalidOperationException("Current inventory evidence was already consumed.");
        observed.Clear();
    }

}

internal sealed partial class OfdSerializedGuard
{
    internal OfdCurrentPayloadEvidence AssertUnchangedAndObservePayloads(OfdDocumentPackage exactPackage,
        OfdControlledEditContext exactContext, CancellationToken token)
    {
        RequireBinding(exactPackage, exactContext);
        try
        {
            var sink = new GuardSink(context.Material, token, entries, true);
            OfdSerializedInventory.Walk(package, sink, context.Material, token, rejectRepeatedOwners: true);
            sink.Complete(); return sink.IssueCurrentPayloadEvidence(package, context, token);
        }
        catch { context.Fail(); throw; }
    }
    internal static OfdSerializedGuard CaptureAndObservePayloads(OfdDocumentPackage actualPackage,
        OfdControlledEditContext exactContext, out OfdCurrentPayloadEvidence actualEvidence, CancellationToken token)
    {
        exactContext.AssertActive();
        try
        {
            exactContext.Material.InspectStorage(actualPackage, token);
            exactContext.Material.ChargeCopy(128, token);
            var entries = new List<Entry>();
            var sink = new GuardSink(exactContext.Material, token, entries, false);
            OfdSerializedInventory.Walk(actualPackage, sink, exactContext.Material, token, rejectRepeatedOwners: true);
            actualEvidence = sink.IssueCurrentPayloadEvidence(actualPackage, exactContext, token);
            return new OfdSerializedGuard(actualPackage, exactContext, entries);
        }
        catch { exactContext.Fail(); throw; }
    }
}
