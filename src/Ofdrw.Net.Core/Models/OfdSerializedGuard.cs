using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// One operation identity; friend assemblies can meter their own stages through
// Material, but cannot replace it inside a candidate certificate.
internal sealed class OfdControlledEditContext : IDisposable
{
    private int state; // active, disposed, failed
    private int cloneIssued;
    internal OfdModelMaterialization Material { get; }
    internal OfdControlledEditContext(OfdModelLimits limits)
    {
        Material = new OfdModelMaterialization(limits); Material.BindControlledOperation(this);
    }
    internal void AssertActive()
    {
        if (Volatile.Read(ref state) == 1) throw new ObjectDisposedException(nameof(OfdControlledEditContext));
        if (Volatile.Read(ref state) == 2) throw new InvalidOperationException("Controlled edit operation has failed and cannot be reset.");
    }
    internal void Fail() => Interlocked.CompareExchange(ref state, 2, 0);
    internal void ClaimClone()
    {
        AssertActive();
        if (Interlocked.CompareExchange(ref cloneIssued, 1, 0) != 0) throw new InvalidOperationException("This operation has already issued its owned clone.");
    }
    internal OfdSerializedGuard CaptureSource(OfdDocumentPackage source, CancellationToken token = default)
    {
        AssertActive();
        try
        {
            if (source is null) throw new ArgumentNullException(nameof(source));
            Material.GuardPreservedXml(source, token); return OfdSerializedGuard.Capture(source, this, token);
        }
        catch { Fail(); throw; }
    }
    public void Dispose() => Interlocked.Exchange(ref state, 1);
}

internal sealed class OfdSerializedGuard
{
    private readonly OfdDocumentPackage package;
    private readonly OfdControlledEditContext context;
    private readonly List<Entry> entries;
    private readonly struct Entry(string role, object? value, bool identity)
    { internal readonly string Role = role; internal readonly object? Value = value; internal readonly bool Identity = identity; }
    private OfdSerializedGuard(OfdDocumentPackage package, OfdControlledEditContext context, List<Entry> entries)
    { this.package = package; this.context = context; this.entries = entries; }

    internal static OfdSerializedGuard Capture(OfdDocumentPackage package, OfdControlledEditContext context, CancellationToken token)
    {
        context.AssertActive(); context.Material.InspectStorage(package, token);
        context.Material.ChargeCopy(128, token);
        var entries = new List<Entry>(); var sink = new GuardSink(context.Material, token, entries, false);
        OfdSerializedInventory.Walk(package, sink, context.Material, token, rejectRepeatedOwners: true);
        return new OfdSerializedGuard(package, context, entries);
    }

    internal void AssertUnchanged(OfdDocumentPackage exactPackage, OfdControlledEditContext exactContext, CancellationToken token = default)
    {
        RequireBinding(exactPackage, exactContext);
        try
        {
            var sink = new GuardSink(context.Material, token, entries, true);
            OfdSerializedInventory.Walk(package, sink, context.Material, token, rejectRepeatedOwners: true); sink.Complete();
        }
        catch { context.Fail(); throw; }
    }

    // Identity only. This never establishes freshness or replaces a byte pass.
    internal void RequireBinding(OfdDocumentPackage exactPackage, OfdControlledEditContext exactContext)
    {
        context.AssertActive();
        if (!ReferenceEquals(package, exactPackage) || !ReferenceEquals(context, exactContext))
            throw new InvalidOperationException("Serialized guard belongs to another package or operation.");
    }

    private sealed class GuardSink(OfdModelMaterialization material, CancellationToken token, List<Entry> entries, bool comparing) : OfdSerializedInventory.ISink
    {
        private int position;
        private readonly Dictionary<byte[], byte[]> digests = new(); // fresh for every verification pass
        public void Item(string role, object? value, bool identity = false)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            // IEEE bits preserve negative zero and NaN payload changes; dates
            // include their offset, which DateTimeOffset.Equals alone ignores.
            if (value is double number) value = BitConverter.DoubleToInt64Bits(number);
            else if (value is DateTimeOffset date) value = (date.Ticks, date.Offset.Ticks);
            material.ChargeWork(value is string text ? 2L * text.Length : 16, token);
            if (!comparing) { material.ChargeCopy(128, token); entries.Add(new Entry(role, value, identity)); }
            else
            {
                if (position >= entries.Count) Changed(); var expected = entries[position++];
                if (role != expected.Role || identity != expected.Identity ||
                    (identity ? !ReferenceEquals(value, expected.Value) : !Equals(value, expected.Value))) Changed();
            }
        }
        public void Bytes(string role, byte[] value, object? owner = null)
        {
            if (value is null) throw new InvalidDataException("Null payload array.");
            Item(role, value, true); Item("length", value.LongLength);
            if (!digests.TryGetValue(value, out var digest))
            { material.ChargeCopy(64, token); digests.Add(value, digest = material.HashBytes(value, token)); }
            // Digest bytes are immutable private values; compare each byte as
            // scalars so the source buffer itself is never a baseline snapshot.
            for (var index = 0; index < digest.Length; index++) Item("digest", digest[index]);
        }
        public void Matrix(double[]? value)
        { Item("matrix", value, true); if (value is null) return; Item("length", value.Length); foreach (var number in value) Item("matrixNumber", number); }
        public void Xml(string? value) { }
        public void Path(string? value) { }
        internal void Complete() { if (position != entries.Count) Changed(); }
        private static void Changed() => throw new InvalidDataException("Serialized package input changed during the controlled operation.");
    }
}
