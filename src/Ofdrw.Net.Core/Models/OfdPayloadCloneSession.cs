using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Only the whole-owned-clone factory supplies this session. It preserves byte
// array identity within one independent candidate, never across candidates or
// by comparing contents. It is separate from the model ownership map.
internal sealed class OfdPayloadCloneSession : IDisposable
{
    private readonly OfdControlledEditContext context;
    private Dictionary<byte[], byte[]> copies;
    private int capacity;
    private bool disposed;

    private OfdPayloadCloneSession(OfdControlledEditContext context)
    { this.context = context; copies = new Dictionary<byte[], byte[]>(OfdReferenceIdentity.Instance); }

    internal static OfdPayloadCloneSession Create(OfdControlledEditContext context, CancellationToken token)
    {
        context.AssertActive();
        context.Material.ChargeCopy(512, token); // session and initially empty dictionary
        return new OfdPayloadCloneSession(context);
    }

    internal byte[] Copy(byte[] original, OfdModelMaterialization material, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (disposed || !ReferenceEquals(material, context.Material))
            throw new InvalidOperationException("Payload clone session is inactive or belongs to another operation.");
        if (original is null) throw new InvalidDataException("Null payload array.");
        // Fund even a colliding identity-key lookup before traversing the table.
        material.ChargeWork(checked((copies.Count + 1L) * IntPtr.Size), token);
        if (copies.TryGetValue(original, out var existing)) return existing;
        if (copies.Count == capacity)
        {
            var nextCapacity = capacity == 0 ? 1 : checked(capacity * 2);
            // Rebuild explicitly before the requested capacity is exceeded. The
            // allowance covers .NET dictionary buckets, entries, prime rounding
            // and migration; Add below cannot trigger an unreserved resize.
            material.ChargeCopy(checked(256L + 128L * nextCapacity), token);
            var next = new Dictionary<byte[], byte[]>(nextCapacity, OfdReferenceIdentity.Instance);
            foreach (var pair in copies)
            {
                material.ChargeWork(checked((next.Count + 1L) * IntPtr.Size), token);
                next.Add(pair.Key, pair.Value);
            }
            copies = next; capacity = nextCapacity;
        }
        var clone = material.AllocateOwnedPayload(original.Length, token);
        for (var offset = 0; offset < original.Length;)
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            var count = Math.Min(64 * 1024, original.Length - offset);
            Array.Copy(original, offset, clone, offset, count); offset += count;
        }
        material.ChargeWork(checked((copies.Count + 1L) * IntPtr.Size), token);
        copies.Add(original, clone);
        return clone;
    }

    public void Dispose() { disposed = true; copies.Clear(); }
}
