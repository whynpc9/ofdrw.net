using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Archive;

// Issued only inside the actual ZipArchive enumeration. It includes physical
// entries the ordinary model projection ignores, and retains the exact arrays
// that the controlled loader allocated. Never accepted from a caller dictionary.
internal sealed class OfdOwnedZipCensus
{
    internal readonly struct Entry(string physicalName, string? normalizedName,
        long compressedLength, long expandedLength, bool directory, byte[]? bytes)
    {
        internal readonly string PhysicalName = physicalName;
        internal readonly string? NormalizedName = normalizedName;
        internal readonly long CompressedLength = compressedLength, ExpandedLength = expandedLength;
        internal readonly bool Directory = directory;
        internal readonly byte[]? Bytes = bytes;
    }
    private readonly OfdControlledEditContext context;
    private readonly List<Entry> entries = new();
    private int expectedCount = -1, sealedState;
    private OfdOwnedZipCensus(OfdControlledEditContext context) { this.context = context; }
    internal static OfdOwnedZipCensus Begin(OfdControlledEditContext context, CancellationToken token)
    {
        context.AssertActive(); context.Material.ChargeCopy(1024, token); return new(context);
    }
    internal void SetExpectedPhysicalCount(int count, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (expectedCount != -1 || count < 0) throw new InvalidOperationException("Census count was already bound.");
        expectedCount = count;
    }
    internal void Record(string physical, string? normalized, long compressed, long expanded,
        bool directory, byte[]? actualBuffer, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (sealedState != 0 || expectedCount < 0 || entries.Count >= expectedCount)
            throw new InvalidDataException("Physical ZIP census exceeds its actual parser count.");
        context.Material.ChargeCopy(checked(512L + 4L * physical.Length), token);
        context.Material.ChargeWork(checked(32L + 2L * physical.Length), token);
        if (directory && (expanded != 0 || !physical.EndsWith("/", StringComparison.Ordinal)) ||
            !directory && (normalized is null || actualBuffer is null || actualBuffer.LongLength != expanded))
            throw new InvalidDataException("Physical ZIP census contains hidden or inconsistent payload.");
        entries.Add(new Entry(physical, normalized, compressed, expanded, directory, actualBuffer));
    }
    internal void Seal(CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (entries.Count != expectedCount || Interlocked.CompareExchange(ref sealedState, 1, 0) != 0)
            throw new InvalidDataException("Physical ZIP census is incomplete or already sealed.");
    }
    internal Entry[] TakeForPrivateVerifier(OfdControlledEditContext exactContext, CancellationToken token)
    {
        context.AssertActive(); token.ThrowIfCancellationRequested();
        if (!ReferenceEquals(context, exactContext) || Volatile.Read(ref sealedState) != 1)
            throw new InvalidOperationException("Unsealed or foreign physical ZIP census.");
        context.Material.ChargeCopy(checked(128L + 64L * entries.Count), token); return entries.ToArray();
    }
}
