using System;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Input values, not ownership or freshness authority. The batch copies Slots.
internal readonly struct OfdRetainedTextEdit(OfdTextElement owner, OfdTextSlotEdit[] slots)
{
    internal OfdTextElement Owner { get; } = owner;
    internal OfdTextSlotEdit[] Slots { get; } = slots;
}

// A prepaid, exact-array snapshot. There is no record/model/map getter and no
// arbitrary collection, mutation delegate or transferable validation stamp.
internal sealed class OfdRetainedTextEditBatch : IDisposable
{
    private OfdDocumentPackage? source;
    private OfdSerializedGuard? sourceGuard;
    private OfdControlledEditContext? context;
    private OfdRetainedTextEdit[]? edits;
    private int state;

    private OfdRetainedTextEditBatch(OfdDocumentPackage source, OfdSerializedGuard sourceGuard,
        OfdControlledEditContext context, OfdRetainedTextEdit[] edits)
    { this.source = source; this.sourceGuard = sourceGuard; this.context = context; this.edits = edits; }

    internal static OfdRetainedTextEditBatch Capture(OfdDocumentPackage source, OfdSerializedGuard sourceGuard,
        OfdControlledEditContext context, OfdRetainedTextEdit[] edits, CancellationToken token = default)
    {
        if (context is null) throw new ArgumentNullException(nameof(context));
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            if (source is null || sourceGuard is null || edits is null)
                throw new ArgumentNullException(source is null ? nameof(source) : sourceGuard is null ? nameof(sourceGuard) : nameof(edits));
            sourceGuard.RequireBinding(source, context);
            if (edits.Length == 0) throw new InvalidDataException("Continuous text editing requires a nonempty edit batch.");
            var material = context.Material;
            material.ChargeCopy(256L + 128L * edits.Length, token);
            var frozen = new OfdRetainedTextEdit[edits.Length];
            for (var index = 0; index < frozen.Length; index++)
            {
                token.ThrowIfCancellationRequested(); material.Structure.Reference();
                var edit = edits[index];
                if (edit.Owner is null || edit.Slots is null) throw new InvalidDataException("A retained edit requires an original owner and concrete slots.");
                material.ChargeCopy(128L + 128L * edit.Slots.Length, token);
                var slots = new OfdTextSlotEdit[edit.Slots.Length];
                for (var ordinal = 0; ordinal < slots.Length; ordinal++)
                {
                    token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
                    var slot = edit.Slots[ordinal];
                    if (slot.Ordinal != ordinal || slot.RawOfdText is null ||
                        double.IsNaN(slot.X) || double.IsInfinity(slot.X) || double.IsNaN(slot.Y) || double.IsInfinity(slot.Y))
                        throw new InvalidDataException("Retained edits require ordered slots, raw text and finite explicit anchors.");
                    material.StoreString(slot.RawOfdText); material.StoreString(slot.DeltaX); material.StoreString(slot.DeltaY);
                    slots[ordinal] = slot; // Strings are immutable; mutable arrays are independent.
                }
                frozen[index] = new OfdRetainedTextEdit(edit.Owner, slots);
            }
            token.ThrowIfCancellationRequested(); context.AssertActive();
            return new OfdRetainedTextEditBatch(source, sourceGuard, context, frozen);
        }
        catch { context.Fail(); throw; }
    }

    internal void Claim(OfdDocumentPackage exactSource, OfdSerializedGuard exactGuard,
        OfdControlledEditContext exactContext, CancellationToken token)
    {
        token.ThrowIfCancellationRequested(); exactContext.AssertActive();
        if (Interlocked.CompareExchange(ref state, 1, 0) != 0 ||
            !ReferenceEquals(source, exactSource) || !ReferenceEquals(sourceGuard, exactGuard) || !ReferenceEquals(context, exactContext))
            throw new InvalidOperationException("Retained edit batch is consumed or belongs to another source, guard or context.");
    }

    // Binding only: the owned job uses this before registering its concrete
    // compiler batch. It does not establish source or candidate freshness.
    internal void RequireUnclaimedBinding(OfdDocumentPackage exactSource, OfdSerializedGuard exactGuard,
        OfdControlledEditContext exactContext)
    {
        exactContext.AssertActive();
        if (Volatile.Read(ref state) != 0 || edits is null ||
            !ReferenceEquals(source, exactSource) || !ReferenceEquals(sourceGuard, exactGuard) ||
            !ReferenceEquals(context, exactContext))
            throw new InvalidOperationException("Retained edit batch is not an unclaimed batch of this owned job.");
    }

    internal void RequireCompleteAffectedRoots(OfdTextElement[] supplied, CancellationToken token)
    {
        var actual = edits ?? throw new InvalidOperationException("Consumed edit batch.");
        var package = source ?? throw new InvalidOperationException("Consumed edit batch.");
        var material = context!.Material;
        material.ChargeCopy(checked(512L + 256L * actual.Length + package.Pages.Count), token);
        var owners = new System.Collections.Generic.HashSet<OfdTextElement>();
        foreach (var edit in actual)
            if (!owners.Add(edit.Owner)) throw new InvalidDataException("Repeated edit owner.");
        var affected = new bool[package.Pages.Count]; var found = 0;
        for (var pageIndex = 0; pageIndex < package.Pages.Count; pageIndex++)
            foreach (var element in package.Pages[pageIndex].Elements)
            {
                material.ChargeWork(16, token);
                if (element is OfdTextElement text && owners.Contains(text))
                { affected[pageIndex] = true; found++; }
            }
        if (found != owners.Count) throw new InvalidDataException("Edit owners do not bind unique source page roots.");
        var ordinal = 0;
        for (var pageIndex = 0; pageIndex < package.Pages.Count; pageIndex++)
        {
            if (!affected[pageIndex]) continue;
            foreach (var element in package.Pages[pageIndex].Elements)
            {
                material.ChargeWork(16, token);
                if (element is not OfdTextElement text) continue;
                if (ordinal >= supplied.Length || !ReferenceEquals(supplied[ordinal++], text))
                    throw new InvalidDataException("Incomplete or reordered affected-page text-root coverage.");
            }
        }
        if (ordinal != supplied.Length) throw new InvalidDataException("Foreign affected-page text-root coverage.");
    }

    internal void Apply(OfdOwnedEditCandidate candidate, CancellationToken token)
    {
        if (Volatile.Read(ref state) != 1 || edits is null || !ReferenceEquals(context, candidate.Context))
            throw new InvalidOperationException("Retained edit batch is not claimed for this candidate.");
        foreach (var edit in edits)
        {
            token.ThrowIfCancellationRequested();
            // The real clone map and patcher's source contract reject foreign or
            // repeated roots; no caller ID/path can stand in for this owner.
            candidate.ApplyTextSlots(candidate.GetMapped(edit.Owner), edit.Slots, token);
        }
    }

    internal void Complete()
    {
        if (Interlocked.CompareExchange(ref state, 2, 1) != 1)
            throw new InvalidOperationException("Retained edit batch was revoked during continuation.");
        Clear();
    }

    public void Dispose()
    {
        Interlocked.Exchange(ref state, 2);
        Clear();
    }

    private void Clear()
    {
        var release = Interlocked.Exchange(ref edits, null);
        if (release is not null) Array.Clear(release, 0, release.Length);
        source = null; sourceGuard = null; context = null;
    }
}
