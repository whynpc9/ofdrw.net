using System;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Reader.Readers;

internal sealed class OfdOwnedInputRead : IDisposable
{
    private readonly OfdOwnedModelsLease lease;
    private readonly OfdOwnedInputSnapshot snapshot;
    private int state;
    private OfdOwnedInputRead(OfdOwnedModelsLease lease, OfdOwnedInputSnapshot snapshot)
    { this.lease = lease; this.snapshot = snapshot; }
    internal static OfdOwnedInputRead Issue(OfdOwnedModelsLease lease, OfdOwnedInputSnapshot snapshot) => new(lease, snapshot);
    internal void Consume(out OfdOwnedModelsLease actualLease, out OfdOwnedInputSnapshot actualSnapshot)
    {
        try
        {
            if (Interlocked.CompareExchange(ref state, 1, 0) != 0) throw new InvalidOperationException("Initial read was already consumed or disposed.");
            actualLease = lease; actualSnapshot = snapshot;
        }
        catch { lease.Fail(); throw; }
    }
    public void Dispose()
    {
        if (Interlocked.CompareExchange(ref state, 2, 0) != 0) return;
        lease.Fail(); snapshot.Dispose();
    }
}

internal static partial class OfdOwnedReader
{
    internal static async Task<OfdOwnedInputRead> ReadInitialAsync(OfdOwnedJobTicket ticket,
        string originalInputPath, OfdPackageLoadOptions options, CancellationToken token)
    {
        OfdOwnedInputSnapshot? original = null;
        try
        {
            var claim = OfdOwnedSaveAuthority.BeginInputRead(ticket, out var context);
            context.Material.ChargeCopy(512, token);
            var limits = OfdPackageLoader.SnapshotOptions(options);
            original = await OfdOwnedInputSnapshot.CaptureActualAsync(claim, context, originalInputPath,
                512L * 1024 * 1024, limits.MaxInputBytes, token).ConfigureAwait(false);
            var acquisition = OfdOwnedSaveAuthority.BeginLoaderAcquisition(claim);
            original.EnterActualInitialReader(claim, out var stream, out var capture, token);
            var actual = await new OfdReader().ReadOwnedInitialAsync(stream, limits, context, acquisition, token).ConfigureAwait(false);
            var lease = OfdOwnedSaveAuthority.CompleteInputRead(claim, actual, acquisition, capture, token);
            original.CompleteActualInitialReader(claim);
            context.Material.ChargeCopy(256, token);
            return OfdOwnedInputRead.Issue(lease, original);
        }
        catch { ticket.Owner.Fail(); original?.AbortActualInitialReader(); original?.Dispose(); throw; }
    }
    internal static async Task<OfdActualReadback> ReadOutputAsync(OfdOwnedWrittenSpool written, CancellationToken token)
    {
        written.EnterActualOutputReader(out var stream, out var context, out var census, token);
        try
        {
            var actualR = await new OfdReader().ReadOwnedOutputAsync(stream, context, census, token).ConfigureAwait(false);
            return written.CompleteActualOutputReader(actualR, census, token);
        }
        catch { context.Fail(); written.AbortActualOutputReader(); throw; }
    }
    internal static void ConsumeInitial(OfdOwnedInputRead actualRead,
        out OfdOwnedModelsLease lease, out OfdOwnedInputSnapshot snapshot) => actualRead.Consume(out lease, out snapshot);
}
