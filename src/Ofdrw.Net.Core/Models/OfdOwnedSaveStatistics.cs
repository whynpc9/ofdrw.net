using System;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Preallocated once-fill numeric records. No model/path/stream/context/claim/digest access.
internal sealed class OfdOwnedSaveStatistics
{
    private OfdOwnedSaveStatistics() { }
    internal static OfdOwnedSaveStatistics Preallocate() => new();
    internal bool Available { get; private set; }
    internal int Sequence { get; private set; }
    internal int PhaseCode { get; private set; }
    internal long BudgetBytes { get; private set; }
    internal long StorageBytes { get; private set; }
    internal long CloneCopyBytes { get; private set; }
    internal long ValidationWorkBytes { get; private set; }
    internal long OutputUpperBoundBytes { get; private set; }
    internal long NodeCount { get; private set; }
    internal long ReferenceCount { get; private set; }
    internal long SourceXmlBytes { get; private set; }
    internal long PathValueCount { get; private set; }
    internal long LogicalWriterReservationU { get; private set; }
    internal bool NoopIssued { get; private set; }
    internal long ActiveActualFileLengths { get; private set; }
    internal long ActiveReservedFilePositions { get; private set; }
    internal long PeakActualFileLengths { get; private set; }
    internal long PeakReservedFilePositions { get; private set; }
    internal bool AllActualPositionsObserved { get; private set; }
    internal bool InputSnapshotCreated { get; private set; }
    internal bool InputSnapshotRemoved { get; private set; }
    internal bool InputSnapshotPositionFullyObserved { get; private set; }
    internal long InputSnapshotReadActualBytes { get; private set; }
    internal long InputSnapshotReadReservedBytes { get; private set; }
    internal long InputSnapshotOriginalReadActualBytes { get; private set; }
    internal long InputSnapshotOriginalReadReservedBytes { get; private set; }
    internal long InputSnapshotWriteActualBytes { get; private set; }
    internal long InputSnapshotWriteReservedBytes { get; private set; }
    internal long InputSnapshotSeekActualCount { get; private set; }
    internal long InputSnapshotSeekReservedCount { get; private set; }
    internal long InputSnapshotFlushActualCount { get; private set; }
    internal long InputSnapshotFlushReservedCount { get; private set; }
    internal long InputSnapshotReservedHighWater { get; private set; }
    internal long InputSnapshotActualPositionHighWater { get; private set; }
    internal long InputSnapshotActualLengthHighWater { get; private set; }
    internal long InputSnapshotActiveActualLength { get; private set; }
    internal long InputSnapshotActiveReservedPosition { get; private set; }
    internal bool OutputSpoolCreated { get; private set; }
    internal bool OutputSpoolRemoved { get; private set; }
    internal bool OutputSpoolPositionFullyObserved { get; private set; }
    internal long OutputSpoolReadActualBytes { get; private set; }
    internal long OutputSpoolReadReservedBytes { get; private set; }
    internal long OutputSpoolOriginalReadActualBytes { get; private set; }
    internal long OutputSpoolOriginalReadReservedBytes { get; private set; }
    internal long OutputSpoolWriteActualBytes { get; private set; }
    internal long OutputSpoolWriteReservedBytes { get; private set; }
    internal long OutputSpoolSeekActualCount { get; private set; }
    internal long OutputSpoolSeekReservedCount { get; private set; }
    internal long OutputSpoolFlushActualCount { get; private set; }
    internal long OutputSpoolFlushReservedCount { get; private set; }
    internal long OutputSpoolReservedHighWater { get; private set; }
    internal long OutputSpoolActualPositionHighWater { get; private set; }
    internal long OutputSpoolActualLengthHighWater { get; private set; }
    internal long OutputSpoolActiveActualLength { get; private set; }
    internal long OutputSpoolActiveReservedPosition { get; private set; }
    internal bool StageCreated { get; private set; }
    internal bool StageRemoved { get; private set; }
    internal bool StagePositionFullyObserved { get; private set; }
    internal long StageReadActualBytes { get; private set; }
    internal long StageReadReservedBytes { get; private set; }
    internal long StageOriginalReadActualBytes { get; private set; }
    internal long StageOriginalReadReservedBytes { get; private set; }
    internal long StageWriteActualBytes { get; private set; }
    internal long StageWriteReservedBytes { get; private set; }
    internal long StageSeekActualCount { get; private set; }
    internal long StageSeekReservedCount { get; private set; }
    internal long StageFlushActualCount { get; private set; }
    internal long StageFlushReservedCount { get; private set; }
    internal long StageReservedHighWater { get; private set; }
    internal long StageActualPositionHighWater { get; private set; }
    internal long StageActualLengthHighWater { get; private set; }
    internal long StageActiveActualLength { get; private set; }
    internal long StageActiveReservedPosition { get; private set; }
    internal void FillActual(int sequence, int phase, long budget, OfdModelMaterialization material, OfdOwnedIoTelemetry telemetry)
    {
        Available = true; Sequence = sequence; PhaseCode = phase; BudgetBytes = budget;
        StorageBytes = material.StorageBytes; CloneCopyBytes = material.CloneCopyBytes;
        ValidationWorkBytes = material.ValidationWorkBytes; OutputUpperBoundBytes = material.OutputUpperBoundBytes;
        NodeCount = material.Structure.Nodes; ReferenceCount = material.Structure.References;
        SourceXmlBytes = material.Structure.SourceXmlBytes; PathValueCount = material.Structure.PathValues;
        LogicalWriterReservationU = telemetry.LogicalU; NoopIssued = telemetry.NoopIssued;
        ActiveActualFileLengths = telemetry.ActiveActualLengths; ActiveReservedFilePositions = telemetry.ActiveReservedPositions;
        PeakActualFileLengths = telemetry.PeakActualLengths; PeakReservedFilePositions = telemetry.PeakReservedPositions;
        AllActualPositionsObserved = telemetry.AllActualPositionsObserved;
        InputSnapshotCreated = telemetry.Input.Created;
        InputSnapshotRemoved = telemetry.Input.Removed;
        InputSnapshotPositionFullyObserved = telemetry.Input.PositionFullyObserved;
        InputSnapshotReadActualBytes = telemetry.Input.ReadActualBytes;
        InputSnapshotReadReservedBytes = telemetry.Input.ReadReservedBytes;
        InputSnapshotOriginalReadActualBytes = telemetry.Input.OriginalReadActualBytes;
        InputSnapshotOriginalReadReservedBytes = telemetry.Input.OriginalReadReservedBytes;
        InputSnapshotWriteActualBytes = telemetry.Input.WriteActualBytes;
        InputSnapshotWriteReservedBytes = telemetry.Input.WriteReservedBytes;
        InputSnapshotSeekActualCount = telemetry.Input.SeekActualCount;
        InputSnapshotSeekReservedCount = telemetry.Input.SeekReservedCount;
        InputSnapshotFlushActualCount = telemetry.Input.FlushActualCount;
        InputSnapshotFlushReservedCount = telemetry.Input.FlushReservedCount;
        InputSnapshotReservedHighWater = telemetry.Input.ReservedHighWater;
        InputSnapshotActualPositionHighWater = telemetry.Input.ActualPositionHighWater;
        InputSnapshotActualLengthHighWater = telemetry.Input.ActualLengthHighWater;
        InputSnapshotActiveActualLength = telemetry.Input.ActiveActualLength;
        InputSnapshotActiveReservedPosition = telemetry.Input.ActiveReservedPosition;
        OutputSpoolCreated = telemetry.Output.Created;
        OutputSpoolRemoved = telemetry.Output.Removed;
        OutputSpoolPositionFullyObserved = telemetry.Output.PositionFullyObserved;
        OutputSpoolReadActualBytes = telemetry.Output.ReadActualBytes;
        OutputSpoolReadReservedBytes = telemetry.Output.ReadReservedBytes;
        OutputSpoolOriginalReadActualBytes = telemetry.Output.OriginalReadActualBytes;
        OutputSpoolOriginalReadReservedBytes = telemetry.Output.OriginalReadReservedBytes;
        OutputSpoolWriteActualBytes = telemetry.Output.WriteActualBytes;
        OutputSpoolWriteReservedBytes = telemetry.Output.WriteReservedBytes;
        OutputSpoolSeekActualCount = telemetry.Output.SeekActualCount;
        OutputSpoolSeekReservedCount = telemetry.Output.SeekReservedCount;
        OutputSpoolFlushActualCount = telemetry.Output.FlushActualCount;
        OutputSpoolFlushReservedCount = telemetry.Output.FlushReservedCount;
        OutputSpoolReservedHighWater = telemetry.Output.ReservedHighWater;
        OutputSpoolActualPositionHighWater = telemetry.Output.ActualPositionHighWater;
        OutputSpoolActualLengthHighWater = telemetry.Output.ActualLengthHighWater;
        OutputSpoolActiveActualLength = telemetry.Output.ActiveActualLength;
        OutputSpoolActiveReservedPosition = telemetry.Output.ActiveReservedPosition;
        StageCreated = telemetry.Stage.Created;
        StageRemoved = telemetry.Stage.Removed;
        StagePositionFullyObserved = telemetry.Stage.PositionFullyObserved;
        StageReadActualBytes = telemetry.Stage.ReadActualBytes;
        StageReadReservedBytes = telemetry.Stage.ReadReservedBytes;
        StageOriginalReadActualBytes = telemetry.Stage.OriginalReadActualBytes;
        StageOriginalReadReservedBytes = telemetry.Stage.OriginalReadReservedBytes;
        StageWriteActualBytes = telemetry.Stage.WriteActualBytes;
        StageWriteReservedBytes = telemetry.Stage.WriteReservedBytes;
        StageSeekActualCount = telemetry.Stage.SeekActualCount;
        StageSeekReservedCount = telemetry.Stage.SeekReservedCount;
        StageFlushActualCount = telemetry.Stage.FlushActualCount;
        StageFlushReservedCount = telemetry.Stage.FlushReservedCount;
        StageReservedHighWater = telemetry.Stage.ReservedHighWater;
        StageActualPositionHighWater = telemetry.Stage.ActualPositionHighWater;
        StageActualLengthHighWater = telemetry.Stage.ActualLengthHighWater;
        StageActiveActualLength = telemetry.Stage.ActiveActualLength;
        StageActiveReservedPosition = telemetry.Stage.ActiveReservedPosition;
    }
}

internal sealed class OfdOwnedIoTelemetry
{
    internal readonly object EventSync = new();
    private readonly OfdOwnedSaveStatistics[] snapshots;
    private readonly OfdOwnedSaveStatistics unavailable;
    private int next;
    internal readonly OfdOwnedIoObservation Input, Output, Stage;
    internal long LogicalU { get; private set; }
    internal bool NoopIssued { get; private set; }
    internal long ActiveActualLengths { get; private set; }
    internal long ActiveReservedPositions { get; private set; }
    internal long PeakActualLengths { get; private set; }
    internal long PeakReservedPositions { get; private set; }
    internal bool AllActualPositionsObserved => Input.PositionFullyObserved && Output.PositionFullyObserved && Stage.PositionFullyObserved;
    private OfdOwnedIoTelemetry()
    {
        snapshots = new OfdOwnedSaveStatistics[32];
        for (var i = 0; i < snapshots.Length; i++) snapshots[i] = OfdOwnedSaveStatistics.Preallocate();
        unavailable = OfdOwnedSaveStatistics.Preallocate();
        Input = OfdOwnedIoObservation.Preallocate(this); Output = OfdOwnedIoObservation.Preallocate(this); Stage = OfdOwnedIoObservation.Preallocate(this);
    }
    internal static OfdOwnedIoTelemetry CreateAfterPrepayment() => new();
    internal void RecordLogicalWriterReservation(long actual) { lock (EventSync) LogicalU = actual; }
    internal void RecordActualNoop() { lock (EventSync) NoopIssued = true; }
    internal void ObserveFileEvent()
    {
        lock (EventSync)
        {
            ActiveActualLengths = Input.ActiveActualLength + Output.ActiveActualLength + Stage.ActiveActualLength;
            ActiveReservedPositions = Input.ActiveReservedPosition + Output.ActiveReservedPosition + Stage.ActiveReservedPosition;
            PeakActualLengths = Math.Max(PeakActualLengths, ActiveActualLengths);
            PeakReservedPositions = Math.Max(PeakReservedPositions, ActiveReservedPositions);
        }
    }
    internal OfdOwnedSaveStatistics CaptureActual(int phase, long budget, OfdModelMaterialization material)
    {
        lock (EventSync)
        {
            if (next == snapshots.Length) return unavailable; // explicit unavailable, never stale counters or post-failure allocation
            var result = snapshots[next]; result.FillActual(++next, phase, budget, material, this); return result;
        }
    }
}

// Numeric event sink retained only by the actual private concrete file owner.
// It neither certifies completion nor exposes any file or model capability.
internal sealed class OfdOwnedIoObservation
{
    private readonly OfdOwnedIoTelemetry telemetry;
    private long position, length;
    private bool pendingRead, pendingWrite;
    private OfdOwnedIoObservation(OfdOwnedIoTelemetry telemetry) { this.telemetry = telemetry; }
    internal static OfdOwnedIoObservation Preallocate(OfdOwnedIoTelemetry telemetry) => new(telemetry);
    internal bool Created { get; private set; }
    internal bool Removed { get; private set; }
    internal bool PositionFullyObserved { get; private set; } = true;
    internal long ReadActualBytes { get; private set; }
    internal long ReadReservedBytes { get; private set; }
    internal long OriginalReadActualBytes { get; private set; }
    internal long OriginalReadReservedBytes { get; private set; }
    internal long WriteActualBytes { get; private set; }
    internal long WriteReservedBytes { get; private set; }
    internal long SeekActualCount { get; private set; }
    internal long SeekReservedCount { get; private set; }
    internal long FlushActualCount { get; private set; }
    internal long FlushReservedCount { get; private set; }
    internal long ReservedHighWater { get; private set; }
    internal long ActualPositionHighWater { get; private set; }
    internal long ActualLengthHighWater { get; private set; }
    internal long ActiveActualLength { get; private set; }
    internal long ActiveReservedPosition { get; private set; }
    private void Publish()
    {
        ActiveActualLength = Created && !Removed ? length : 0;
        ActiveReservedPosition = Created && !Removed ? ReservedHighWater : 0;
        telemetry.ObserveFileEvent();
    }
    internal void RecordOpened() { lock (telemetry.EventSync) { Created = true; Publish(); } }
    internal void ReservePosition(long target) { lock (telemetry.EventSync) { ReservedHighWater = Math.Max(ReservedHighWater, target); Publish(); } }
    internal void ReserveRead(int count) { lock (telemetry.EventSync) { ReadReservedBytes += count; pendingRead = true; } }
    internal void CompleteRead(int actual) { lock (telemetry.EventSync) { pendingRead = false; ReadActualBytes += actual; position += actual; ActualPositionHighWater = Math.Max(ActualPositionHighWater, position); } }
    internal void ReserveOriginalRead(int count) { lock (telemetry.EventSync) { OriginalReadReservedBytes += count; } }
    internal void CompleteOriginalRead(int actual) { lock (telemetry.EventSync) { OriginalReadActualBytes += actual; } }
    internal void ReserveWrite(int count) { lock (telemetry.EventSync) { WriteReservedBytes += count; } }
    internal void BeginActualWrite() { lock (telemetry.EventSync) { pendingWrite = true; } }
    internal void CompleteWrite(int actual) { lock (telemetry.EventSync) { pendingWrite = false; WriteActualBytes += actual; position += actual; length = Math.Max(length, position); ActualPositionHighWater = Math.Max(ActualPositionHighWater, position); ActualLengthHighWater = Math.Max(ActualLengthHighWater, length); Publish(); } }
    internal void ReserveSeek() { lock (telemetry.EventSync) { SeekReservedCount++; } }
    internal void CompleteSeek(long actual) { lock (telemetry.EventSync) { SeekActualCount++; position = actual; ActualPositionHighWater = Math.Max(ActualPositionHighWater, actual); } }
    internal void CompleteLength(long actual) { lock (telemetry.EventSync) { length = actual; ActualLengthHighWater = Math.Max(ActualLengthHighWater, actual); Publish(); } }
    internal void ReserveFlush() { lock (telemetry.EventSync) { FlushReservedCount++; } }
    internal void CompleteFlush() { lock (telemetry.EventSync) { FlushActualCount++; } }
    internal void RecordIoFailure() { lock (telemetry.EventSync) { if (pendingRead || pendingWrite) PositionFullyObserved = false; } }
    internal void RecordRemoved() { lock (telemetry.EventSync) { Removed = true; Publish(); } }
}
