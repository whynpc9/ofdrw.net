using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Layout.Editing.Internal;

// A prepaid once-fill copy of actual Core numeric observations. It contains no
// ticket/context/model/stream/path/claim/hash handle and grants no save authority.
internal sealed class OfdPrivateSaveStatistics
{
    private bool filled;
    internal string? Stage { get; private set; }
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
    internal void Fill(string stage, OfdOwnedSaveStatistics actual)
    {
        if (filled) return; // no post-rename diagnostic exception
        filled = true; Stage = stage;
        Available = actual.Available;
        Sequence = actual.Sequence;
        PhaseCode = actual.PhaseCode;
        BudgetBytes = actual.BudgetBytes;
        StorageBytes = actual.StorageBytes;
        CloneCopyBytes = actual.CloneCopyBytes;
        ValidationWorkBytes = actual.ValidationWorkBytes;
        OutputUpperBoundBytes = actual.OutputUpperBoundBytes;
        NodeCount = actual.NodeCount;
        ReferenceCount = actual.ReferenceCount;
        SourceXmlBytes = actual.SourceXmlBytes;
        PathValueCount = actual.PathValueCount;
        LogicalWriterReservationU = actual.LogicalWriterReservationU;
        NoopIssued = actual.NoopIssued;
        ActiveActualFileLengths = actual.ActiveActualFileLengths;
        ActiveReservedFilePositions = actual.ActiveReservedFilePositions;
        PeakActualFileLengths = actual.PeakActualFileLengths;
        PeakReservedFilePositions = actual.PeakReservedFilePositions;
        AllActualPositionsObserved = actual.AllActualPositionsObserved;
        InputSnapshotCreated = actual.InputSnapshotCreated;
        InputSnapshotRemoved = actual.InputSnapshotRemoved;
        InputSnapshotPositionFullyObserved = actual.InputSnapshotPositionFullyObserved;
        InputSnapshotReadActualBytes = actual.InputSnapshotReadActualBytes;
        InputSnapshotReadReservedBytes = actual.InputSnapshotReadReservedBytes;
        InputSnapshotOriginalReadActualBytes = actual.InputSnapshotOriginalReadActualBytes;
        InputSnapshotOriginalReadReservedBytes = actual.InputSnapshotOriginalReadReservedBytes;
        InputSnapshotWriteActualBytes = actual.InputSnapshotWriteActualBytes;
        InputSnapshotWriteReservedBytes = actual.InputSnapshotWriteReservedBytes;
        InputSnapshotSeekActualCount = actual.InputSnapshotSeekActualCount;
        InputSnapshotSeekReservedCount = actual.InputSnapshotSeekReservedCount;
        InputSnapshotFlushActualCount = actual.InputSnapshotFlushActualCount;
        InputSnapshotFlushReservedCount = actual.InputSnapshotFlushReservedCount;
        InputSnapshotReservedHighWater = actual.InputSnapshotReservedHighWater;
        InputSnapshotActualPositionHighWater = actual.InputSnapshotActualPositionHighWater;
        InputSnapshotActualLengthHighWater = actual.InputSnapshotActualLengthHighWater;
        InputSnapshotActiveActualLength = actual.InputSnapshotActiveActualLength;
        InputSnapshotActiveReservedPosition = actual.InputSnapshotActiveReservedPosition;
        OutputSpoolCreated = actual.OutputSpoolCreated;
        OutputSpoolRemoved = actual.OutputSpoolRemoved;
        OutputSpoolPositionFullyObserved = actual.OutputSpoolPositionFullyObserved;
        OutputSpoolReadActualBytes = actual.OutputSpoolReadActualBytes;
        OutputSpoolReadReservedBytes = actual.OutputSpoolReadReservedBytes;
        OutputSpoolOriginalReadActualBytes = actual.OutputSpoolOriginalReadActualBytes;
        OutputSpoolOriginalReadReservedBytes = actual.OutputSpoolOriginalReadReservedBytes;
        OutputSpoolWriteActualBytes = actual.OutputSpoolWriteActualBytes;
        OutputSpoolWriteReservedBytes = actual.OutputSpoolWriteReservedBytes;
        OutputSpoolSeekActualCount = actual.OutputSpoolSeekActualCount;
        OutputSpoolSeekReservedCount = actual.OutputSpoolSeekReservedCount;
        OutputSpoolFlushActualCount = actual.OutputSpoolFlushActualCount;
        OutputSpoolFlushReservedCount = actual.OutputSpoolFlushReservedCount;
        OutputSpoolReservedHighWater = actual.OutputSpoolReservedHighWater;
        OutputSpoolActualPositionHighWater = actual.OutputSpoolActualPositionHighWater;
        OutputSpoolActualLengthHighWater = actual.OutputSpoolActualLengthHighWater;
        OutputSpoolActiveActualLength = actual.OutputSpoolActiveActualLength;
        OutputSpoolActiveReservedPosition = actual.OutputSpoolActiveReservedPosition;
        StageCreated = actual.StageCreated;
        StageRemoved = actual.StageRemoved;
        StagePositionFullyObserved = actual.StagePositionFullyObserved;
        StageReadActualBytes = actual.StageReadActualBytes;
        StageReadReservedBytes = actual.StageReadReservedBytes;
        StageOriginalReadActualBytes = actual.StageOriginalReadActualBytes;
        StageOriginalReadReservedBytes = actual.StageOriginalReadReservedBytes;
        StageWriteActualBytes = actual.StageWriteActualBytes;
        StageWriteReservedBytes = actual.StageWriteReservedBytes;
        StageSeekActualCount = actual.StageSeekActualCount;
        StageSeekReservedCount = actual.StageSeekReservedCount;
        StageFlushActualCount = actual.StageFlushActualCount;
        StageFlushReservedCount = actual.StageFlushReservedCount;
        StageReservedHighWater = actual.StageReservedHighWater;
        StageActualPositionHighWater = actual.StageActualPositionHighWater;
        StageActualLengthHighWater = actual.StageActualLengthHighWater;
        StageActiveActualLength = actual.StageActiveActualLength;
        StageActiveReservedPosition = actual.StageActiveReservedPosition;
    }
}
