using System.Text;
using System.Text.Json;
using Ofdrw.Net.Layout.Editing.Internal;

// Artifact-only additive partial in the instrumented CLI copy. Every actual
// original method body and source file is unchanged. No reflection or new IVT.
internal static partial class Cli
{
    internal static async Task<(int ExitCode, string Json)> CaptureOwnedDispatchForDiagnosticsAsync(string[] replacementArgs, CancellationToken token)
    {
        try
        {
            var result = await DispatchPrivateOwnedSaveAsync(TextReplacementArguments.Parse(replacementArgs), token).ConfigureAwait(false);
            return (0, WriteOwnedDiagnosticJson(result, null));
        }
        catch (Exception primary)
        {
            try
            {
                // This report was already copied after terminalization, before
                // cleanup. Reading it supplies no owner or source capability.
                var result = primary.Data["Ofdrw.TextReplacement.OwnedStatistics"] as OfdPrivateSaveResult;
                return (1, WriteOwnedDiagnosticJson(result, primary.Message));
            }
            catch
            {
                System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(primary).Throw();
                throw;
            }
        }
    }
    private static string WriteOwnedDiagnosticJson(OfdPrivateSaveResult? report, string? error)
    {
        // Diagnostic serialization is outside the actual job and its counters.
        // It begins only after success/terminal failure and actual cleanup.
        using var output = new MemoryStream();
        using (var json = new Utf8JsonWriter(output, new JsonWriterOptions { Indented = true }))
        {
            json.WriteStartObject();
            json.WriteString("EvidenceKind", "instrumented typed dispatch diagnostic; not runtime acceptance");
            json.WriteBoolean("InstrumentedCli", true);
            json.WriteString("Source06", "439cf4404c63362ceeceb56e77078fcabb966763"); json.WriteString("Source09", "f7e3a62b6e63522a836643e59ff8144162ab5bfa");
            json.WriteString("UninstrumentedSourceBindingSha256", "26412a23b742eb7094f48386d0d51598023a5f92c55214f6452ac19edbd6a3e5");
            json.WriteBoolean("ReportPresent", report is not null);
            json.WriteString("Error", error);
            if (report is not null)
            {
                json.WriteBoolean("Completed", report.Completed);
                if (report.Changed is bool changed) json.WriteBoolean("Changed", changed); else json.WriteNull("Changed");
                if (report.Requests is int requests) json.WriteNumber("Requests", requests); else json.WriteNull("Requests");
                if (report.CopiedOriginalFontBytes is long copied) json.WriteNumber("CopiedOriginalFontBytes", copied); else json.WriteNull("CopiedOriginalFontBytes");
                json.WriteString("CopiedOriginalFontBytesMeaning", "P0 copied-byte total; not arbitrary single-font F");
                json.WriteNumber("SnapshotCapacity", 32); json.WriteNumber("StatisticsCount", report.StatisticsCount);
                json.WriteBoolean("SlotCapacityReached", report.StatisticsCount >= 32);
                json.WriteBoolean("LatestUnavailable", !report.Statistics.Available);
                json.WritePropertyName("LatestStatistics"); WriteOwnedNumericSnapshot(json, report.Statistics);
                json.WriteStartArray("PhaseSnapshots");
                for (var i = 0; i < report.StatisticsCount; i++) WriteOwnedNumericSnapshot(json, report.StatisticsAt(i));
                json.WriteEndArray();
                json.WriteStartArray("PreparationPhases");
                if (report.Preparation is not null)
                    foreach (var phase in report.Preparation.Phases)
                    {
                        json.WriteStartObject(); json.WriteString("Name", phase.Name);
                        json.WriteNumber("StorageBytes", phase.StorageBytes); json.WriteNumber("CopyBytes", phase.CopyBytes);
                        json.WriteNumber("OutputUpperBoundBytes", phase.OutputUpperBoundBytes);
                        json.WriteNumber("ValidationWorkBytes", phase.ValidationWorkBytes);
                        json.WriteNumber("TraversalReferences", phase.TraversalReferences); json.WriteNumber("TraversalNodes", phase.TraversalNodes);
                        json.WriteEndObject();
                    }
                json.WriteEndArray();
            }
            json.WriteEndObject(); json.Flush();
        }
        return Encoding.UTF8.GetString(output.ToArray());
    }
    private static void WriteOwnedNumericSnapshot(Utf8JsonWriter json, OfdPrivateSaveStatistics value)
    {
        json.WriteStartObject(); json.WriteString("Stage", value.Stage);
        json.WriteBoolean("Available", value.Available);
        // Available=false is explicit unknown, never zero/false budget success.
        if (value.Available) json.WriteNumber("Sequence", value.Sequence); else json.WriteNull("Sequence");
        if (value.Available) json.WriteNumber("PhaseCode", value.PhaseCode); else json.WriteNull("PhaseCode");
        if (value.Available) json.WriteNumber("BudgetBytes", value.BudgetBytes); else json.WriteNull("BudgetBytes");
        if (value.Available) json.WriteNumber("StorageBytes", value.StorageBytes); else json.WriteNull("StorageBytes");
        if (value.Available) json.WriteNumber("CloneCopyBytes", value.CloneCopyBytes); else json.WriteNull("CloneCopyBytes");
        if (value.Available) json.WriteNumber("ValidationWorkBytes", value.ValidationWorkBytes); else json.WriteNull("ValidationWorkBytes");
        if (value.Available) json.WriteNumber("OutputUpperBoundBytes", value.OutputUpperBoundBytes); else json.WriteNull("OutputUpperBoundBytes");
        if (value.Available) json.WriteNumber("NodeCount", value.NodeCount); else json.WriteNull("NodeCount");
        if (value.Available) json.WriteNumber("ReferenceCount", value.ReferenceCount); else json.WriteNull("ReferenceCount");
        if (value.Available) json.WriteNumber("SourceXmlBytes", value.SourceXmlBytes); else json.WriteNull("SourceXmlBytes");
        if (value.Available) json.WriteNumber("PathValueCount", value.PathValueCount); else json.WriteNull("PathValueCount");
        if (value.Available) json.WriteNumber("LogicalWriterReservationU", value.LogicalWriterReservationU); else json.WriteNull("LogicalWriterReservationU");
        if (value.Available) json.WriteBoolean("NoopIssued", value.NoopIssued); else json.WriteNull("NoopIssued");
        if (value.Available) json.WriteNumber("ActiveActualFileLengths", value.ActiveActualFileLengths); else json.WriteNull("ActiveActualFileLengths");
        if (value.Available) json.WriteNumber("ActiveReservedFilePositions", value.ActiveReservedFilePositions); else json.WriteNull("ActiveReservedFilePositions");
        if (value.Available) json.WriteNumber("PeakActualFileLengths", value.PeakActualFileLengths); else json.WriteNull("PeakActualFileLengths");
        if (value.Available) json.WriteNumber("PeakReservedFilePositions", value.PeakReservedFilePositions); else json.WriteNull("PeakReservedFilePositions");
        if (value.Available) json.WriteBoolean("AllActualPositionsObserved", value.AllActualPositionsObserved); else json.WriteNull("AllActualPositionsObserved");
        if (value.Available) json.WriteBoolean("InputSnapshotCreated", value.InputSnapshotCreated); else json.WriteNull("InputSnapshotCreated");
        if (value.Available) json.WriteBoolean("InputSnapshotRemoved", value.InputSnapshotRemoved); else json.WriteNull("InputSnapshotRemoved");
        if (value.Available) json.WriteBoolean("InputSnapshotPositionFullyObserved", value.InputSnapshotPositionFullyObserved); else json.WriteNull("InputSnapshotPositionFullyObserved");
        if (value.Available) json.WriteNumber("InputSnapshotReadActualBytes", value.InputSnapshotReadActualBytes); else json.WriteNull("InputSnapshotReadActualBytes");
        if (value.Available) json.WriteNumber("InputSnapshotReadReservedBytes", value.InputSnapshotReadReservedBytes); else json.WriteNull("InputSnapshotReadReservedBytes");
        if (value.Available) json.WriteNumber("InputSnapshotOriginalReadActualBytes", value.InputSnapshotOriginalReadActualBytes); else json.WriteNull("InputSnapshotOriginalReadActualBytes");
        if (value.Available) json.WriteNumber("InputSnapshotOriginalReadReservedBytes", value.InputSnapshotOriginalReadReservedBytes); else json.WriteNull("InputSnapshotOriginalReadReservedBytes");
        if (value.Available) json.WriteNumber("InputSnapshotWriteActualBytes", value.InputSnapshotWriteActualBytes); else json.WriteNull("InputSnapshotWriteActualBytes");
        if (value.Available) json.WriteNumber("InputSnapshotWriteReservedBytes", value.InputSnapshotWriteReservedBytes); else json.WriteNull("InputSnapshotWriteReservedBytes");
        if (value.Available) json.WriteNumber("InputSnapshotSeekActualCount", value.InputSnapshotSeekActualCount); else json.WriteNull("InputSnapshotSeekActualCount");
        if (value.Available) json.WriteNumber("InputSnapshotSeekReservedCount", value.InputSnapshotSeekReservedCount); else json.WriteNull("InputSnapshotSeekReservedCount");
        if (value.Available) json.WriteNumber("InputSnapshotFlushActualCount", value.InputSnapshotFlushActualCount); else json.WriteNull("InputSnapshotFlushActualCount");
        if (value.Available) json.WriteNumber("InputSnapshotFlushReservedCount", value.InputSnapshotFlushReservedCount); else json.WriteNull("InputSnapshotFlushReservedCount");
        if (value.Available) json.WriteNumber("InputSnapshotReservedHighWater", value.InputSnapshotReservedHighWater); else json.WriteNull("InputSnapshotReservedHighWater");
        if (value.Available) json.WriteNumber("InputSnapshotActualPositionHighWater", value.InputSnapshotActualPositionHighWater); else json.WriteNull("InputSnapshotActualPositionHighWater");
        if (value.Available) json.WriteNumber("InputSnapshotActualLengthHighWater", value.InputSnapshotActualLengthHighWater); else json.WriteNull("InputSnapshotActualLengthHighWater");
        if (value.Available) json.WriteNumber("InputSnapshotActiveActualLength", value.InputSnapshotActiveActualLength); else json.WriteNull("InputSnapshotActiveActualLength");
        if (value.Available) json.WriteNumber("InputSnapshotActiveReservedPosition", value.InputSnapshotActiveReservedPosition); else json.WriteNull("InputSnapshotActiveReservedPosition");
        if (value.Available) json.WriteBoolean("OutputSpoolCreated", value.OutputSpoolCreated); else json.WriteNull("OutputSpoolCreated");
        if (value.Available) json.WriteBoolean("OutputSpoolRemoved", value.OutputSpoolRemoved); else json.WriteNull("OutputSpoolRemoved");
        if (value.Available) json.WriteBoolean("OutputSpoolPositionFullyObserved", value.OutputSpoolPositionFullyObserved); else json.WriteNull("OutputSpoolPositionFullyObserved");
        if (value.Available) json.WriteNumber("OutputSpoolReadActualBytes", value.OutputSpoolReadActualBytes); else json.WriteNull("OutputSpoolReadActualBytes");
        if (value.Available) json.WriteNumber("OutputSpoolReadReservedBytes", value.OutputSpoolReadReservedBytes); else json.WriteNull("OutputSpoolReadReservedBytes");
        if (value.Available) json.WriteNumber("OutputSpoolOriginalReadActualBytes", value.OutputSpoolOriginalReadActualBytes); else json.WriteNull("OutputSpoolOriginalReadActualBytes");
        if (value.Available) json.WriteNumber("OutputSpoolOriginalReadReservedBytes", value.OutputSpoolOriginalReadReservedBytes); else json.WriteNull("OutputSpoolOriginalReadReservedBytes");
        if (value.Available) json.WriteNumber("OutputSpoolWriteActualBytes", value.OutputSpoolWriteActualBytes); else json.WriteNull("OutputSpoolWriteActualBytes");
        if (value.Available) json.WriteNumber("OutputSpoolWriteReservedBytes", value.OutputSpoolWriteReservedBytes); else json.WriteNull("OutputSpoolWriteReservedBytes");
        if (value.Available) json.WriteNumber("OutputSpoolSeekActualCount", value.OutputSpoolSeekActualCount); else json.WriteNull("OutputSpoolSeekActualCount");
        if (value.Available) json.WriteNumber("OutputSpoolSeekReservedCount", value.OutputSpoolSeekReservedCount); else json.WriteNull("OutputSpoolSeekReservedCount");
        if (value.Available) json.WriteNumber("OutputSpoolFlushActualCount", value.OutputSpoolFlushActualCount); else json.WriteNull("OutputSpoolFlushActualCount");
        if (value.Available) json.WriteNumber("OutputSpoolFlushReservedCount", value.OutputSpoolFlushReservedCount); else json.WriteNull("OutputSpoolFlushReservedCount");
        if (value.Available) json.WriteNumber("OutputSpoolReservedHighWater", value.OutputSpoolReservedHighWater); else json.WriteNull("OutputSpoolReservedHighWater");
        if (value.Available) json.WriteNumber("OutputSpoolActualPositionHighWater", value.OutputSpoolActualPositionHighWater); else json.WriteNull("OutputSpoolActualPositionHighWater");
        if (value.Available) json.WriteNumber("OutputSpoolActualLengthHighWater", value.OutputSpoolActualLengthHighWater); else json.WriteNull("OutputSpoolActualLengthHighWater");
        if (value.Available) json.WriteNumber("OutputSpoolActiveActualLength", value.OutputSpoolActiveActualLength); else json.WriteNull("OutputSpoolActiveActualLength");
        if (value.Available) json.WriteNumber("OutputSpoolActiveReservedPosition", value.OutputSpoolActiveReservedPosition); else json.WriteNull("OutputSpoolActiveReservedPosition");
        if (value.Available) json.WriteBoolean("StageCreated", value.StageCreated); else json.WriteNull("StageCreated");
        if (value.Available) json.WriteBoolean("StageRemoved", value.StageRemoved); else json.WriteNull("StageRemoved");
        if (value.Available) json.WriteBoolean("StagePositionFullyObserved", value.StagePositionFullyObserved); else json.WriteNull("StagePositionFullyObserved");
        if (value.Available) json.WriteNumber("StageReadActualBytes", value.StageReadActualBytes); else json.WriteNull("StageReadActualBytes");
        if (value.Available) json.WriteNumber("StageReadReservedBytes", value.StageReadReservedBytes); else json.WriteNull("StageReadReservedBytes");
        if (value.Available) json.WriteNumber("StageOriginalReadActualBytes", value.StageOriginalReadActualBytes); else json.WriteNull("StageOriginalReadActualBytes");
        if (value.Available) json.WriteNumber("StageOriginalReadReservedBytes", value.StageOriginalReadReservedBytes); else json.WriteNull("StageOriginalReadReservedBytes");
        if (value.Available) json.WriteNumber("StageWriteActualBytes", value.StageWriteActualBytes); else json.WriteNull("StageWriteActualBytes");
        if (value.Available) json.WriteNumber("StageWriteReservedBytes", value.StageWriteReservedBytes); else json.WriteNull("StageWriteReservedBytes");
        if (value.Available) json.WriteNumber("StageSeekActualCount", value.StageSeekActualCount); else json.WriteNull("StageSeekActualCount");
        if (value.Available) json.WriteNumber("StageSeekReservedCount", value.StageSeekReservedCount); else json.WriteNull("StageSeekReservedCount");
        if (value.Available) json.WriteNumber("StageFlushActualCount", value.StageFlushActualCount); else json.WriteNull("StageFlushActualCount");
        if (value.Available) json.WriteNumber("StageFlushReservedCount", value.StageFlushReservedCount); else json.WriteNull("StageFlushReservedCount");
        if (value.Available) json.WriteNumber("StageReservedHighWater", value.StageReservedHighWater); else json.WriteNull("StageReservedHighWater");
        if (value.Available) json.WriteNumber("StageActualPositionHighWater", value.StageActualPositionHighWater); else json.WriteNull("StageActualPositionHighWater");
        if (value.Available) json.WriteNumber("StageActualLengthHighWater", value.StageActualLengthHighWater); else json.WriteNull("StageActualLengthHighWater");
        if (value.Available) json.WriteNumber("StageActiveActualLength", value.StageActiveActualLength); else json.WriteNull("StageActiveActualLength");
        if (value.Available) json.WriteNumber("StageActiveReservedPosition", value.StageActiveReservedPosition); else json.WriteNull("StageActiveReservedPosition");
        json.WriteEndObject();
    }
}
