using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Keywords.Internal;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Editing.Internal;

internal sealed class OfdPrivateGeometryHost
{
    internal OfdPrivateGeometryHost(OfdPrivateSaveBootstrap exactBootstrap, OfdOwnedGlyphGeometryHost actualIssuedHost)
        => exactBootstrap.RegisterHost(this, actualIssuedHost);
}

internal sealed class OfdPrivateSaveResult
{
    // All shells are constructed under bootstrap admission before model/provider
    // creation or dispatch. Success/failure only fill existing numeric records.
    private readonly OfdPrivateSaveStatistics[] phases = new OfdPrivateSaveStatistics[32];
    private readonly OfdPrivateSaveStatistics unavailable = new();
    private OfdPrivateSaveStatistics latest;
    private int count;
    internal OfdPrivateSaveResult()
    {
        for (var i = 0; i < phases.Length; i++) phases[i] = new OfdPrivateSaveStatistics();
        latest = unavailable;
    }
    internal void Prepare(bool changed, int requests, OfdTextReplacementStatistics preparation)
    {
        Changed = changed; Requests = requests; Preparation = preparation;
        CopiedOriginalFontBytes = preparation.Analyses.Count == 0 ? null : preparation.Analyses[0].CopiedFontBytes;
    }
    internal void Record(string stage, OfdOwnedSaveStatistics actual)
    {
        if (count == phases.Length) { latest = unavailable; return; }
        var copied = phases[count++]; copied.Fill(stage, actual); latest = copied;
    }
    internal void Complete() => Completed = true;
    internal bool Completed { get; private set; }
    internal bool? Changed { get; private set; }
    internal int? Requests { get; private set; }
    internal OfdTextReplacementStatistics? Preparation { get; private set; }
    internal long? CopiedOriginalFontBytes { get; private set; }
    internal int StatisticsCount => count;
    internal OfdPrivateSaveStatistics StatisticsAt(int index) => phases[index];
    internal OfdPrivateSaveStatistics Statistics => latest;
    internal long? StorageBytes => latest.Available ? latest.StorageBytes : null;
    internal long? CloneCopyBytes => latest.Available ? latest.CloneCopyBytes : null;
    internal long? ValidationWorkBytes => latest.Available ? latest.ValidationWorkBytes : null;
    internal long? OutputBytes => latest.Available ? latest.OutputUpperBoundBytes : null;
}

// Owns the fresh ticket before any model, provider or async worker exists.
internal sealed class OfdPrivateSaveBootstrap : IDisposable
{
    private const long FixedConfigurationAndDispatch = 16_384;
    private const long CopiedReportAndStatisticsShells = 65_536;
    private readonly object sync = new();
    private readonly OfdOwnedJobTicket ticket;
    private readonly CancellationToken preparationToken;
    private readonly int identities, selectedPages;
    private readonly OfdPrivateSaveResult report;
    private enum State { Fresh, Configured, PathsChecked, HostReserved, HostBound, Running, Completed, Failed, Disposed }
    private State state;
    private OfdPrivateGeometryHost? facade;
    private OfdOwnedGlyphGeometryHost? host;
    private readonly OfdOwnedNoopKernelIssuer kernel;
    private string? configuredInput, configuredOutput, configuredKeyword, configuredReplacement;
    private OfdPrivateSaveBootstrap(OfdOwnedJobTicket ticket, int identities, int pages, CancellationToken token)
    { this.ticket = ticket; this.identities = identities; selectedPages = pages; preparationToken = token;
      report = new OfdPrivateSaveResult();
      kernel = OfdTextReplacer.CreatePrivateKernel(ticket); }

    internal static OfdPrivateSaveBootstrap Create(long maxMaterializedBytes, int fontIdentityCount, int pageCount, CancellationToken token)
    {
        if (fontIdentityCount < 0 || fontIdentityCount > 64 || pageCount < 0 || pageCount > 1000)
            throw new ArgumentOutOfRangeException(nameof(fontIdentityCount));
        var ticket = OfdOwnedSaveAuthority.BeginFreshJob(new OfdModelLimits { MaxMaterializedBytes = maxMaterializedBytes }, token);
        try
        {
            // Fixed bootstrap/facade/dispatch state, delegates, StartNew/Unwrap,
            // worker/read/write/comparison/bridge Tasks and bounded final report.
            // Separate 64KiB covers result + 32 numeric once-fill snapshots and
            // the unavailable shell before any report allocation. Core prepays
            // its own statistics records independently in the fresh ticket.
            // Copies contain existing immutable references, not new SHA strings.
            OfdOwnedSaveAuthority.ReserveBootstrap(ticket,
                checked(FixedConfigurationAndDispatch + CopiedReportAndStatisticsShells + 32L * fontIdentityCount + 16L * pageCount), token);
            return new OfdPrivateSaveBootstrap(ticket, fontIdentityCount, pageCount, token);
        }
        catch { OfdOwnedSaveAuthority.Fail(ticket); throw; }
    }
    internal void Configure(string inputPath, string outputPath, string keyword, string replacement, CancellationToken token)
    {
        lock (sync)
        {
            try
            {
                Check(token); if (state != State.Fresh) throw new InvalidOperationException("PrivateConfigurationAlreadyBound");
                OfdOwnedSaveAuthority.AdmitBootstrapString(ticket, inputPath, token);
                OfdOwnedSaveAuthority.AdmitBootstrapString(ticket, outputPath, token);
                OfdOwnedSaveAuthority.AdmitBootstrapString(ticket, keyword, token);
                OfdOwnedSaveAuthority.AdmitBootstrapString(ticket, replacement, token);
                OfdOwnedSaveAuthority.ReserveBootstrapWork(ticket, 2L * (inputPath.Length + outputPath.Length), token);
                OfdOwnedSaveAuthority.BindFileIntent(ticket, inputPath, outputPath, token);
                configuredInput = inputPath; configuredOutput = outputPath;
                configuredKeyword = keyword; configuredReplacement = replacement; state = State.Configured;
                CaptureStatistics("Configured");
            }
            catch { Fail(); throw; }
        }
    }
    internal void AdmitConfiguredString(string value, CancellationToken token)
    {
        lock (sync) { try { Check(token); OfdOwnedSaveAuthority.AdmitBootstrapString(ticket, value, token); } catch { Fail(); throw; } }
    }
    internal void ReservePathDecomposition(int characters, CancellationToken token)
    {
        lock (sync)
        {
            try { Check(token); OfdOwnedSaveAuthority.ReserveBootstrap(ticket, checked(4096L + 32L * (characters + 1L)), token);
                OfdOwnedSaveAuthority.ReserveBootstrapWork(ticket, 2L * characters, token); }
            catch { Fail(); throw; }
        }
    }
    internal void ReservePathLookup(int characters, int previousLinks, CancellationToken token)
    {
        lock (sync)
        {
            try { Check(token); OfdOwnedSaveAuthority.ReserveBootstrap(ticket, checked(1024L + 4L * (characters + 32_768L)), token);
                OfdOwnedSaveAuthority.ReserveBootstrapWork(ticket, checked(2L * characters * (previousLinks + 2L)), token); }
            catch { Fail(); throw; }
        }
    }
    internal void RequireDistinctCanonicalPaths(string input, string output, CancellationToken token)
    {
        lock (sync)
        {
            try
            {
                Check(token); if (state != State.Configured) throw new InvalidOperationException("PrivatePathsAlreadyChecked");
                OfdOwnedSaveAuthority.ReserveBootstrapWork(ticket, 2L * Math.Max(input.Length, output.Length), token);
                if (string.Equals(input, output, PathComparison)) throw new ArgumentException("Replacement input and output resolve to the same path.");
                state = State.PathsChecked;
            }
            catch { Fail(); throw; }
        }
    }
    internal void ReserveHostConstruction(int exactIdentityCount, CancellationToken token)
    {
        lock (sync)
        {
            try
            {
                Check(token);
                if (state != State.PathsChecked || exactIdentityCount != identities) throw new InvalidOperationException("ForeignHostConfiguration");
                // Frozen identity references/dictionary, provider and its actual
                // Session wrappers/dictionaries; no face bytes are opened here.
                OfdOwnedSaveAuthority.ReserveBootstrap(ticket, checked(8192L + 1024L * exactIdentityCount), token);
                state = State.HostReserved;
            }
            catch { Fail(); throw; }
        }
    }
    internal void RegisterHost(OfdPrivateGeometryHost actualFacade, OfdOwnedGlyphGeometryHost actualHost)
    {
        lock (sync)
        {
            try
            {
                Check(default);
                if (state != State.HostReserved || facade is not null || host is not null || actualFacade is null || actualHost is null)
                    throw new InvalidOperationException("ForeignOrRepeatedHostRegistration");
                OfdOwnedSaveAuthority.BindHost(ticket, actualHost);
                facade = actualFacade; host = actualHost; state = State.HostBound;
                CaptureStatistics("HostBound");
            }
            catch { Fail(); throw; }
        }
    }
    internal Task<OfdPrivateSaveResult> ExecuteAsync(OfdPrivateGeometryHost exactFacade, string inputPath,
        string outputPath, string keyword, string replacement, int[]? pages, CancellationToken token)
    {
        lock (sync)
        {
            try
            {
                Check(token);
                if (state != State.HostBound || !ReferenceEquals(facade, exactFacade) || (pages?.Length ?? 0) != selectedPages)
                    throw new InvalidOperationException("ForeignOrReusedPrivateWorker");
                OfdOwnedSaveAuthority.ReserveBootstrapWork(ticket,
                    2L * (inputPath.Length + outputPath.Length + keyword.Length + replacement.Length), token);
                if (inputPath != configuredInput || outputPath != configuredOutput || keyword != configuredKeyword || replacement != configuredReplacement)
                    throw new InvalidOperationException("ForeignPrivateConfiguration");
                state = State.Running;
                return ExecuteCoreAsync(inputPath, outputPath, keyword, replacement, pages, token);
            }
            catch { Fail(); throw; }
        }
    }
    private async Task<OfdPrivateSaveResult> ExecuteCoreAsync(string inputPath, string outputPath,
        string keyword, string replacement, int[]? pages, CancellationToken token)
    {
        OfdOwnedInputRead? initial = null; OfdOwnedInputSnapshot? original = null;
        OfdTextReplacementOperation? operation = null; OfdOwnedWrittenSpool? written = null;
        OfdActualReadback? actualRead = null; OfdVerifiedReadback? verified = null;
        OfdOwnedGeometryExpectation? expectation = null; OfdPrivateAtomicBridge? atomic = null;
        Exception? primary = null;
        try
        {
            Check(token);
            initial = await OfdOwnedReader.ReadInitialAsync(ticket, inputPath, new OfdPackageLoadOptions(), token).ConfigureAwait(false);
            CaptureStatistics("InitialRead");
            OfdOwnedReader.ConsumeInitial(initial, out var lease, out var input); original = input;
            var analysis = OfdOwnedSaveAuthority.EnterAnalysis(lease, host!, out var source, out var context, out var guard);
            Check(token);
            operation = OfdTextReplacementOperation.ForPrivateSave(source, context, guard, new OfdTextReplacementOptions(), token);
            var compiled = await OfdTextReplacer.WritePrivateOwnedAsync(operation, analysis, host!, kernel, pages, keyword, replacement, token).ConfigureAwait(false);
            var changed = compiled.Written is not null;
            operation.Mark(changed ? "ActualOwnedWritten" : "ActualOwnedNoop");
            CaptureStatistics(changed ? "ActualOwnedWritten" : "ActualOwnedNoop");
            if (!changed)
            {
                var noop = compiled.Noop ?? throw new InvalidDataException("MissingActualNoopScope");
                atomic = OfdPrivateAtomicBridge.CreateNoop(noop, original, outputPath, token);
            }
            else
            {
                written = compiled.Written!; expectation = compiled.Expectation ?? throw new InvalidDataException("MissingActualCompilerExpectation");
                actualRead = await OfdOwnedReader.ReadOutputAsync(written, token).ConfigureAwait(false);
                operation.Mark("ActualOwnedReadback");
                CaptureStatistics("ActualOwnedReadback");
                verified = OfdOwnedPackageSave.VerifyActualReadback(written, actualRead, token);
                operation.Mark("ActualCompoundVerified");
                CaptureStatistics("ActualCompoundVerified");
                var comparison = OfdOwnedPackageSave.EnterGeometryComparison(verified, out var actual, out var addresses, token);
                var stageScope = expectation.CompareActual(comparison, actual, addresses, token);
                operation.Mark("ActualGeometryCompared");
                CaptureStatistics("ActualGeometryCompared");
                atomic = OfdPrivateAtomicBridge.Create(stageScope, written, original, outputPath, token);
            }
            // Report shells/arrays and any fallible local owner cleanup precede
            // the last input/source/token checks and the single real rename.
            report.Prepare(changed, compiled.Requests, operation.Snapshot());
            await atomic.CopyAsync(token).ConfigureAwait(false); // shared copy closes/deletes actual temporary files before proof issuance
            CaptureStatistics("ActualStageCopied");
            Exception? cleanup = null;
            DisposePrecommit(verified, ref cleanup); verified = null;
            DisposePrecommit(actualRead, ref cleanup); actualRead = null;
            DisposePrecommit(written, ref cleanup); written = null;
            DisposePrecommit(initial, ref cleanup); initial = null;
            DisposePrecommit(expectation, ref cleanup); expectation = null;
            DisposePrecommit(host, ref cleanup); host = null;
            if (cleanup is not null) System.Runtime.ExceptionServices.ExceptionDispatchInfo.Capture(cleanup).Throw();
            await atomic.ClaimAsync(token).ConfigureAwait(false); // owns original cleanup and final current path/S/token validation
            original = null; // consumed and closed by the actual shared claim frontier
            CaptureStatistics("ActualCommitClaimed");
            await atomic.CommitAsync(token).ConfigureAwait(false);
            CaptureStatistics("ActualCommitted"); // prepaid sealed numeric fields only; no allocation/token/I/O/model walk
            report.Complete();
            atomic = null; // stage was already physically closed; rename leaves no fallible file cleanup
            lock (sync) state = State.Completed;
            return report;
        }
        catch (Exception error)
        {
            // Revoke before Exception.Data, provider disposal and file cleanup.
            RecordFailure(error); primary = error;
            operation?.Fail(error); throw;
        }
        finally
        {
            if (primary is not null)
            {
                DisposeAfterFailure(atomic); DisposeAfterFailure(verified); DisposeAfterFailure(actualRead); DisposeAfterFailure(written);
                DisposeAfterFailure(original); DisposeAfterFailure(initial); DisposeAfterFailure(expectation);
            }
            // Only deterministic managed-state retirement remains after rename.
            // This operation never owns the shared candidate or physical files.
            if (primary is null) operation?.Dispose(); else DisposeAfterFailure(operation);
        }
    }
    private void Check(CancellationToken token)
    {
        preparationToken.ThrowIfCancellationRequested(); token.ThrowIfCancellationRequested();
        if (state is State.Failed or State.Disposed) throw new InvalidOperationException("PrivateSaveJobTerminal");
    }
    internal void RecordFailure(Exception primary)
    {
        Fail(); // actual terminal snapshot is copied before any callback or owner cleanup
        try { primary.Data["Ofdrw.TextReplacement.OwnedStatistics"] = report; } catch { }
    }
    private void CaptureStatistics(string stage)
        => report.Record(stage, OfdOwnedSaveAuthority.CaptureStatistics(ticket));
    private void Fail()
    {
        OfdOwnedSaveAuthority.Fail(ticket);
        lock (sync)
        {
            if (state is State.Failed or State.Disposed) return;
            state = State.Failed; CaptureStatistics("FailedBeforeCleanup");
        }
    }
    private void DisposePrecommit(IDisposable? value, ref Exception? first)
    {
        try { value?.Dispose(); }
        catch (Exception error) { Fail(); first ??= error; }
    }
    private static void DisposeAfterFailure(IDisposable? value) { try { value?.Dispose(); } catch { } }
    public void Dispose()
    {
        lock (sync)
        {
            if (state == State.Disposed) return;
            if (state != State.Completed) Fail(); state = State.Disposed;
        }
        host?.Dispose(); host = null; facade = null;
    }

    private static StringComparison PathComparison => System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.Windows) ||
        System.Runtime.InteropServices.RuntimeInformation.IsOSPlatform(System.Runtime.InteropServices.OSPlatform.OSX) ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;

}

internal static class OfdPrivateSaveWorker
{
    internal static Task<OfdPrivateSaveResult> ExecuteAsync(OfdPrivateSaveBootstrap bootstrap, OfdPrivateGeometryHost host,
        string inputPath, string outputPath, string keyword, string replacement, int[]? pages, CancellationToken token)
        => bootstrap.ExecuteAsync(host, inputPath, outputPath, keyword, replacement, pages, token);
}
