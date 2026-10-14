using Ofdrw.Net.Converter.Pdf.Geometry;
using Ofdrw.Net.Layout.Editing.Internal;

internal static partial class Cli
{
    private static async Task<int> ReplaceTextAsync(string[] args, CancellationToken token)
    {
        if (args.Length == 1 && IsHelp(args[0]))
        {
            PrintHelp();
            Console.WriteLine("Replacement requires explicit vetted original font SHA256:BYTECOUNT identities. Repeat --trusted-font for additional identities. Empty --replace deletes matching text. Pages are distinct 1-based positions; input and output must differ. Unsupported geometry or insufficient original space rejects the complete batch.");
            return 0;
        }
        var options = TextReplacementArguments.Parse(args);
        token.ThrowIfCancellationRequested();
        if (!File.Exists(options.InputPath)) throw new FileNotFoundException("Input file does not exist.", options.InputPath);
        return await SaveReplacementAsync(options, token).ConfigureAwait(false);
    }

    private static async Task<int> SaveReplacementAsync(TextReplacementArguments options, CancellationToken token)
    {
        _ = await DispatchPrivateOwnedSaveAsync(options, token).ConfigureAwait(false);
        Console.WriteLine($"Replaced {options.InputPath} -> {options.OutputPath}");
        return 0;
    }

    // Sole ticket-owned command route. Source-only integration candidate;
    // release acceptance requires the independent full runtime and visual gates.
    private static Task<OfdPrivateSaveResult> DispatchPrivateOwnedSaveAsync(TextReplacementArguments options, CancellationToken token)
    {
        var bootstrap = OfdPrivateSaveBootstrap.Create(512L * 1024 * 1024, options.TrustedFonts.Count, options.Pages?.Count ?? 0, token);
        try
        {
            bootstrap.Configure(options.InputPath, options.OutputPath, options.Keyword, options.Replacement, token);
            var input = ResolvePrivateOwnedPath(options.InputPath, bootstrap, token);
            var output = ResolvePrivateOwnedPath(options.OutputPath, bootstrap, token);
            bootstrap.RequireDistinctCanonicalPaths(input, output, token);
            for (var i = 0; i < options.TrustedFonts.Count; i++) bootstrap.AdmitConfiguredString(options.TrustedFonts[i].Sha256, token);
            var state = new PrivateSaveDispatch(bootstrap, options.InputPath, options.OutputPath, options.Keyword, options.Replacement,
                options.TrustedFonts.ToArray(), options.Pages?.ToArray(), token);
            Task<Task<OfdPrivateSaveResult>> queued;
            if (ExecutionContext.IsFlowSuppressed()) queued = QueuePrivateSave(state);
            else
            {
                // Synchronous queue only; restore the caller's flow before any
                // await or asynchronous completion/disposal continuation.
                using (ExecutionContext.SuppressFlow()) queued = QueuePrivateSave(state);
            }
            return FinishPrivateDispatchAsync(queued.Unwrap(), bootstrap);
        }
        catch (Exception error) { bootstrap.RecordFailure(error); try { bootstrap.Dispose(); } catch { } throw; }
    }
    private static Task<Task<OfdPrivateSaveResult>> QueuePrivateSave(PrivateSaveDispatch state)
        => Task.Factory.StartNew(static value => ((PrivateSaveDispatch)value!).RunAsync(), state,
            CancellationToken.None, TaskCreationOptions.DenyChildAttach, TaskScheduler.Default);
    private static async Task<OfdPrivateSaveResult> FinishPrivateDispatchAsync(Task<OfdPrivateSaveResult> work, OfdPrivateSaveBootstrap bootstrap)
    {
        Exception? primary = null;
        try { return await work.ConfigureAwait(false); }
        catch (Exception error) { bootstrap.RecordFailure(error); primary = error; throw; }
        finally { if (primary is null) bootstrap.Dispose(); else { try { bootstrap.Dispose(); } catch { } } }
    }
    private sealed class PrivateSaveDispatch(OfdPrivateSaveBootstrap bootstrap, string input, string output,
        string keyword, string replacement, OfdTrustedEmbeddedFontIdentity[] identities, int[]? pages, CancellationToken token)
    {
        internal Task<OfdPrivateSaveResult> RunAsync()
        {
            token.ThrowIfCancellationRequested();
            // First provider construction occurs on the fixed default worker;
            // initial models remain wholly inside the actual Layout/Reader job.
            var host = OfdPrivateGeometryFactory.Create(bootstrap, identities);
            return OfdPrivateSaveWorker.ExecuteAsync(bootstrap, host, input, output, keyword, replacement, pages, token);
        }
    }
    private static string ResolvePrivateOwnedPath(string path, OfdPrivateSaveBootstrap bootstrap, CancellationToken token)
        => ResolvePrivateOwnedPath(path, new HashSet<string>(OperatingSystem.IsWindows() || OperatingSystem.IsMacOS()
            ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal), bootstrap, token);
    private static string ResolvePrivateOwnedPath(string path, HashSet<string> links, OfdPrivateSaveBootstrap bootstrap, CancellationToken token)
    {
        bootstrap.ReservePathDecomposition(path.Length, token);
        var absolute = Path.GetFullPath(path); var root = Path.GetPathRoot(absolute)!;
        var components = absolute[root.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
        var current = root;
        for (var i = 0; i < components.Length; i++)
        {
            token.ThrowIfCancellationRequested(); bootstrap.ReservePathLookup(current.Length + components[i].Length, links.Count, token);
            current = Path.Combine(current, components[i]);
            FileSystemInfo entry = i == components.Length - 1 ? new FileInfo(current) : new DirectoryInfo(current);
            if (entry.LinkTarget is not null)
            {
                bootstrap.AdmitConfiguredString(current, token);
                if (!links.Add(current)) throw new IOException("Output path contains a link cycle.");
                var target = entry.ResolveLinkTarget(returnFinalTarget: true)?.FullName ?? throw new IOException("Output path link could not be resolved.");
                current = ResolvePrivateOwnedPath(target, links, bootstrap, token);
            }
        }
        return current;
    }

    // Existing full-guard component composition. No provider, Stream, model,
    // result, context, callback or credential is accepted from a caller.
    private sealed class OwnedTextReplacementJob : IDisposable
    {
        private readonly TextReplacementArguments options;
        private readonly OfdTrustedEmbeddedFontGeometryProvider provider;
        private readonly int[]? pages;
        private readonly OfdTextReplacementFileSession session;
        private OwnedTextReplacementJob(TextReplacementArguments options, OfdTextReplacementFileSession session)
        {
            this.options = options; this.session = session;
            // Parse owns immutable identity values and readonly page positions.
            // Copy the containers so the job retains no external selection.
            pages = options.Pages?.ToArray();
            provider = new OfdTrustedEmbeddedFontGeometryProvider(options.TrustedFonts.ToArray());
        }
        internal static async Task<OwnedTextReplacementJob> CreateAsync(TextReplacementArguments options, CancellationToken token)
        {
            var session = await OfdTextReplacementFileSession.OpenForPrivateCliJobAsync(options.InputPath, token).ConfigureAwait(false);
            try
            {
                session.ReserveOwnedConfiguration(options.TrustedFonts.Count, options.Pages?.Count ?? 0, token);
                return new OwnedTextReplacementJob(options, session);
            }
            catch { session.FailOwnedPreparation(); session.Dispose(); throw; }
        }
        internal async Task SaveAsync(CancellationToken token)
        {
            try
            {
                var report = await session.PrepareAndWriteOwnedAsync(provider, options.Keyword, options.Replacement, pages, token).ConfigureAwait(false);
                if (report.Changed)
                {
                    await session.ReadOwnedOutputAsync(token).ConfigureAwait(false);
                    session.VerifyOwnedGeometryInputs(token);
                    // Actual same-context Reader and prepared-input equivalence
                    // remain insufficient without the separately owned receipt.
                    // Until receipt/preservation/final-copy gates are wired, fail before any
                    // AtomicOutput stage is created for a changed document.
                    throw new NotSupportedException("Changed text output requires an approved Writer receipt, preservation and final copy verification.");
                }
                session.ReserveOwnedAtomicMetadata(options.OutputPath, token);
                var input = ResolveOwnedPath(options.InputPath, token);
                var output = ResolveOwnedPath(options.OutputPath, token);
                var comparison = OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
                if (string.Equals(input, output, comparison)) throw new ArgumentException("Replacement input and output resolve to the same path.");
                AtomicOutput? atomic = null; Exception? primary = null;
                try
                {
                    atomic = new AtomicOutput(options.OutputPath);
                    await session.CopyOwnedNoOpToStageAsync(atomic.Stream, token).ConfigureAwait(false);
                    await session.ClaimOwnedCommitAsync(token).ConfigureAwait(false);
                    await atomic.CommitAsync(token).ConfigureAwait(false);
                }
                catch (Exception error) { primary = error; session.FailOwnedPreparation(); throw; }
                finally
                {
                    if (atomic is not null)
                    {
                        if (primary is null) await atomic.DisposeAsync().ConfigureAwait(false);
                        else { try { await atomic.DisposeAsync().ConfigureAwait(false); } catch { } }
                    }
                }
            }
            catch { session.FailOwnedPreparation(); throw; }
        }
        private string ResolveOwnedPath(string path, CancellationToken token)
            => ResolveOwnedPath(path, new HashSet<string>(OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() ? StringComparer.OrdinalIgnoreCase : StringComparer.Ordinal), token);
        private string ResolveOwnedPath(string path, HashSet<string> links, CancellationToken token)
        {
            session.ReserveOwnedPathDecomposition(path.Length, token);
            var absolute = Path.GetFullPath(path); var root = Path.GetPathRoot(absolute)!;
            var components = absolute[root.Length..].Split(Path.DirectorySeparatorChar, StringSplitOptions.RemoveEmptyEntries);
            var current = root;
            for (var i = 0; i < components.Length; i++)
            {
                token.ThrowIfCancellationRequested(); session.ReserveOwnedPathLookup(current.Length + components[i].Length, token);
                current = Path.Combine(current, components[i]);
                FileSystemInfo entry = i == components.Length - 1 ? new FileInfo(current) : new DirectoryInfo(current);
                if (entry.LinkTarget is not null)
                {
                    if (!links.Add(current)) throw new IOException("Output path contains a link cycle.");
                    var target = entry.ResolveLinkTarget(returnFinalTarget: true)?.FullName ?? throw new IOException("Output path link could not be resolved.");
                    current = ResolveOwnedPath(target, links, token);
                }
            }
            return current;
        }
        public void Dispose() => session.Dispose();
    }
}
