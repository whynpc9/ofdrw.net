using System.Globalization;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Layout.Editing.Internal;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Layout.Editing;

/// <summary>Replaces genuine keyword hits with conservative original-font geometry and fixed source slots.</summary>
/// <remarks>Supports literal simple scalars, horizontal positive progression and translation only.
/// Unsupported sources, overlapping requests or insufficient original space reject the entire batch.
/// Keep input models stable through result consumption; the service never mutates the input.</remarks>
public sealed class OfdTextReplacer
{
    private readonly IOfdGlyphGeometryProvider provider;
    private readonly OfdKeywordLocatorLimits locatorLimits;
    public OfdTextReplacer(IOfdGlyphGeometryProvider geometryProvider, OfdKeywordLocatorLimits? locatorLimits = null)
    { provider = geometryProvider ?? throw new ArgumentNullException(nameof(geometryProvider)); this.locatorLimits = locatorLimits ?? new(); }

    public OfdTextReplacementResult Replace(OfdDocumentPackage source, IReadOnlyList<OfdTextReplacement> replacements,
        OfdTextReplacementOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        var policies = options ?? new OfdTextReplacementOptions();
        var operation = new OfdTextReplacementOperation(source, policies, cancellationToken);
        try { return Prepare(operation, Freeze(replacements, policies.Limits, operation.Context.Material, cancellationToken), cancellationToken); }
        catch (Exception error) { operation.Fail(error); operation.Dispose(); throw; }
    }
    public OfdTextReplacementResult Replace(OfdDocumentPackage source, string keyword, string replacementText,
        OfdTextReplacementOptions? options = null, CancellationToken cancellationToken = default)
        => ReplaceKeyword(source, null, keyword, replacementText, options, cancellationToken);
    public OfdTextReplacementResult Replace(OfdDocumentPackage source, IReadOnlyList<int> pageIndexes,
        string keyword, string replacementText, OfdTextReplacementOptions? options = null, CancellationToken cancellationToken = default)
    {
        if (pageIndexes is null) throw new ArgumentNullException(nameof(pageIndexes));
        return ReplaceKeyword(source, pageIndexes, keyword, replacementText, options, cancellationToken);
    }
    private OfdTextReplacementResult ReplaceKeyword(OfdDocumentPackage source, IReadOnlyList<int>? pageIndexes,
        string keyword, string replacementText, OfdTextReplacementOptions? options, CancellationToken cancellationToken)
    {
        if (source is null) throw new ArgumentNullException(nameof(source));
        if (keyword is null) throw new ArgumentNullException(nameof(keyword));
        if (keyword.Length == 0) throw new ArgumentException("Keyword cannot be empty.", nameof(keyword));
        if (replacementText is null) throw new ArgumentNullException(nameof(replacementText));
        var policies = options ?? new OfdTextReplacementOptions();
        if (replacementText.Length > policies.Limits.MaxReplacementUtf16Units) throw new InvalidDataException("ReplacementTextLimit");
        var operation = new OfdTextReplacementOperation(source, policies, cancellationToken);
        return PrepareKeywordOperation(operation, pageIndexes, keyword, replacementText, cancellationToken)!;
    }

    internal OfdTextReplacementResult ReplaceControlled(OfdTextReplacementFileSession session, IReadOnlyList<int>? pageIndexes,
        string keyword, string replacementText, CancellationToken cancellationToken)
    {
        if (keyword is null) throw new ArgumentNullException(nameof(keyword));
        if (keyword.Length == 0) throw new ArgumentException("Keyword cannot be empty.", nameof(keyword));
        if (replacementText is null) throw new ArgumentNullException(nameof(replacementText));
        if (replacementText.Length > session.Options.Limits.MaxReplacementUtf16Units) throw new InvalidDataException("ReplacementTextLimit");
        var operation = OfdTextReplacementOperation.ForFileSession(session, session.BoundSource, session.BoundContext, cancellationToken);
        return PrepareKeywordOperation(operation, pageIndexes, keyword, replacementText, cancellationToken)!;
    }

    internal async Task PrepareAndWriteKeywordForOwnedAsync(OfdTextReplacementFileSession session, OfdDocumentPackage source,
        OfdControlledEditContext context, int[]? pageIndexes, string keyword, string replacementText, CancellationToken token)
    {
        if (keyword is null) throw new ArgumentNullException(nameof(keyword));
        if (keyword.Length == 0) throw new ArgumentException("Keyword cannot be empty.", nameof(keyword));
        if (replacementText is null) throw new ArgumentNullException(nameof(replacementText));
        if (replacementText.Length > session.Options.Limits.MaxReplacementUtf16Units) throw new InvalidDataException("ReplacementTextLimit");
        var operation = OfdTextReplacementOperation.ForOwnedFileSession(session, source, context, token);
        try
        {
            var analysis = AnalyzeKeywordOperation(operation, pageIndexes, keyword, replacementText, token, session);
            await FinishOwnedWriteAsync(session, operation, analysis.Requests, analysis.Roots, token).ConfigureAwait(false);
        }
        catch (Exception error)
        {
            // Terminal state precedes diagnostics and all owned cleanup.
            session.FailOwnedPreparation(); operation.Fail(error);
            try { operation.Dispose(); } catch { }
            throw;
        }
    }

    private OfdTextReplacementResult PrepareKeywordOperation(OfdTextReplacementOperation operation, IReadOnlyList<int>? pageIndexes,
        string keyword, string replacementText, CancellationToken cancellationToken)
    {
        try
        {
            var analysis = AnalyzeKeywordOperation(operation, pageIndexes, keyword, replacementText, cancellationToken);
            operation.Validate(cancellationToken); operation.Mark("MeasuredAndClosedProvider");
            return Finish(operation, analysis.Requests, analysis.Roots, cancellationToken);
        }
        catch (Exception error) { operation.Fail(error); operation.Dispose(); throw; }
    }

    // The sole keyword analysis/compiler kernel serves both public replacement
    // and the private write-through route. It never returns a Writer or candidate.
    private (OfdTextReplacement[] Requests, List<OfdTextReplacementRoot>? Roots) AnalyzeKeywordOperation(
        OfdTextReplacementOperation operation, IReadOnlyList<int>? pageIndexes, string keyword, string replacementText,
        CancellationToken cancellationToken, OfdTextReplacementFileSession? ownedFile = null)
    {
        var source = operation.Source!; var policies = operation.Options;
        operation.Begin("ValidateReplacement");
        operation.Context.Material.ChargeWork(2L * replacementText.Length, cancellationToken);
        ValidateLiteral(replacementText, cancellationToken);
        operation.Begin("KeywordQuery");
        if (pageIndexes is null)
        {
            if (source.Pages.Count > locatorLimits.MaxPages) throw new InvalidDataException("ReplacementPageLimit");
            operation.Context.Material.ChargeCopy(128L + 8L * source.Pages.Count, cancellationToken);
            pageIndexes = Enumerable.Range(0, source.Pages.Count).ToArray();
        }
        OfdTextReplacement[]? requests = null; List<OfdTextReplacementRoot>? roots = null;
        Action<OfdKeywordAnalysisCapture> captureMatches = capture =>
            {
                operation.Begin("AuthenticateKeywordMatches");
                var count = capture.CurrentHits.Count;
                if (count > policies.Limits.MaxRequests) throw new InvalidDataException("ReplacementRequestLimit");
                if (checked((long)count * replacementText.Length) > policies.Limits.MaxReplacementUtf16Units)
                    throw new InvalidDataException("ReplacementTextLimit");
                operation.Context.Material.ChargeCopy(256L + 256L * count, cancellationToken);
                requests = new OfdTextReplacement[count];
                for (var i = 0; i < count; i++)
                {
                    cancellationToken.ThrowIfCancellationRequested(); var hit = capture.CurrentHits[i];
                    capture.AuthenticateCurrentHit(hit); requests[i] = new OfdTextReplacement(hit, replacementText);
                }
                RejectOverlaps(requests, operation.Context.Material, cancellationToken);
                operation.Validate(cancellationToken); operation.Mark("KeywordMatchesCaptured");
                if (requests.Any(value => value.ReplacementText != value.Hit.MatchedText))
                {
                    var pages = AffectedPages(requests, operation, cancellationToken);
                    roots = OfdTextReplacementFitAdapter.Plan(capture, requests, operation, cancellationToken, pages);
                    // File creation stays inside the private capture phase,
                    // before final provider Dispose and its source/font checks.
                    ownedFile?.CreateOwnedOutputSpool(operation, cancellationToken);
                }
            };
        var found = ownedFile is null
            ? OfdKeywordEditSource.Analyze(source, pageIndexes, keyword, provider, locatorLimits, policies.SearchOptions,
                operation.Context.Material, () => operation.Validate(cancellationToken), captureMatches, cancellationToken, operation.BeforeCallback)
            : OfdKeywordEditSource.AnalyzeForOwnedFile(source, (int[])pageIndexes, keyword, provider, locatorLimits, policies.SearchOptions,
                operation.Context.Material, () => operation.Validate(cancellationToken), captureMatches, cancellationToken, operation.BeforeCallback);
        operation.RecordAnalysis(found);
        // No callback or await is inserted after the final actual provider
        // Dispose. The caller proceeds directly to its private batch/factory.
        if (ownedFile is not null) operation.Mark("MeasuredAndClosedProvider");
        return (requests!, roots);
    }

    private OfdTextReplacementResult Prepare(OfdTextReplacementOperation operation, OfdTextReplacement[] requests, CancellationToken token)
    {
        var source = operation.Source!; operation.Begin("AuthenticateHits");
        var material = operation.Context.Material;
        material.ChargeCopy(256L + 256L * requests.Length, token);
        var hits = requests.Select(value => value.Hit).ToArray();
        var digests = OfdKeywordEditSource.ValidateHits(source, hits, locatorLimits, operation.Options.SearchOptions, material, token);
        operation.RecordRetainedDigests(digests.Bytes, digests.Operations);
        RejectOverlaps(requests, material, token); operation.Validate(token); operation.Mark("AuthenticatedHits");
        var changed = requests.Any(value => value.ReplacementText != value.Hit.MatchedText);
        if (changed)
        {
            var affectedPages = AffectedPages(requests, operation, token);
            material.ChargeCopy(128L + 8L * affectedPages.Count, token);
            var pages = affectedPages.OrderBy(value => value).ToArray();
            operation.Begin("MeasureReplacement"); List<OfdTextReplacementRoot>? roots = null;
            var analyzed = OfdKeywordEditSource.Analyze(source, pages, hits[0].MatchedText, provider, locatorLimits, operation.Options.SearchOptions,
                material, () => operation.Validate(token), capture => roots = OfdTextReplacementFitAdapter.Plan(capture, requests, operation, token, affectedPages), token, operation.BeforeCallback);
            operation.RecordAnalysis(analyzed);
            operation.Validate(token); operation.Mark("MeasuredAndClosedProvider");
            return Finish(operation, requests, roots, token);
        }
        return Finish(operation, requests, null, token);
    }
    private static HashSet<int> AffectedPages(OfdTextReplacement[] requests, OfdTextReplacementOperation operation, CancellationToken token)
    {
        var material = operation.Context.Material;
        material.ChargeCopy(256L + 128L * requests.Length, token);
        var pages = new HashSet<int>();
        foreach (var request in requests)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(32, token);
            var page = request.Hit.PageIndex;
            if (page < 0 || page >= operation.Source!.Pages.Count) throw new InvalidDataException("ReplacementPageMismatch");
            pages.Add(page);
        }
        foreach (var page in pages)
        {
            token.ThrowIfCancellationRequested();
            if (operation.Source!.Pages[page].Templates.Count != 0 || operation.Source.Pages[page].AnnotationAppearances.Count != 0)
                throw new NotSupportedException("UnprovedReplacementOverlayNeighbors");
        }
        return pages;
    }
    private static async Task FinishOwnedWriteAsync(OfdTextReplacementFileSession session, OfdTextReplacementOperation operation,
        OfdTextReplacement[] requests, List<OfdTextReplacementRoot>? roots, CancellationToken token)
    {
        if (roots is null)
        {
            // No-op has no clone/Writer entry to discharge final freshness.
            operation.Validate(token); operation.Mark("PreparedOwnedNoOp");
            operation.Context.Material.ChargeCopy(512L * requests.Length + 256, token);
            session.AcceptOwnedNoOp(operation, requests.Select(value => new OfdTextReplacementObservation(value)).ToArray());
            return;
        }
        operation.Begin("ContinuousOwnedWriteFactory");
        var changedCount = 0;
        foreach (var root in roots)
        {
            token.ThrowIfCancellationRequested();
            if (!root.Plan.Requests.All(value => value.Disposition == OfdTextFitDisposition.Unchanged)) changedCount++;
        }
        operation.Context.Material.ChargeCopy(128L + 32L * changedCount, token);
        var edits = new OfdRetainedTextEdit[changedCount]; var index = 0;
        foreach (var root in roots)
        {
            token.ThrowIfCancellationRequested();
            if (!root.Plan.Requests.All(value => value.Disposition == OfdTextFitDisposition.Unchanged))
                edits[index++] = new OfdRetainedTextEdit(root.Owner, root.Slots);
        }
        using var batch = OfdRetainedTextEditBatch.Capture(operation.Source!, operation.SourceGuard!, operation.Context, edits, token);
        // The real shared factory starts on this stack. No prepared capability,
        // callback, reporting return or await occurs before its closed prefix.
        await session.WriteContinuousOwnedAsync(operation, batch, requests, token).ConfigureAwait(false);
    }

    private static OfdTextReplacementResult Finish(OfdTextReplacementOperation operation, OfdTextReplacement[] requests,
        List<OfdTextReplacementRoot>? roots, CancellationToken token)
    {
        var material = operation.Context.Material;
        if (roots is not null)
        {
            operation.Begin("CloneAndPatch");
            operation.CloneCandidate(token);
            var candidate = operation.Candidate!;
            foreach (var root in roots!)
            {
                token.ThrowIfCancellationRequested();
                // Identity requests participate in conflict/fit checks but must
                // not normalize their source XML as an incidental edit.
                if (root.Plan.Requests.All(value => value.Disposition == OfdTextFitDisposition.Unchanged)) continue;
                candidate.ApplyTextSlots(candidate.GetMapped(root.Owner), root.Slots, token);
            }
            operation.Validate(token); operation.CaptureCandidate(token);
            operation.Mark("PreparedCandidate");
        }
        else operation.Mark("PreparedNoOp");
        material.ChargeCopy(512L * requests.Length + 256, token);
        return new OfdTextReplacementResult(operation, requests.Select(value => new OfdTextReplacementObservation(value)).ToArray());
    }
    private static OfdTextReplacement[] Freeze(IReadOnlyList<OfdTextReplacement> requests, OfdTextReplacementLimits limits,
        OfdModelMaterialization material, CancellationToken token)
    {
        if (requests is null) throw new ArgumentNullException(nameof(requests));
        token.ThrowIfCancellationRequested(); var count = requests.Count;
        if (count < 0 || count > limits.MaxRequests) throw new InvalidDataException("ReplacementRequestLimit");
        material.ChargeCopy(256L + 16L * count, token);
        var copy = new OfdTextReplacement[count]; long units = 0;
        for (var i = 0; i < count; i++)
        {
            token.ThrowIfCancellationRequested(); var request = requests[i] ?? throw new ArgumentException("Null replacement request.", nameof(requests));
            units = checked(units + request.ReplacementText.Length);
            if (units > limits.MaxReplacementUtf16Units) throw new InvalidDataException("ReplacementTextLimit");
            material.ChargeWork(2L * request.ReplacementText.Length, token);
            ValidateLiteral(request.ReplacementText, token); copy[i] = request;
        }
        token.ThrowIfCancellationRequested(); return copy;
    }
    internal static int ValidateLiteral(string text, CancellationToken token)
    {
        var nextCheck = 0; var checkpoints = 0;
        for (var i = 0; i < text.Length; i++)
        {
            if (i >= nextCheck) { token.ThrowIfCancellationRequested(); nextCheck = i + 256; checkpoints++; }
            var start = i; int scalar = text[i];
            if (char.IsHighSurrogate(text[i]))
            {
                if (++i >= text.Length || !char.IsLowSurrogate(text[i])) throw new ArgumentException("Unpaired replacement surrogate.");
                scalar = char.ConvertToUtf32(text[start], text[i]);
            }
            else if (char.IsLowSurrogate(text[i])) throw new ArgumentException("Unpaired replacement surrogate.");
            var category = CharUnicodeInfo.GetUnicodeCategory(text, start);
            if (!Ordinary(scalar) || char.IsWhiteSpace(text, start) || scalar is '\\' or 0xa4 ||
                category is UnicodeCategory.NonSpacingMark or UnicodeCategory.SpacingCombiningMark or UnicodeCategory.EnclosingMark or UnicodeCategory.Format or UnicodeCategory.Control)
                throw new NotSupportedException("UnsafeReplacementScalarProfile");
        }
        token.ThrowIfCancellationRequested(); return checkpoints + 1;
    }
    private static bool Ordinary(int scalar) => scalar is >= 0x20 and <= 0x24f or >= 0x2000 and <= 0x206f or
        >= 0x3000 and <= 0x303f or >= 0x3400 and <= 0x4dbf or >= 0x4e00 and <= 0x9fff or
        >= 0xf900 and <= 0xfaff or >= 0xff00 and <= 0xffef or >= 0x20000 and <= 0x2ffff;
    private static void RejectOverlaps(OfdTextReplacement[] requests, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(256L + requests.Sum(value => 128L * value.Hit.Fragments.Count), token);
        var seen = new HashSet<(int Page, string Path, int Offset)>();
        foreach (var request in requests)
            foreach (var fragment in request.Hit.Fragments)
            {
                token.ThrowIfCancellationRequested(); material.ChargeWork(64, token);
                material.ChargeWork(2L * fragment.TextCodePath.Length, token);
                var key = (request.Hit.PageIndex, fragment.TextCodePath, fragment.LogicalUtf16Start);
                if (!seen.Add(key)) throw new InvalidDataException("OverlappingReplacementRequests");
            }
    }
}
