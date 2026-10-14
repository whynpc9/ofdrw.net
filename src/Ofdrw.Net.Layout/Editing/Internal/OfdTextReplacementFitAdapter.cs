using System.IO;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Layout.Editing.Internal;

internal sealed class OfdTextReplacementRoot
{
    internal OfdTextElement Owner = null!;
    internal OfdKeywordTextSource Source = null!;
    internal OfdTextFitObject Fit = null!;
    internal OfdTextFitObjectPlan Plan = null!;
    internal OfdTextSlotEdit[] Slots = null!;
}

internal static class OfdTextReplacementFitAdapter
{
    internal static List<OfdTextReplacementRoot> Plan(OfdKeywordAnalysisCapture capture,
        OfdTextReplacement[] requests, OfdTextReplacementOperation operation, CancellationToken token, HashSet<int> affectedPages, OfdTextGeometryExpectation.Builder? expectation = null)
    {
        var budget = capture.Budget; var objectCount = 0;
        foreach (var occurrence in capture.Traversal.Occurrences)
        {
            token.ThrowIfCancellationRequested(); budget.Correlate();
            if (affectedPages.Contains(occurrence.OccurrencePath[0]) && ++objectCount > OfdTextFitLimits.HardObjects)
                throw new InvalidDataException("ReplacementPageObjectLimit");
        }
        budget.Copy(4096L + 1024L * objectCount + 512L * requests.Length);
        var roots = new Dictionary<OfdKeywordTextSource, OfdTextReplacementRoot>();
        var rootsByOccurrence = new Dictionary<string, OfdTextReplacementRoot>(StringComparer.Ordinal);
        var fitObjects = new List<OfdTextFitObject>(); var fitRequests = new List<OfdTextFitRequest>();
        foreach (var occurrence in capture.Traversal.Occurrences)
        {
            token.ThrowIfCancellationRequested(); budget.Correlate();
            if (!affectedPages.Contains(occurrence.OccurrencePath[0])) continue;
            if (occurrence.Element is not OfdTextElement owner || !capture.Sources.TryGetValue(occurrence, out var source) ||
                !source.Visible || !source.HasXml || !source.Editable || occurrence.IsInComposite)
                throw new NotSupportedException("UnprovedReplacementNeighbors");
            if (source.Matrix[0] != 1 || source.Matrix[1] != 0 || source.Matrix[2] != 0 || source.Matrix[3] != 1)
                throw new NotSupportedException("UnsupportedReplacementTransform");
            foreach (var run in source.Runs)
            {
                if (run.RawText != run.Text) throw new NotSupportedException("UnsafeLegacyRawTextEncoding");
                budget.Work(2L * run.Text.Length);
                OfdTextReplacer.ValidateLiteral(run.Text, token);
            }
            var positions = capture.Positions[source];
            if (positions.Count > OfdTextFitLimits.HardClusters) throw new InvalidDataException("ReplacementSourceClusterLimit");
            long units = 0; foreach (var run in source.Runs) units += run.Text.Length;
            if (units > OfdTextFitLimits.HardUtf16) throw new InvalidDataException("ReplacementSourceTextLimit");
            budget.Copy(checked(2048L + 768L * positions.Count + 8L * units));
            var text = string.Concat(source.Runs.Select(run => run.Text));
            var clusters = new List<OfdTextFitSourceCluster>(positions.Count); var utf16 = 0;
            budget.Copy(128L * source.Runs.Count + 256);
            var runIndices = new Dictionary<OfdKeywordSourceRun, int>();
            for (var r = 0; r < source.Runs.Count; r++) { budget.Correlate(); runIndices.Add(source.Runs[r], r); }
            foreach (var position in positions)
            {
                token.ThrowIfCancellationRequested();
                budget.Correlate(); var runIndex = runIndices[position.Run]; var geometry = position.Geometry;
                var x = position.RawX + source.Matrix[4]; var y = position.RawY + source.Matrix[5];
                var tx = source.X - source.PageX; var ty = source.Y - source.PageY;
                if (!OfdGlyphBox.Finite(x) || !OfdGlyphBox.Finite(y) || x + tx != position.Baseline.X || y + ty != position.Baseline.Y ||
                    x - source.Matrix[4] != position.RawX || y - source.Matrix[5] != position.RawY)
                    throw new NotSupportedException("ReplacementCoordinatePrecisionLoss");
                var metric = Metric(char.ConvertToUtf32(position.Run.Text, geometry.Utf16Start), utf16, geometry);
                clusters.Add(new OfdTextFitSourceCluster(metric, x, y, "s" + runIndex,
                    position.Run.SourceOffsets[geometry.Utf16Start],
                    position.Run.SourceOffsets[geometry.Utf16Start + geometry.Utf16Length] - position.Run.SourceOffsets[geometry.Utf16Start]));
                utf16 += geometry.Utf16Length;
            }
            var key = "o" + roots.Count;
            var fit = new OfdTextFitObject(source.PageIndex, key, key, "edit", "body", key, text, clusters,
                new OfdTextFitAffine(1, 0, 0, 1, source.X - source.PageX, source.Y - source.PageY),
                new OfdTextFitBox(0, 0, source.Width, source.Height), new OfdTextFitBox(0, 0, source.PageWidth, source.PageHeight),
                cancellationToken: token);
            var root = new OfdTextReplacementRoot { Owner = owner, Source = source, Fit = fit };
            roots.Add(source, root); rootsByOccurrence.Add(key, root); fitObjects.Add(fit);
            expectation?.CaptureRoot(capture, occurrence, root);
        }
        for (var i = 0; i < requests.Length; i++)
        {
            token.ThrowIfCancellationRequested(); var request = requests[i];
            var mapped = OfdKeywordEditSource.Authenticate(capture, request.Hit);
            if (!affectedPages.Contains(mapped.Source.PageIndex)) throw new InvalidDataException("ReplacementAffectedPageMismatch");
            var source = mapped.Source; var root = roots[source];
            var geometry = capture.Measure(source, request.ReplacementText);
            budget.Copy(512L + 256L * geometry.Clusters.Count + 4L * request.ReplacementText.Length);
            var metrics = geometry.Clusters.Select(value => Metric(char.ConvertToUtf32(request.ReplacementText, value.Utf16Start), value.Utf16Start, value)).ToArray();
            expectation?.RecordReplacementMetrics(root, geometry, metrics);
            fitRequests.Add(new OfdTextFitRequest(root.Fit.Occurrence, mapped.Start, mapped.Count,
                request.Hit.MatchedText, request.ReplacementText, root.Fit.FontStyle, metrics, token));
        }
        operation.Begin("Fit");
        // Pre-reserve the planner's copied input/output positions and local
        // temporary structures; its own independent finite counters also apply.
        long inputCount = fitObjects.Sum(value => (long)value.Clusters.Count);
        long replacementCount = fitRequests.Sum(value => (long)value.Metrics.Count);
        budget.Copy(checked(4096L + 1024L * (inputCount + replacementCount) + 256L * fitRequests.Count));
        var input = new OfdTextFitInput("edit", "body", fitObjects, fitRequests, Array.Empty<OfdTextFitNeighbor>(), token);
        var result = OfdTextFitPlanner.Plan(input, new OfdTextFitLimits(), token, budget.Work);
        operation.RecordFit(result.Statistics);
        if (!result.FitsSuppliedGeometry) throw new NotSupportedException("ReplacementFit:" + result.Failure);
        operation.Mark("Fitted"); operation.Begin("CompileSlots");
        var planned = new List<OfdTextReplacementRoot>();
        foreach (var plan in result.Plans)
        {
            budget.Correlate(); var root = rootsByOccurrence[plan.Occurrence];
            root.Plan = plan; root.Slots = OfdTextSlotCompiler.Compile(root, requests, budget, token, expectation); planned.Add(root);
        }
        expectation?.FreezeBeforeProviderClosure();
        operation.Mark("CompiledSlots"); return planned;
    }
    private static OfdTextFitMetric Metric(int scalar, int offset, OfdGlyphClusterGeometry geometry)
    {
        var box = geometry.FontBoundsMillimeters;
        return new OfdTextFitMetric(scalar, offset, geometry.Utf16Length, geometry.GlyphId,
            geometry.AdvanceXMillimeters, geometry.AdvanceYMillimeters,
            box is null ? null : new OfdTextFitBox(box.Left, box.Top, box.Right, box.Bottom));
    }
}
