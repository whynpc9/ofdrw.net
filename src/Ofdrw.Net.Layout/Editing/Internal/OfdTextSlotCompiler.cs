using System.Globalization;
using System.IO;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Reader.Keywords.Internal;

namespace Ofdrw.Net.Layout.Editing.Internal;

internal static class OfdTextSlotCompiler
{
    internal static OfdTextSlotEdit[] Compile(OfdTextReplacementRoot root, OfdTextReplacement[] requests,
        OfdKeywordLocatorBudget budget, CancellationToken token, OfdTextGeometryExpectation.Builder? expectation = null)
    {
        var source = root.Source; var plan = root.Plan;
        budget.Copy(checked(4096L + 256L * source.Runs.Count + 512L * plan.Placements.Count + 16L * plan.Text.Length));
        var slots = new List<OfdTextFitPlacement>[source.Runs.Count];
        for (var i = 0; i < slots.Length; i++) slots[i] = new List<OfdTextFitPlacement>();
        var inserted = new List<(int Start, int End, int Slot)>(); var shift = 0;
        foreach (var requestPlan in plan.Requests)
        {
            token.ThrowIfCancellationRequested(); var request = requests[requestPlan.RequestIndex];
            var first = root.Fit.Clusters[requestPlan.ScalarStart];
            if (requestPlan.Disposition != OfdTextFitDisposition.Unchanged)
            {
                var start = first.Metric.Utf16Start + shift;
                inserted.Add((start, start + request.ReplacementText.Length, Slot(first.Slot)));
                shift += request.ReplacementText.Length - request.Hit.MatchedText.Length;
            }
        }
        var insertedCursor = 0;
        foreach (var placement in plan.Placements)
        {
            token.ThrowIfCancellationRequested(); budget.Correlate(); int slot;
            if (placement.Source is { } original) slot = Slot(original.Slot);
            else
            {
                while (insertedCursor < inserted.Count && placement.OutputUtf16Start >= inserted[insertedCursor].End)
                { budget.Correlate(); insertedCursor++; }
                if (insertedCursor == inserted.Count || placement.OutputUtf16Start < inserted[insertedCursor].Start)
                    throw new InvalidDataException("ReplacementInsertionRangeMismatch");
                slot = inserted[insertedCursor].Slot;
            }
            slots[slot].Add(placement);
        }
        var edits = new OfdTextSlotEdit[slots.Length]; var whole = new StringBuilder(plan.Text.Length);
        for (var i = 0; i < slots.Length; i++)
        {
            token.ThrowIfCancellationRequested(); var placements = slots[i]; var original = source.Runs[i];
            if (placements.Count == 0)
            {
                edits[i] = new OfdTextSlotEdit(i, string.Empty, original.X, original.Y); continue;
            }
            budget.Copy(checked(512L + placements.Count * 256L));
            var text = new StringBuilder(); var dx = new StringBuilder(); var dy = new StringBuilder();
            var xs = new double[placements.Count]; var ys = new double[placements.Count];
            for (var p = 0; p < placements.Count; p++)
            {
                token.ThrowIfCancellationRequested(); var placement = placements[p];
                xs[p] = placement.X - source.Matrix[4]; ys[p] = placement.Y - source.Matrix[5];
                if (!Ofdrw.Net.Reader.Keywords.OfdGlyphBox.Finite(xs[p]) || !Ofdrw.Net.Reader.Keywords.OfdGlyphBox.Finite(ys[p]) ||
                    xs[p] + source.Matrix[4] != placement.X || ys[p] + source.Matrix[5] != placement.Y)
                    throw new NotSupportedException("ReplacementCoordinatePrecisionLoss");
                text.Append(plan.Text, placement.OutputUtf16Start, placement.Metric.Utf16Length);
                if (p == 0) continue;
                if (p > 1) { dx.Append(' '); dy.Append(' '); }
                dx.Append((xs[p] - xs[p - 1]).ToString("R", CultureInfo.InvariantCulture));
                dy.Append((ys[p] - ys[p - 1]).ToString("R", CultureInfo.InvariantCulture));
            }
            var raw = text.ToString(); var deltaX = placements.Count == 1 ? null : dx.ToString();
            var deltaY = placements.Count == 1 ? null : dy.ToString();
            // Replay the actual serialized strings with the same strict decoder
            // and cumulative delta arithmetic used by the public locator.
            var decoded = OfdStrictTextGeometry.Decode(raw, budget);
            if (decoded.Text != raw) throw new InvalidDataException("ReplacementSlotEncodingMismatch");
            var replayX = OfdStrictTextGeometry.Deltas(deltaX, placements.Count - 1, budget);
            var replayY = OfdStrictTextGeometry.Deltas(deltaY, placements.Count - 1, budget);
            var x = double.Parse(xs[0].ToString("R", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            var y = double.Parse(ys[0].ToString("R", CultureInfo.InvariantCulture), CultureInfo.InvariantCulture);
            for (var p = 0; p < placements.Count; p++)
            {
                token.ThrowIfCancellationRequested();
                if (p > 0) { x += replayX[p - 1]; y += replayY[p - 1]; }
                if (x != xs[p] || y != ys[p] || x + source.Matrix[4] != placements[p].X || y + source.Matrix[5] != placements[p].Y)
                    throw new NotSupportedException("ReplacementDeltaPrecisionLoss");
            }
            whole.Append(raw); edits[i] = new OfdTextSlotEdit(i, raw, xs[0], ys[0], deltaX, deltaY);
        }
        if (whole.ToString() != plan.Text) throw new InvalidDataException("ReplacementSlotOrderMismatch");
        expectation?.RecordCompiled(root, edits, slots);
        return edits;
    }
    private static int Slot(string token) => int.Parse(token.Substring(1), NumberStyles.None, CultureInfo.InvariantCulture);
}
