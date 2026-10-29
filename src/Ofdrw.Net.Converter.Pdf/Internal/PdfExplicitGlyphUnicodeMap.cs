using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using PdfSharpCore.Pdf;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// A private PDF code can carry one source character span to one original GID.
// The visible CID must recover the exact source span, rather than the alias.
internal static class PdfExplicitGlyphUnicodeMap
{
    internal static void Install(PdfDocument document,
        IEnumerable<(string Identity, IReadOnlyDictionary<int, string> UnicodeByGlyph, bool NeedsOverride,
            IReadOnlyDictionary<int, int> WidthByGlyph)> bindings)
    {
        var groups = bindings.GroupBy(binding => binding.Identity, StringComparer.Ordinal).ToArray();
        var pending = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            var first = group.First();
            if (group.Skip(1).Any(binding =>
                    !SameMap(first.UnicodeByGlyph, binding.UnicodeByGlyph) ||
                    !SameWidths(first.WidthByGlyph, binding.WidthByGlyph)))
                throw new NotSupportedException("One PDF font alias has conflicting CGTransform Unicode or width semantics.");
            if (group.Any(binding => binding.NeedsOverride)) pending.Add(group.Key, 0);
        }
        if (pending.Count == 0) return;
        foreach (var font in document.Internals.GetAllObjects().OfType<PdfDictionary>())
        {
            if (font.Elements.GetName("/Subtype") != "/Type0") continue;
            var baseName = font.Elements.GetName("/BaseFont");
            foreach (var group in groups)
            {
                if (!pending.ContainsKey(group.Key) ||
                    baseName.IndexOf(group.Key.Substring(0, 62), StringComparison.OrdinalIgnoreCase) < 0) continue;
                var unicode = new PdfDictionary(document);
                unicode.CreateStream(PdfSupplementaryFontBinding.BuildToUnicode(group.First().UnicodeByGlyph));
                document.Internals.AddObject(unicode);
                font.Elements["/ToUnicode"] = unicode.Reference;
                pending[group.Key]++;
                break;
            }
        }
        if (pending.Any(item => item.Value == 0))
            throw new InvalidDataException("CGTransform source-span PDF font was not realized with its Unicode map.");
    }

    // PDFsharp builds descendant /W during Save, after Install has run. Patch
    // the saved document in the same pass that restores bounded FontFile2 faces.
    internal static void InstallWidths(PdfDocument document,
        IEnumerable<(string Identity, IReadOnlyDictionary<int, int> WidthByGlyph)> bindings)
    {
        var groups = bindings.Where(binding => binding.WidthByGlyph.Count > 0)
            .GroupBy(binding => binding.Identity, StringComparer.Ordinal).ToArray();
        var widthPending = groups.ToDictionary(group => group.Key, _ => 0, StringComparer.Ordinal);
        foreach (var group in groups)
            if (group.Skip(1).Any(binding => !SameWidths(group.First().WidthByGlyph, binding.WidthByGlyph)))
                throw new NotSupportedException("One PDF font alias has conflicting positioned widths.");
        foreach (var font in document.Internals.GetAllObjects().OfType<PdfDictionary>())
        {
            if (font.Elements.GetName("/Subtype") != "/CIDFontType2") continue;
            var baseName = font.Elements.GetName("/BaseFont");
            foreach (var group in groups)
            {
                if (!widthPending.ContainsKey(group.Key) ||
                    baseName.IndexOf(group.Key.Substring(0, 62), StringComparison.OrdinalIgnoreCase) < 0) continue;
                var width = group.First().WidthByGlyph.Single();
                if (font.Elements["/W"] is not PdfArray existing || existing.Elements.Count != 2 ||
                    existing.Elements[0] is not PdfInteger code || code.Value != width.Key ||
                    existing.Elements[1] is not PdfArray values || values.Elements.Count != 1 ||
                    values.Elements[0] is not PdfInteger)
                    throw new NotSupportedException("PDF source-span width override requires one verified CID width.");
                var replacement = new PdfArray(document);
                replacement.Elements.Add(new PdfInteger(width.Key));
                var widths = new PdfArray(document);
                widths.Elements.Add(new PdfInteger(width.Value));
                replacement.Elements.Add(widths);
                font.Elements["/W"] = replacement;
                widthPending[group.Key]++;
                break;
            }
        }
        if (widthPending.Any(item => item.Value == 0))
            throw new InvalidDataException("Hebrew PDF source-span width was not installed on its descendant font.");
    }

    private static bool SameMap(IReadOnlyDictionary<int, string> left, IReadOnlyDictionary<int, string> right) =>
        left.Count == right.Count && left.All(item =>
            right.TryGetValue(item.Key, out var value) &&
            string.Equals(item.Value, value, StringComparison.Ordinal));

    private static bool SameWidths(IReadOnlyDictionary<int, int> left, IReadOnlyDictionary<int, int> right) =>
        left.Count == right.Count && left.All(item =>
            right.TryGetValue(item.Key, out var value) && item.Value == value);
}
