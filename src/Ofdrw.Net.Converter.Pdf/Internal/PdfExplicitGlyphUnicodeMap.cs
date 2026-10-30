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
        IEnumerable<(string Identity, IReadOnlyDictionary<int, string> UnicodeByGlyph, bool NeedsOverride)> bindings)
    {
        var groups = bindings.GroupBy(binding => binding.Identity, StringComparer.Ordinal).ToArray();
        var pending = new Dictionary<string, int>(StringComparer.Ordinal);
        foreach (var group in groups)
        {
            var first = group.First();
            if (group.Skip(1).Any(binding =>
                    !SameMap(first.UnicodeByGlyph, binding.UnicodeByGlyph)))
                throw new NotSupportedException("One PDF font alias has conflicting CGTransform Unicode semantics.");
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

    private static bool SameMap(IReadOnlyDictionary<int, string> left, IReadOnlyDictionary<int, string> right) =>
        left.Count == right.Count && left.All(item =>
            right.TryGetValue(item.Key, out var value) &&
            string.Equals(item.Value, value, StringComparison.Ordinal));
}
