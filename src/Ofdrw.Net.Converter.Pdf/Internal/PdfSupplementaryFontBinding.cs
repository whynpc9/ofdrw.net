using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Pdf;

namespace Ofdrw.Net.Converter.Pdf.Internal;

// PDFsharp 1.3.67 indexes its cmap and ToUnicode by UTF-16 char. A supplementary
// scalar would otherwise become two .notdef glyphs. This limited static glyf
// profile uses a spare BMP private-use code only inside the PDF font. The OFD
// remains unchanged; the PDF's CID (the original GID) maps back to the exact
// UTF-16 surrogate pair in a separate ToUnicode stream.
internal sealed class PdfSupplementaryFontBinding
{
    private const int MaximumScalars = 1024;
    private static readonly string[] UnsupportedOutlineTables =
    [
        "CFF ", "CFF2", "fvar", "gvar", "avar", "HVAR", "VVAR", "MVAR",
        "COLR", "CPAL", "SVG ", "CBDT", "CBLC", "sbix"
    ];
    private readonly IReadOnlyDictionary<int, int> _pdfCodeByScalar;
    private readonly IReadOnlyDictionary<int, int> _unicodeByGlyph;

    private PdfSupplementaryFontBinding(string family, string identity, byte[] originalFont,
        IReadOnlyDictionary<int, int> pdfCodeByScalar, IReadOnlyDictionary<int, int> unicodeByGlyph)
    {
        Family = family;
        Identity = identity;
        OriginalFont = originalFont;
        _pdfCodeByScalar = pdfCodeByScalar;
        _unicodeByGlyph = unicodeByGlyph;
    }

    internal string Family { get; }
    internal string Identity { get; }
    internal byte[] OriginalFont { get; }

    internal bool HasSameUnicodeMap(PdfSupplementaryFontBinding other) =>
        _unicodeByGlyph.Count == other._unicodeByGlyph.Count &&
        _unicodeByGlyph.All(item => other._unicodeByGlyph.TryGetValue(item.Key, out var scalar) && scalar == item.Value);

    internal static PdfSupplementaryFontBinding Create(byte[] originalFont, OpenTypeCmap coverage,
        OfdTextElement text, bool bold, bool italic, FontParseBudget budget, CancellationToken cancellationToken)
    {
        if (originalFont.LongLength > 64L * 1024 * 1024)
            throw new NotSupportedException("Supplementary PDF font exceeds the bounded face profile.");
        var face = new OpenTypeFace(originalFont);
        if (face.Version != 0x00010000 || !face.Tables.ContainsKey("glyf") || !face.Tables.ContainsKey("loca") ||
            UnsupportedOutlineTables.Any(face.Tables.ContainsKey))
            throw new NotSupportedException("Supplementary PDF text requires a static TrueType glyf face; the OFD font remains preserved.");
        if (!face.Tables.TryGetValue("OS/2", out var os2) || os2.Length < 10 ||
            (OpenTypeFace.U16(os2, 8) & (0x0002 | 0x0100 | 0x0200)) != 0)
            throw new NotSupportedException("Supplementary PDF font does not permit the verified embedding and cmap-alias profile.");
        var glyphCount = OpenTypeFace.U16(face.Table("maxp", 6), 4);
        var glyf = face.Table("glyf");
        var loca = face.Table("loca");
        var locaFormat = OpenTypeFace.U16(face.Table("head", 54), 50);
        if (glyphCount == 0 || locaFormat is not (0 or 1))
            throw new NotSupportedException("Supplementary PDF font has unsupported glyph locations.");
        OpenTypeFace.Require(loca, 0, (glyphCount + 1) * (locaFormat == 0 ? 2 : 4));

        var scalars = new SortedSet<int>();
        var scalarCount = 0;
        var strings = text.Runs.Count == 0 ? new[] { text.Text } : text.Runs.Select(run => run.Text);
        foreach (var value in strings)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var visible = PdfTextControlPolicy.VisibleText(value, coverage);
            foreach (var scalar in OpenTypeFace.Scalars(visible))
            {
                if (++scalarCount > MaximumScalars)
                    throw new NotSupportedException("Supplementary PDF text exceeds the bounded scalar profile.");
                scalars.Add(scalar);
            }
            // DrawStyledString may compose a single Latin grapheme. Include the
            // same face's composed glyph so drawing cannot escape this mapping.
            foreach (var scalar in OpenTypeFace.Scalars(visible.Normalize(NormalizationForm.FormC))) scalars.Add(scalar);
            if (scalars.Count > MaximumScalars)
                throw new NotSupportedException("Supplementary PDF text exceeds the bounded scalar profile.");
        }

        var pdfCodeByScalar = new Dictionary<int, int>();
        var glyphByPdfCode = new Dictionary<int, int>();
        var unicodeByGlyph = new Dictionary<int, int>();
        var nextPrivateUse = 0xE000;
        foreach (var scalar in scalars)
        {
            cancellationToken.ThrowIfCancellationRequested();
            budget.Spend(1);
            var glyph = coverage.Glyph(scalar);
            if (glyph <= 0 || glyph >= glyphCount)
                throw new NotSupportedException($"Supplementary PDF text has no distinct outline for U+{scalar:X}.");
            if (scalar > 0xFFFF)
            {
                var stride = locaFormat == 0 ? 2 : 4;
                var start = locaFormat == 0 ? OpenTypeFace.U16(loca, glyph * stride) * 2 :
                    OpenTypeFace.Size(OpenTypeFace.U32(loca, glyph * stride));
                var end = locaFormat == 0 ? OpenTypeFace.U16(loca, (glyph + 1) * stride) * 2 :
                    OpenTypeFace.Size(OpenTypeFace.U32(loca, (glyph + 1) * stride));
                if (start > glyf.Length || end > glyf.Length || end - start < 10 ||
                    unchecked((short)OpenTypeFace.U16(glyf, start)) == 0)
                    throw new NotSupportedException($"Supplementary PDF text has no verified outline for U+{scalar:X}.");
            }
            if (unicodeByGlyph.TryGetValue(glyph, out var previous) && previous != scalar)
                throw new NotSupportedException("Supplementary PDF text has ambiguous Unicode for one glyph ID.");
            unicodeByGlyph[glyph] = scalar;

            var pdfCode = scalar;
            if (scalar > 0xFFFF)
            {
                while (nextPrivateUse <= 0xF8FF &&
                       (scalars.Contains(nextPrivateUse) || coverage.Glyph(nextPrivateUse) != 0))
                {
                    budget.Spend(1);
                    nextPrivateUse++;
                }
                if (nextPrivateUse > 0xF8FF)
                    throw new NotSupportedException("Supplementary PDF text has no spare private-use code.");
                pdfCode = nextPrivateUse++;
            }
            pdfCodeByScalar.Add(scalar, pdfCode);
            glyphByPdfCode.Add(pdfCode, glyph);
        }

        // Keep original GIDs while removing outlines that this text object
        // cannot draw. The registry is process-wide, so retaining a full face
        // for each distinct PDF cmap would quickly exhaust its byte budget.
        var subsetFace = new OpenTypeFace(originalFont);
        var subsetBytes = TrueTypeSubset.Create(subsetFace, scalars, cancellationToken, 2_000_000,
            budget, out _, out _, coverage);
        var aliasedFace = new OpenTypeFace(subsetBytes);
        aliasedFace.Tables["cmap"] = OpenTypeCmap.Build(glyphByPdfCode);
        var aliasedBytes = aliasedFace.Build();
        var identity = BinaryIdentity.Hash(aliasedBytes);
        var family = PdfFontRegistry.RegisterFontFace(aliasedBytes, bold, italic);
        return new PdfSupplementaryFontBinding(family, identity, originalFont, pdfCodeByScalar, unicodeByGlyph);
    }

    internal string MapText(string logicalText)
    {
        var result = new StringBuilder(logicalText.Length);
        foreach (var scalar in OpenTypeFace.Scalars(logicalText))
        {
            if (!_pdfCodeByScalar.TryGetValue(scalar, out var pdfCode))
                throw new NotSupportedException($"Supplementary PDF font lacks an admitted scalar U+{scalar:X}.");
            result.Append((char)pdfCode);
        }
        return result.ToString();
    }

    internal static void InstallToUnicodeMaps(PdfDocument document, IEnumerable<PdfSupplementaryFontBinding> bindings)
    {
        var unique = bindings.GroupBy(binding => binding.Identity, StringComparer.Ordinal)
            .Select(group =>
            {
                var first = group.First();
                if (group.Skip(1).Any(binding => !first.HasSameUnicodeMap(binding)))
                    throw new NotSupportedException("Supplementary PDF text has conflicting Unicode for one registered font identity.");
                return first;
            }).ToArray();
        if (unique.Length == 0) return;
        var pending = unique.ToDictionary(binding => binding.Identity, _ => 0, StringComparer.Ordinal);
        foreach (var font in document.Internals.GetAllObjects().OfType<PdfDictionary>())
        {
            if (font.Elements.GetName("/Subtype") != "/Type0") continue;
            var baseName = font.Elements.GetName("/BaseFont");
            foreach (var binding in unique)
            {
                // The installed resolver gives PDFsharp this content-specific
                // PostScript name. The six-letter PDF subset prefix may vary.
                if (baseName.IndexOf(binding.Identity.Substring(0, 62), StringComparison.OrdinalIgnoreCase) < 0) continue;
                var unicode = new PdfDictionary(document);
                unicode.CreateStream(BuildToUnicode(binding._unicodeByGlyph));
                document.Internals.AddObject(unicode);
                font.Elements["/ToUnicode"] = unicode.Reference;
                pending[binding.Identity]++;
                break;
            }
        }
        if (pending.Any(item => item.Value == 0))
            throw new InvalidDataException("PDF supplementary font was not realized with its verified ToUnicode map.");
    }

    private static byte[] BuildToUnicode(IReadOnlyDictionary<int, int> unicodeByGlyph)
    {
        var builder = new StringBuilder();
        builder.Append("/CIDInit /ProcSet findresource begin\n12 dict begin\nbegincmap\n")
            .Append("/CIDSystemInfo << /Registry (Adobe) /Ordering (UCS) /Supplement 0 >> def\n")
            .Append("/CMapName /Adobe-Identity-UCS def\n/CMapType 2 def\n")
            .Append("1 begincodespacerange\n<0000><FFFF>\nendcodespacerange\n");
        var ordered = unicodeByGlyph.OrderBy(item => item.Key).ToArray();
        for (var start = 0; start < ordered.Length; start += 100)
        {
            var count = Math.Min(100, ordered.Length - start);
            builder.Append(count).Append(" beginbfchar\n");
            for (var index = start; index < start + count; index++)
            {
                var item = ordered[index];
                var unicode = Encoding.BigEndianUnicode.GetBytes(char.ConvertFromUtf32(item.Value));
                builder.Append('<').Append(item.Key.ToString("X4")).Append("><")
                    .Append(BitConverter.ToString(unicode).Replace("-", string.Empty)).Append(">\n");
            }
            builder.Append("endbfchar\n");
        }
        builder.Append("endcmap\nCMapName currentdict /CMap defineresource pop\nend\nend\n");
        return Encoding.ASCII.GetBytes(builder.ToString());
    }
}
