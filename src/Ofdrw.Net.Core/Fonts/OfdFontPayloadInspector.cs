using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;

namespace Ofdrw.Net.Core.Fonts;

/// <summary>Font payload format recognized by <see cref="OfdFontPayloadInspector"/>.</summary>
public enum OfdFontPayloadFormat
{
    /// <summary>The bytes do not start with a supported single-face sfnt signature.</summary>
    Unknown,
    /// <summary>A single-face TrueType font (signature 0x00010000 or 'true') with <c>glyf</c> outlines.</summary>
    TrueType,
    /// <summary>A single-face OpenType font with CFF (version 1) outlines (signature 'OTTO').</summary>
    OpenTypeCff
}

/// <summary>Whether the inspection could interpret the payload.</summary>
public enum OfdFontPayloadStatus
{
    /// <summary>The table directory and requested coverage were read; see the report for missing items.</summary>
    Parsed,
    /// <summary>
    /// The format or a feature is outside the supported set (collections, WOFF, CFF2, variable, color or bitmap
    /// outlines, unrecognized signatures). This is an evidence gap, not a defect finding.
    /// </summary>
    Unsupported,
    /// <summary>The payload carries a supported signature but its structure cannot be read.</summary>
    Malformed,
    /// <summary>The configured byte or work budget was exhausted before the inspection finished.</summary>
    BudgetExceeded
}

/// <summary>What to verify in a font payload, and the bounds of that work. Values are copied when inspection starts.</summary>
public sealed class OfdFontPayloadRequest
{
    /// <summary>Unicode scalar values that must map, through the font's own Unicode cmap, to a readable glyph.</summary>
    public ICollection<int> UnicodeScalars { get; } = new List<int>();
    /// <summary>Glyph IDs (for example from OFD <c>CGTransform</c>) that must exist and have a readable outline.</summary>
    public ICollection<int> GlyphIds { get; } = new List<int>();
    /// <summary>Maximum payload size in bytes. Larger payloads are reported as <see cref="OfdFontPayloadStatus.BudgetExceeded"/>.</summary>
    public int MaxPayloadBytes { get; set; } = 32 * 1024 * 1024;
    /// <summary>Maximum cmap parse work units (bytes and records visited).</summary>
    public long MaxCmapWork { get; set; } = 64L * 1024 * 1024;
    /// <summary>Maximum distinct scalars plus glyph IDs checked in one call.</summary>
    public int MaxGlyphChecks { get; set; } = 65_536;
}

/// <summary>Read-only result of <see cref="OfdFontPayloadInspector.Inspect"/>. It never states licensing or authorisation.</summary>
public sealed class OfdFontPayloadReport
{
    private static readonly IReadOnlyList<string> None = new ReadOnlyCollection<string>(Array.Empty<string>());
    private static readonly IReadOnlyList<int> NoValues = new ReadOnlyCollection<int>(Array.Empty<int>());

    internal OfdFontPayloadReport(OfdFontPayloadStatus status, OfdFontPayloadFormat format, string reason)
    { Status = status; Format = format; Reason = reason; }

    /// <summary>Whether the payload could be interpreted.</summary>
    public OfdFontPayloadStatus Status { get; }
    /// <summary>Detected format; <see cref="OfdFontPayloadFormat.Unknown"/> when the signature is not supported.</summary>
    public OfdFontPayloadFormat Format { get; }
    /// <summary>Stable reason code, for example <c>Parsed</c>, <c>FontCollectionNotSupported</c> or <c>MalformedTableDirectory</c>.</summary>
    public string Reason { get; }
    /// <summary>Table tags present in the table directory, ordinal order.</summary>
    public IReadOnlyList<string> Tables { get; internal set; } = None;
    /// <summary>Outline tables required by <see cref="Format"/> (<c>glyf</c>/<c>loca</c> or <c>CFF </c>) that are absent.</summary>
    public IReadOnlyList<string> MissingOutlineTables { get; internal set; } = None;
    /// <summary>Tables from <see cref="OfdFontPayloadInspector.RequiredTables"/> that are absent.</summary>
    public IReadOnlyList<string> MissingRequiredTables { get; internal set; } = None;
    /// <summary>TrueType instruction tables present (<c>cvt </c>, <c>fpgm</c>, <c>prep</c>); their necessity is not judged.</summary>
    public IReadOnlyList<string> ProgramTables { get; internal set; } = None;
    /// <summary>Glyph count from <c>maxp</c>, or zero when unavailable.</summary>
    public int GlyphCount { get; internal set; }
    /// <summary>Whether the requested scalars and glyph IDs were checked against readable outlines.</summary>
    public bool CoverageChecked { get; internal set; }
    /// <summary>Requested scalars without a Unicode cmap mapping. Tab, line feed, carriage return and non-rendering controls are exempt.</summary>
    public IReadOnlyList<int> MissingScalars { get; internal set; } = NoValues;
    /// <summary>Requested glyph IDs that are zero (.notdef) or outside the font.</summary>
    public IReadOnlyList<int> MissingGlyphIds { get; internal set; } = NoValues;
    /// <summary>Mapped or requested glyph IDs whose outline data is malformed (for example truncated or inconsistent).</summary>
    public IReadOnlyList<int> UnreadableGlyphIds { get; internal set; } = NoValues;
    /// <summary>
    /// Mapped or requested glyph IDs whose outline the SDK reader does not interpret, such as point-matched or transformed
    /// composite components. Their renderability is undetermined; they are not reported as damaged.
    /// </summary>
    public IReadOnlyList<int> UninterpretedGlyphIds { get; internal set; } = NoValues;
    /// <summary>The OS/2 <c>fsType</c> bits recorded by the font, or null. Parseability never implies an embedding or licence right.</summary>
    public int? EmbeddingPermissions { get; internal set; }
    /// <summary>
    /// True when parsed, outline and required tables are present, and every requested item is confirmed renderable;
    /// false while any glyph is in <see cref="UninterpretedGlyphIds"/>.
    /// </summary>
    public bool CanRenderRequest => Status == OfdFontPayloadStatus.Parsed && CoverageChecked && MissingOutlineTables.Count == 0 &&
        MissingRequiredTables.Count == 0 && MissingScalars.Count == 0 && MissingGlyphIds.Count == 0 && UnreadableGlyphIds.Count == 0 &&
        UninterpretedGlyphIds.Count == 0;
}

/// <summary>
/// Bounded, read-only validation of font bytes, built on the SDK's own OpenType reader. Supported formats are single-face
/// static monochrome TrueType (<c>glyf</c>) and OpenType/CFF1. It does not look up host fonts, return font bytes,
/// render text or judge embedding rights. Malformed or unsupported payloads produce a status, never a format exception.
/// </summary>
public static class OfdFontPayloadInspector
{
    /// <summary>
    /// Tables this inspector treats as required in any embedded sfnt: character mapping, metrics, naming and the
    /// OS/2 and post tables required by the OpenType specification.
    /// </summary>
    public static IReadOnlyList<string> RequiredTables { get; } =
        new ReadOnlyCollection<string>(new[] { "cmap", "head", "hhea", "hmtx", "maxp", "name", "OS/2", "post" });

    private static readonly string[] UnsupportedFeatureTables =
        { "CFF2", "fvar", "gvar", "COLR", "SVG ", "CBDT", "CBLC", "EBDT", "EBLC", "sbix" };

    /// <summary>Inspects <paramref name="payload"/> against <paramref name="request"/>.</summary>
    /// <param name="payload">Font bytes; not modified and not retained.</param>
    /// <param name="request">Coverage to verify and budgets.</param>
    /// <param name="cancellationToken">Cancellation throws <see cref="OperationCanceledException"/>.</param>
    /// <returns>A report whose <see cref="OfdFontPayloadReport.Status"/> records unsupported, malformed or over-budget input.</returns>
    /// <exception cref="ArgumentNullException">An argument is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">A request budget is not positive.</exception>
    public static OfdFontPayloadReport Inspect(byte[] payload, OfdFontPayloadRequest request, CancellationToken cancellationToken = default)
    {
        if (payload == null) throw new ArgumentNullException(nameof(payload));
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.MaxPayloadBytes <= 0 || request.MaxCmapWork <= 0 || request.MaxGlyphChecks <= 0)
            throw new ArgumentOutOfRangeException(nameof(request));
        var scalars = request.UnicodeScalars.Distinct().OrderBy(x => x).ToArray();
        var glyphIds = request.GlyphIds.Distinct().OrderBy(x => x).ToArray();
        cancellationToken.ThrowIfCancellationRequested();
        if (payload.Length > request.MaxPayloadBytes || scalars.Length + (long)glyphIds.Length > request.MaxGlyphChecks)
            return new OfdFontPayloadReport(OfdFontPayloadStatus.BudgetExceeded, OfdFontPayloadFormat.Unknown, "PayloadBudgetExceeded");

        var signature = payload.Length >= 4 ? (uint)(payload[0] << 24 | payload[1] << 16 | payload[2] << 8 | payload[3]) : 0u;
        var format = signature is 0x00010000 or 0x74727565 ? OfdFontPayloadFormat.TrueType :
            signature == 0x4F54544F ? OfdFontPayloadFormat.OpenTypeCff : OfdFontPayloadFormat.Unknown;
        if (format == OfdFontPayloadFormat.Unknown)
            return new OfdFontPayloadReport(OfdFontPayloadStatus.Unsupported, format,
                signature == 0x74746366 ? "FontCollectionNotSupported" :
                signature is 0x774F4646 or 0x774F4632 ? "WoffNotSupported" : "UnrecognizedFontFormat");

        OpenTypeFace face;
        try { face = new OpenTypeFace(payload); }
        catch (Exception error) when (Recoverable(error))
        { return new OfdFontPayloadReport(OfdFontPayloadStatus.Malformed, format, "MalformedTableDirectory"); }
        cancellationToken.ThrowIfCancellationRequested();

        var tables = face.Tables.Keys.OrderBy(x => x, StringComparer.Ordinal).ToArray();
        var outlineTables = format == OfdFontPayloadFormat.TrueType ? new[] { "glyf", "loca" } : new[] { "CFF " };
        var unsupported = UnsupportedFeatureTables.Where(face.Tables.ContainsKey).ToArray();
        OfdFontPayloadReport Report(OfdFontPayloadStatus status, string reason)
        {
            var report = new OfdFontPayloadReport(status, format, reason)
            {
                Tables = Freeze(tables),
                MissingOutlineTables = Freeze(outlineTables.Where(x => !face.Tables.ContainsKey(x))),
                MissingRequiredTables = Freeze(RequiredTables.Where(x => !face.Tables.ContainsKey(x))),
                ProgramTables = Freeze(new[] { "cvt ", "fpgm", "prep" }.Where(face.Tables.ContainsKey))
            };
            try
            {
                report.EmbeddingPermissions = face.EmbeddingPermissions;
                if (face.Tables.TryGetValue("maxp", out var maxp) && maxp.Length >= 6) report.GlyphCount = OpenTypeFace.U16(maxp, 4);
            }
            catch (Exception error) when (Recoverable(error)) { }
            return report;
        }
        if (unsupported.Length > 0)
            return Report(OfdFontPayloadStatus.Unsupported, "UnsupportedOutlineFeature:" + string.Join(",", unsupported).Trim());
        if (outlineTables.Any(x => !face.Tables.ContainsKey(x)))
            return Report(OfdFontPayloadStatus.Parsed, "OutlineTablesMissing");

        IGlyphOutlineSource outlines;
        try { outlines = GlyphOutlineSources.Open(payload); }
        catch (NotSupportedException) { return Report(OfdFontPayloadStatus.Unsupported, "UnsupportedOutlineFeature"); }
        catch (Exception error) when (Recoverable(error)) { return Report(OfdFontPayloadStatus.Malformed, "MalformedOutlineTables"); }

        var budget = new FontParseBudget(request.MaxCmapWork, cancellationToken);
        OpenTypeCmap? cmap = null;
        string reason = "Parsed";
        if (scalars.Length > 0)
        {
            try { cmap = new OpenTypeCmap(face, budget); }
            catch (Exception error) when (Recoverable(error))
            {
                if (budget.IsExhausted) return Report(OfdFontPayloadStatus.BudgetExceeded, "CmapBudgetExceeded");
                if (error is NotSupportedException || !face.Tables.ContainsKey("cmap")) reason = "NoUnicodeCmap";
                else return Report(OfdFontPayloadStatus.Malformed, "MalformedCmap");
            }
        }

        var missingScalars = new List<int>();
        var missingGlyphs = new List<int>();
        var unreadable = new SortedSet<int>();
        var uninterpreted = new SortedSet<int>();
        var readable = new HashSet<int>();
        void Outline(int gid)
        {
            if (readable.Contains(gid) || unreadable.Contains(gid) || uninterpreted.Contains(gid)) return;
            try { outlines.Read(gid, cancellationToken); readable.Add(gid); }
            // The outline readers signal features outside their profile with NotSupportedException: undetermined, not damaged.
            catch (NotSupportedException) { uninterpreted.Add(gid); }
            catch (Exception error) when (Recoverable(error)) { unreadable.Add(gid); }
        }
        try
        {
            foreach (var scalar in scalars)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (scalar is 9 or 10 or 13 || UnicodeFontSubsetProfile.IsNonRenderingControl(scalar) || OpenTypeCmap.IsVariationSelector(scalar))
                    continue;
                var gid = cmap == null || scalar is < 0 or > 0x10FFFF ? 0 : cmap.Glyph(scalar);
                if (gid <= 0 || gid >= outlines.GlyphCount) missingScalars.Add(scalar);
                else Outline(gid);
            }
            foreach (var gid in glyphIds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (gid <= 0 || gid >= outlines.GlyphCount) missingGlyphs.Add(gid);
                else Outline(gid);
            }
        }
        catch (Exception error) when (Recoverable(error))
        {
            return budget.IsExhausted ? Report(OfdFontPayloadStatus.BudgetExceeded, "CmapBudgetExceeded")
                : Report(OfdFontPayloadStatus.Malformed, "MalformedCmap");
        }
        var result = Report(OfdFontPayloadStatus.Parsed, reason);
        result.CoverageChecked = true;
        result.MissingScalars = new ReadOnlyCollection<int>(missingScalars);
        result.MissingGlyphIds = new ReadOnlyCollection<int>(missingGlyphs);
        result.UnreadableGlyphIds = new ReadOnlyCollection<int>(unreadable.ToArray());
        result.UninterpretedGlyphIds = new ReadOnlyCollection<int>(uninterpreted.ToArray());
        return result;
    }

    private static bool Recoverable(Exception error) =>
        error is InvalidDataException or NotSupportedException or ArgumentException or IndexOutOfRangeException or
            OverflowException or FormatException or InvalidOperationException && error is not OperationCanceledException;

    private static IReadOnlyList<string> Freeze(IEnumerable<string> values) => new ReadOnlyCollection<string>(values.ToArray());
}
