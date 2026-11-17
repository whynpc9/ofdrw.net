using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Archival.Internal;

/// <summary>One registered font rule; all four share a single <see cref="ArchivalFontAnalysis"/> per assessment.</summary>
internal sealed class ArchivalFontCheck : IArchivalClauseCheck
{
    internal ArchivalFontCheck(ArchivalOwnedRuleDefinition rule) { Rule = rule; }
    public ArchivalOwnedRuleDefinition Rule { get; }
    // All four font rules own the gap: it closes only when AC-017..AC-020 are all decided.
    public string? GapCode => "FontPayloadNotAssessed";
    public IEnumerable<ArchivalClauseObservation> Check(ArchivalCheckContext context) =>
        context.Shared(() => ArchivalFontAnalysis.Run(context)).For(Rule);
    public bool Decided(IReadOnlyList<ArchivalClauseObservation> observations) =>
        observations.All(x => x.Outcome != PreflightOutcome.Unknown);
}

/// <summary>
/// Font payload and used-glyph analysis for 6.2.6(b)(1), (b)(2), (c)(1) and (c)(2). Fonts are resolved through the
/// declaring resource table and its BaseLoc; used characters and glyph IDs come from TextCode and CGTransform in pages
/// and templates. Bytes are judged with <see cref="OfdFontPayloadInspector"/>; declared names and host fonts are never
/// evidence, and parseability never implies an embedding or licence right.
/// </summary>
internal sealed class ArchivalFontAnalysis
{
    internal static readonly ArchivalOwnedRuleDefinition UsedTextRule = new ArchivalOwnedRuleDefinition("AC-017", "6.2.6(b)(1)",
        "Every font used by a filled or stroked TextObject has an embedded payload in the package that renders the characters and glyph IDs actually used");
    internal static readonly ArchivalOwnedRuleDefinition ClippingTextRule = new ArchivalOwnedRuleDefinition("AC-018", "6.2.6(b)(2)",
        "Every font used by clipping-area text has an embedded payload in the package that renders the characters and glyph IDs actually used");
    internal static readonly ArchivalOwnedRuleDefinition OutlineRule = new ArchivalOwnedRuleDefinition("AC-019", "6.2.6(c)(1)",
        "Each embedded font payload keeps its outline tables (glyf/loca or CFF) and the outlines of the glyphs used");
    internal static readonly ArchivalOwnedRuleDefinition TablesRule = new ArchivalOwnedRuleDefinition("AC-020", "6.2.6(c)(2)",
        "Each embedded font payload keeps cmap, head, hhea, hmtx, maxp, name, OS/2 and post");

    private static readonly HashSet<string> TextScopeFindings = new HashSet<string>(StringComparer.Ordinal)
    {
        "DocumentScopeUnavailable", "UnrecognizedDocumentContent", "PageContentUnavailable", "UnrecognizedPageContent",
        "UnrecognizedPageObject", "CompositeGraphicsPresent", "AnnotationsDeclared", "ExtensionsDeclared", "ScopeInspectionBudgetExceeded"
    };

    private readonly ArchivalCheckContext context;
    private readonly Dictionary<string, List<ArchivalClauseObservation>> observations = new Dictionary<string, List<ArchivalClauseObservation>>(StringComparer.Ordinal);
    private readonly List<FontDeclaration> declarations = new List<FontDeclaration>();
    private readonly Dictionary<string, List<FontDeclaration>> tables = new Dictionary<string, List<FontDeclaration>>(StringComparer.OrdinalIgnoreCase);
    private bool resourceScopeComplete = true, usageScopeComplete = true;
    private bool anyVisibleText, anyClipText;
    private long fontBytes;

    private ArchivalFontAnalysis(ArchivalCheckContext context)
    {
        this.context = context;
        foreach (var rule in new[] { UsedTextRule, ClippingTextRule, OutlineRule, TablesRule })
            observations[rule.RuleId] = new List<ArchivalClauseObservation>();
    }

    internal IEnumerable<ArchivalClauseObservation> For(ArchivalOwnedRuleDefinition rule) => observations[rule.RuleId];

    internal static ArchivalFontAnalysis Run(ArchivalCheckContext context)
    {
        var analysis = new ArchivalFontAnalysis(context);
        analysis.Analyse();
        return analysis;
    }

    private void Add(ArchivalOwnedRuleDefinition rule, PreflightOutcome outcome, string reason, string entry, XElement? node, string observed)
    {
        context.SpendResult();
        observations[rule.RuleId].Add(ArchivalOwnedRules.Observation(rule, outcome, reason, entry,
            node == null ? "<package>" : ArchivalPackagePath.Selector(node), observed));
    }

    private void Analyse()
    {
        var scope = context.Scope;
        if (scope.Document == null || scope.DocumentEntry == null)
        {
            foreach (var rule in new[] { UsedTextRule, ClippingTextRule, OutlineRule, TablesRule })
                Add(rule, PreflightOutcome.Unknown, "DocumentScopeUnavailable", "<document>", null, "Document not resolved by scope inspection");
            return;
        }
        foreach (var finding in scope.Findings.Where(x => TextScopeFindings.Contains(x.Code)))
        {
            usageScopeComplete = false;
            foreach (var rule in new[] { UsedTextRule, ClippingTextRule })
            {
                context.SpendResult();
                observations[rule.RuleId].Add(ArchivalOwnedRules.Observation(rule, PreflightOutcome.Unknown, "TextScopeIncomplete:" + finding.Code,
                    finding.Evidence.Entry, finding.Evidence.Selector, finding.Evidence.Observed));
            }
        }

        var ns = scope.Document.Name.Namespace;
        var documentFonts = new Dictionary<string, FontDeclaration?>(StringComparer.Ordinal);
        foreach (var reference in scope.Document.Elements(ns + "CommonData").SelectMany(x => x.Elements())
                     .Where(x => x.Name == ns + "PublicRes" || x.Name == ns + "DocumentRes"))
            ReadResources(scope.DocumentEntry, reference, documentFonts);

        var uses = new Dictionary<FontDeclaration, Usage>();
        foreach (var page in scope.Pages)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            var pns = page.Value.Name.Namespace;
            var pageFonts = new Dictionary<string, FontDeclaration?>(documentFonts, StringComparer.Ordinal);
            // The reader and preflight treat {pageDirectory}/PageRes.xml as a fixed page resource even without a PageRes link.
            var sibling = page.Key.Substring(0, page.Key.LastIndexOf('/') + 1) + "PageRes.xml";
            if (context.Archive.Contains(sibling)) ReadTable(page.Key, sibling, page.Value, pageFonts);
            foreach (var reference in page.Value.Elements(pns + "PageRes")) ReadResources(page.Key, reference, pageFonts);
            foreach (var text in page.Value.Descendants(pns + "TextObject"))
            {
                // Text that is neither filled nor stroked, or has no code points, draws nothing with its font.
                if (!Painted(text) || Empty(text)) continue;
                anyVisibleText = true;
                Use(page.Key, text, pageFonts, uses, UsedTextRule);
            }
            foreach (var text in page.Value.Descendants(pns + "Text").Where(x => x.Parent?.Name == pns + "Area" && x.Parent.Parent?.Name == pns + "Clip"))
            {
                if (Empty(text)) continue;
                anyClipText = true;
                Use(page.Key, text, pageFonts, uses, ClippingTextRule);
            }
        }

        var embedded = false;
        foreach (var declaration in declarations)
        {
            context.CancellationToken.ThrowIfCancellationRequested();
            uses.TryGetValue(declaration, out var usage);
            embedded |= Payload(declaration, usage);
        }

        if (observations[UsedTextRule.RuleId].Count == 0)
            Add(UsedTextRule, usageScopeComplete && !anyVisibleText ? PreflightOutcome.NotApplicable : PreflightOutcome.Unknown,
                usageScopeComplete && !anyVisibleText ? "NoFilledOrStrokedText" : "TextScopeIncomplete", scope.DocumentEntry, scope.Document,
                "Painted TextObjects=" + anyVisibleText);
        if (observations[ClippingTextRule.RuleId].Count == 0)
            Add(ClippingTextRule, usageScopeComplete && !anyClipText ? PreflightOutcome.NotApplicable : PreflightOutcome.Unknown,
                usageScopeComplete && !anyClipText ? "NoClippingText" : "TextScopeIncomplete", scope.DocumentEntry, scope.Document,
                "Clipping text=" + anyClipText);
        foreach (var rule in new[] { OutlineRule, TablesRule })
            if (observations[rule.RuleId].Count == 0)
                Add(rule, resourceScopeComplete && !embedded ? PreflightOutcome.NotApplicable : PreflightOutcome.Unknown,
                    resourceScopeComplete && !embedded ? "NoEmbeddedFontData" : "ResourceScopeIncomplete", scope.DocumentEntry, scope.Document,
                    "Embedded font payloads=" + embedded);
    }

    // Fill defaults to true and Stroke to false; text that is neither is not drawn by page graphics.
    private static bool Painted(XElement text) =>
        !string.Equals((string?)text.Attribute("Fill"), "false", StringComparison.OrdinalIgnoreCase) ||
        string.Equals((string?)text.Attribute("Stroke"), "true", StringComparison.OrdinalIgnoreCase);

    private static bool Empty(XElement text) =>
        !text.Elements().Any(x => x.Name.LocalName is "TextCode" or "CGTransform" && x.Value.Trim().Length > 0);

    private void ReadResources(string containing, XElement reference, Dictionary<string, FontDeclaration?> fonts)
    {
        context.Xml.Visit();
        var location = reference.Value.Trim();
        if (reference.HasElements || ArchivalPackagePath.Unsafe(location) || !ArchivalPackagePath.TryResolve(containing, location, out var entry))
        { UnknownResources("ResourceTableUnavailable", containing, reference, "Reference=" + location); return; }
        ReadTable(containing, entry, reference, fonts);
    }

    private void ReadTable(string containing, string entry, XElement reference, Dictionary<string, FontDeclaration?> fonts)
    {
        if (!tables.TryGetValue(entry, out var declared))
        {
            tables[entry] = declared = new List<FontDeclaration>();
            var root = context.Xml.Read(entry)?.Root;
            if (root == null || root.Name.LocalName != "Res" || !ArchivalPackagePath.Supported(root.Name.Namespace))
            { UnknownResources("ResourceTableUnavailable", containing, reference, "Target=" + entry); return; }
            ReadFonts(entry, root, declared);
        }
        foreach (var declaration in declared)
            // IDs are document-wide: a second declaration of the same ID makes every use of it ambiguous.
            fonts[declaration.Id] = fonts.TryGetValue(declaration.Id, out var existing) && existing != declaration ? null : declaration;
    }

    private void ReadFonts(string entry, XElement root, List<FontDeclaration> declared)
    {
        var ns = root.Name.Namespace;
        var baseLoc = (string?)root.Attribute("BaseLoc");
        var baseUsable = !root.Attributes().Any(x => x.Name.LocalName == "BaseLoc" && x.Name.Namespace != XNamespace.None) &&
            (string.IsNullOrEmpty(baseLoc) || !ArchivalPackagePath.Unsafe(baseLoc));
        if (root.Elements().Any(x => x.Name.LocalName == "Fonts" && x.Name.Namespace != ns) ||
            root.Elements(ns + "Fonts").SelectMany(x => x.Elements()).Any(x => x.Name != ns + "Font"))
            UnknownResources("UnrecognizedFontDeclaration", entry, root, "Foreign or unknown font declaration");
        foreach (var font in root.Elements(ns + "Fonts").SelectMany(x => x.Elements(ns + "Font")))
        {
            var id = ((string?)font.Attribute("ID"))?.Trim();
            if (string.IsNullOrEmpty(id)) { UnknownResources("AmbiguousFontId", entry, font, "Font without ID"); continue; }
            var declaration = new FontDeclaration(id!, entry, font, baseUsable ? baseLoc : null, baseUsable);
            declared.Add(declaration); // a duplicate ID inside one table becomes ambiguous when mapped
            declarations.Add(declaration);
        }
    }

    private void UnknownResources(string reason, string entry, XElement node, string observed)
    {
        resourceScopeComplete = false; usageScopeComplete = false;
        foreach (var rule in new[] { UsedTextRule, ClippingTextRule, OutlineRule, TablesRule })
            Add(rule, PreflightOutcome.Unknown, reason, entry, node, observed);
    }

    private void Use(string pageEntry, XElement text, Dictionary<string, FontDeclaration?> fonts,
        Dictionary<FontDeclaration, Usage> uses, ArchivalOwnedRuleDefinition rule)
    {
        var id = ((string?)text.Attribute("Font"))?.Trim();
        var objectId = (string?)text.Attribute("ID");
        if (string.IsNullOrEmpty(id) || !fonts.TryGetValue(id!, out var declaration))
        {
            // With every resource table read, a missing ID is a located defect; otherwise it may live in an unread table.
            if (resourceScopeComplete)
                Add(rule, PreflightOutcome.Fail, "UnresolvedFontReference", pageEntry, text,
                    "Object ID=" + objectId + "; Font=" + id + "; no font resource with this ID is declared");
            else Add(rule, PreflightOutcome.Unknown, "FontReferenceUnresolvable", pageEntry, text, "Object ID=" + objectId + "; Font=" + id);
            return;
        }
        if (declaration == null)
        {
            Add(rule, PreflightOutcome.Unknown, "AmbiguousFontId", pageEntry, text, "Object ID=" + objectId + "; Font=" + id);
            return;
        }
        if (!TryReadCodes(text, out var scalars, out var glyphs))
        {
            Add(rule, PreflightOutcome.Unknown, "UnmodeledGlyphMapping", pageEntry, text, "Object ID=" + objectId + "; Font=" + id);
            return;
        }
        if (!uses.TryGetValue(declaration, out var usage)) uses[declaration] = usage = new Usage();
        var target = rule == UsedTextRule ? usage.Painted : usage.Clipping;
        target.Scalars.UnionWith(scalars); target.Glyphs.UnionWith(glyphs);
        target.Objects++;
        target.First ??= pageEntry + " " + ArchivalPackagePath.Selector(text);
    }

    // TextCode content through the font's cmap, except code ranges that CGTransform maps to explicit glyph IDs.
    private static bool TryReadCodes(XElement text, out List<int> scalars, out List<int> glyphs)
    {
        scalars = new List<int>(); glyphs = new List<int>();
        var ns = text.Name.Namespace;
        var codes = text.Elements(ns + "TextCode").ToList();
        var transforms = text.Elements(ns + "CGTransform").ToList();
        if (text.Elements().Any(x => x.Name.LocalName is "TextCode" or "CGTransform" && x.Name.Namespace != ns)) return false;
        var sequences = new List<List<int>>();
        foreach (var code in codes)
        {
            if (code.HasElements || !TryScalars(code.Value, out var values)) return false;
            sequences.Add(values);
        }
        if (transforms.Count == 0) { scalars.AddRange(sequences.SelectMany(x => x)); return true; }
        if (sequences.Count != 1) return false; // positions across several TextCode elements are not modeled
        var sequence = sequences[0];
        var covered = new bool[sequence.Count];
        foreach (var transform in transforms)
        {
            var glyphsNode = transform.Elements().ToList();
            if (glyphsNode.Count != 1 || glyphsNode[0].Name != ns + "Glyphs" || glyphsNode[0].HasElements) return false;
            if (!Number(transform, "CodePosition", -1, out var position) || !Number(transform, "CodeCount", 1, out var count) ||
                !Number(transform, "GlyphCount", 1, out var glyphCount) || position < 0 || count < 0 || glyphCount < 1 ||
                position + (long)count > sequence.Count) return false;
            var tokens = glyphsNode[0].Value.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length != glyphCount) return false;
            foreach (var token in tokens)
            {
                if (!int.TryParse(token, NumberStyles.None, CultureInfo.InvariantCulture, out var gid) || gid > ushort.MaxValue) return false;
                glyphs.Add(gid);
            }
            for (var i = position; i < position + count; i++) covered[i] = true;
        }
        for (var i = 0; i < sequence.Count; i++) if (!covered[i]) scalars.Add(sequence[i]);
        return true;
    }

    private static bool Number(XElement element, string name, int fallback, out int value)
    {
        var text = (string?)element.Attribute(name);
        value = fallback;
        return text == null || int.TryParse(text.Trim(), NumberStyles.None, CultureInfo.InvariantCulture, out value);
    }

    private static bool TryScalars(string text, out List<int> values)
    {
        values = new List<int>(text.Length);
        for (var i = 0; i < text.Length; i++)
        {
            if (!char.IsSurrogate(text[i])) values.Add(text[i]);
            else if (char.IsHighSurrogate(text[i]) && i + 1 < text.Length && char.IsLowSurrogate(text[i + 1])) values.Add(char.ConvertToUtf32(text[i], text[++i]));
            else return false;
        }
        return true;
    }

    // Returns whether the declaration carries an embedded payload in the package.
    private bool Payload(FontDeclaration declaration, Usage? usage)
    {
        var ns = declaration.Element.Name.Namespace;
        var identity = "Font ID=" + declaration.Id + "; FontName=" + (string?)declaration.Element.Attribute("FontName") +
            "; FamilyName=" + (string?)declaration.Element.Attribute("FamilyName");
        var usedBy = usage == null ? Array.Empty<(ArchivalOwnedRuleDefinition, Codes)>() : new[] { (UsedTextRule, usage.Painted), (ClippingTextRule, usage.Clipping) }
            .Where(x => x.Item2.Objects > 0).ToArray();
        void Used(PreflightOutcome outcome, string reason, XElement node, string observed)
        {
            foreach (var (rule, codes) in usedBy)
                Add(rule, outcome, reason, declaration.Entry, node, identity + "; " + observed + "; used by " + codes.Objects + " object(s), first at " + codes.First);
        }
        void Embedded(PreflightOutcome outcome, string reason, XElement node, string observed)
        {
            Add(OutlineRule, outcome, reason, declaration.Entry, node, identity + "; " + observed);
            Add(TablesRule, outcome, reason, declaration.Entry, node, identity + "; " + observed);
        }
        var files = declaration.Element.Elements().Where(x => x.Name.LocalName == "FontFile").ToList();
        if (files.Count == 0)
        {
            if (context.Fonts.Exemption == ArchivalFontExemptionPolicy.NoExemptionsAccepted)
                Used(PreflightOutcome.Fail, "FontPayloadNotEmbedded", declaration.Element, "no FontFile; policy accepts no exemption");
            else
                Used(PreflightOutcome.Unknown, "FontExemptionBasisUnverified", declaration.Element,
                    "no FontFile; GB/T 33190 font exemption edition not fixed");
            return false;
        }
        var file = files[0];
        if (files.Count > 1 || file.Name.Namespace != ns || file.HasElements || !declaration.BaseUsable)
        {
            Used(PreflightOutcome.Unknown, "AmbiguousFontFile", file, "FontFile count=" + files.Count);
            Embedded(PreflightOutcome.Unknown, "AmbiguousFontFile", file, "FontFile count=" + files.Count);
            return true;
        }
        context.Xml.Visit();
        var location = file.Value.Trim();
        if (location.Length == 0 || location.Contains(":") || location.StartsWith("//", StringComparison.Ordinal) ||
            location.StartsWith("\\\\", StringComparison.Ordinal) ||
            !ArchivalPackagePath.Unsafe(location) && !OfdResourceBase.TryResolve(declaration.Entry, declaration.BaseLoc, location, out _))
        {
            Used(PreflightOutcome.Fail, "FontFileOutsidePackage", file, "FontFile is empty, external or escapes the package root");
            return false;
        }
        if (ArchivalPackagePath.Unsafe(location))
        {
            Used(PreflightOutcome.Unknown, "AmbiguousFontFileReference", file, "Ambiguous FontFile reference (redacted)");
            Embedded(PreflightOutcome.Unknown, "AmbiguousFontFileReference", file, "Ambiguous FontFile reference (redacted)");
            return true;
        }
        OfdResourceBase.TryResolve(declaration.Entry, declaration.BaseLoc, location, out var target);
        if (!context.Archive.TryGetBytes(target, out var bytes))
        {
            Used(PreflightOutcome.Fail, "FontFileMissing", file, "FontFile=" + target + " is not in the package");
            return false;
        }
        if (bytes.Length > context.Fonts.MaxTotalBytes - fontBytes) throw new ArchivalBudgetExceededException("FontBytes");
        fontBytes += bytes.Length;

        var report = Inspect(bytes, usage == null ? Array.Empty<Codes>() : new[] { usage.Painted, usage.Clipping });
        var observed = "FontFile=" + target + "; length=" + bytes.Length + "; format=" + report.Format + "; status=" + report.Status + "; reason=" + report.Reason;
        switch (report.Status)
        {
            case OfdFontPayloadStatus.Unsupported:
            case OfdFontPayloadStatus.BudgetExceeded:
                var reason = report.Status == OfdFontPayloadStatus.Unsupported ? "FontFormatUnsupported" : "FontPayloadBudgetExceeded";
                Used(PreflightOutcome.Unknown, reason, file, observed);
                Embedded(PreflightOutcome.Unknown, reason, file, observed);
                return true;
            case OfdFontPayloadStatus.Malformed:
                Used(PreflightOutcome.Fail, "FontPayloadUnparseable", file, observed);
                Add(OutlineRule, PreflightOutcome.Fail, "FontPayloadUnparseable", declaration.Entry, file, identity + "; " + observed);
                if (report.Tables.Count == 0)
                    Add(TablesRule, PreflightOutcome.Fail, "FontPayloadUnparseable", declaration.Entry, file, identity + "; " + observed);
                else Tables(report, declaration, file, identity, observed);
                return true;
        }

        if (report.MissingOutlineTables.Count > 0)
        {
            var missing = "missing=" + string.Join(",", report.MissingOutlineTables).Trim();
            Add(OutlineRule, PreflightOutcome.Fail, "OutlineTableMissing", declaration.Entry, file, identity + "; " + observed + "; " + missing);
            Used(PreflightOutcome.Fail, "OutlineTableMissing", file, observed + "; " + missing);
        }
        else if (report.UnreadableGlyphIds.Count > 0)
            Add(OutlineRule, PreflightOutcome.Fail, "GlyphOutlineUnreadable", declaration.Entry, file,
                identity + "; " + observed + "; GID " + string.Join(" ", report.UnreadableGlyphIds.Take(8)));
        else if (report.UninterpretedGlyphIds.Count > 0)
            Add(OutlineRule, PreflightOutcome.Unknown, "GlyphOutlineUninterpreted", declaration.Entry, file,
                identity + "; " + observed + "; GID " + string.Join(" ", report.UninterpretedGlyphIds.Take(8)) + " use an outline feature not interpreted");
        else
            Add(OutlineRule, PreflightOutcome.Pass, "OutlineTablesRetained", declaration.Entry, file,
                identity + "; " + observed + "; checked glyphs of used text");
        Tables(report, declaration, file, identity, observed);

        if (report.MissingOutlineTables.Count == 0)
            foreach (var (rule, codes) in usedBy)
            {
                var used = identity + "; " + observed + "; used by " + codes.Objects + " object(s), first at " + codes.First;
                // Outline results are per glyph: with both categories present, judge each category by its own glyphs only.
                var own = usedBy.Length == 1 ? report : Inspect(bytes, new[] { codes });
                if (own.Status != OfdFontPayloadStatus.Parsed)
                {
                    Add(rule, PreflightOutcome.Unknown, own.Status == OfdFontPayloadStatus.BudgetExceeded ? "FontPayloadBudgetExceeded" : "FontFormatUnsupported",
                        declaration.Entry, file, used + "; category status=" + own.Status + "; reason=" + own.Reason);
                    continue;
                }
                var missingScalars = codes.Scalars.Where(own.MissingScalars.Contains).OrderBy(x => x).ToList();
                var missingGlyphs = codes.Glyphs.Where(own.MissingGlyphIds.Contains).OrderBy(x => x).ToList();
                if (missingScalars.Count > 0 || missingGlyphs.Count > 0 || own.UnreadableGlyphIds.Count > 0)
                    Add(rule, PreflightOutcome.Fail, "UsedGlyphMissing", declaration.Entry, file, used + "; missing " +
                        string.Join(" ", missingScalars.Take(8).Select(x => "U+" + x.ToString("X4", CultureInfo.InvariantCulture))
                            .Concat(missingGlyphs.Take(8).Select(x => "GID " + x))
                            .Concat(own.UnreadableGlyphIds.Take(8).Select(x => "unreadable GID " + x))));
                else if (own.UninterpretedGlyphIds.Count > 0)
                    Add(rule, PreflightOutcome.Unknown, "GlyphOutlineUninterpreted", declaration.Entry, file, used + "; GID " +
                        string.Join(" ", own.UninterpretedGlyphIds.Take(8)) + " use an outline feature not interpreted");
                else
                    Add(rule, PreflightOutcome.Pass, "UsedGlyphsCovered", declaration.Entry, file, used + "; scalars=" + codes.Scalars.Count +
                        "; glyphIds=" + codes.Glyphs.Count + "; fsType=" + Permissions(report) + " (recorded only; licensing not assessed)");
            }
        return true;
    }

    private OfdFontPayloadReport Inspect(byte[] bytes, IEnumerable<Codes> categories)
    {
        var request = new OfdFontPayloadRequest { MaxPayloadBytes = context.Fonts.MaxPayloadBytes };
        foreach (var codes in categories)
        {
            foreach (var value in codes.Scalars) request.UnicodeScalars.Add(value);
            foreach (var value in codes.Glyphs) request.GlyphIds.Add(value);
        }
        return OfdFontPayloadInspector.Inspect(bytes, request, context.CancellationToken);
    }

    private void Tables(OfdFontPayloadReport report, FontDeclaration declaration, XElement file, string identity, string observed)
    {
        if (report.MissingRequiredTables.Count > 0)
            Add(TablesRule, PreflightOutcome.Fail, "RequiredTableMissing", declaration.Entry, file,
                identity + "; " + observed + "; missing=" + string.Join(",", report.MissingRequiredTables));
        else
            Add(TablesRule, PreflightOutcome.Pass, "RequiredTablesRetained", declaration.Entry, file,
                identity + "; " + observed + "; program tables=" + (report.ProgramTables.Count == 0 ? "none" : string.Join(",", report.ProgramTables).Trim()) +
                " (necessity not judged)");
    }

    private static string Permissions(OfdFontPayloadReport report) =>
        report.EmbeddingPermissions is { } bits ? "0x" + bits.ToString("X4", CultureInfo.InvariantCulture) : "absent";

    private sealed class FontDeclaration
    {
        internal FontDeclaration(string id, string entry, XElement element, string? baseLoc, bool baseUsable)
        { Id = id; Entry = entry; Element = element; BaseLoc = baseLoc; BaseUsable = baseUsable; }
        internal string Id { get; }
        internal string Entry { get; }
        internal XElement Element { get; }
        internal string? BaseLoc { get; }
        internal bool BaseUsable { get; }
    }

    private sealed class Codes
    {
        internal HashSet<int> Scalars { get; } = new HashSet<int>();
        internal HashSet<int> Glyphs { get; } = new HashSet<int>();
        internal int Objects { get; set; }
        internal string? First { get; set; }
    }

    private sealed class Usage
    {
        internal Codes Painted { get; } = new Codes();
        internal Codes Clipping { get; } = new Codes();
    }
}
