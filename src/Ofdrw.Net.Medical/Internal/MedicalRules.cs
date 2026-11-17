using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Ofdrw.Net.Archival;
using Ofdrw.Net.Packaging.Preflight;

namespace Ofdrw.Net.Medical.Internal;

internal sealed class MedicalRuleContext
{
    internal MedicalRuleContext(MedicalPackage package, MedicalStructure structure, MedicalAssessmentOptions options,
        MedicalMetadataMapping? mapping, ArchivalAssessmentReport? archival)
    { Package = package; Structure = structure; Options = options; Mapping = mapping; Archival = archival; }
    internal MedicalPackage Package { get; }
    internal MedicalStructure Structure { get; }
    internal MedicalAssessmentOptions Options { get; }
    internal MedicalMetadataMapping? Mapping { get; }
    internal ArchivalAssessmentReport? Archival { get; }
}

/// <summary>A package-owned executable rule. Every rule returns at least one observation per mapped clause.</summary>
internal sealed class MedicalRule
{
    internal MedicalRule(string id, string[] clauses, string description, Func<MedicalRuleContext, IEnumerable<MedicalRuleObservation>> check)
    { Id = id; Clauses = clauses; Description = description; Check = check; }
    internal string Id { get; }
    internal string[] Clauses { get; }
    internal string Description { get; }
    internal Func<MedicalRuleContext, IEnumerable<MedicalRuleObservation>> Check { get; }
}

internal static class MedicalRules
{
    internal const string Sm2Oid = "1.2.156.10197.1.501";
    internal const string Sm3Oid = "1.2.156.10197.1.401";
    private const string Package = "<package>";

    private static readonly Regex PageFolder = new Regex(@"^Doc_0/Pages/Page_\d+/", RegexOptions.CultureInvariant);

    internal static IReadOnlyList<MedicalRule> All { get; } = new[]
    {
        new MedicalRule("MH-001", new[] { "7.1(a)" }, "Root DocType is the ASCII value OFD-H", DocType),
        new MedicalRule("MH-002", new[] { "6.1.1", "6.1.2(a)" }, "Single document at Doc_0 inside the OFD container", SingleDocument),
        new MedicalRule("MH-003", new[] { "6.2(c)", "6.2(d)" }, "Namespace matches the declared GB/T 33190 edition; ofd prefix", Namespaces),
        new MedicalRule("MH-004", new[] { "6.1.3(b)" }, "Compressible entries use Deflate", Deflate),
        new MedicalRule("MH-005", new[] { "6.1.3(a)" }, "Physical entry order is linearized", Linearized),
        new MedicalRule("MH-006", new[] { "6.2(a)", "7.1(b)", "7.1(d)", "7.1(e)", "7.1(f)" }, "Canonical folders and in-container resources", Folders),
        new MedicalRule("MH-007", new[] { "6.2(b)", "7.2(a)" }, "Fonts and color spaces in the prescribed resource tables", ResourceTables),
        new MedicalRule("MH-008", new[] { "7.2(b)" }, "Fonts embedded (GB/T 42133 AC-017/AC-018 evidence)", Fonts),
        new MedicalRule("MH-009", new[] { "7.2(c)", "7.2(d)" }, "Image payloads present and in a permitted format", Images),
        new MedicalRule("MH-010", new[] { "7.2(e)", "7.2(f)", "7.2(g)", "7.4(c)" }, "Audio/video formats and embedded object size limit", MediaAndSize),
        new MedicalRule("MH-011", new[] { "7.1(c)", "7.3(a)", "7.3(b)" }, "Required description and identification metadata", Metadata),
        new MedicalRule("MH-012", new[] { "7.5(a)", "7.5(b)" }, "Machine-readable payloads are UTF-8 XML or JSON", MachineReadable),
        new MedicalRule("MH-013", new[] { "6.3(e)", "6.3(g)", "8(c)", "7.3(d)" }, "Latest signature covers all content except annotation and signature files", SignatureCoverage),
        new MedicalRule("MH-014", new[] { "6.3(b)", "8(d)" }, "Signatures use SM2/SM3 with GB/T 33560 identifiers", Algorithms),
        new MedicalRule("MH-015", new[] { "7.3(g)" }, "Permissions do not prohibit annotation or signing", Permissions),
        new MedicalRule("MH-016", new[] { "8(b)" }, "Encryption declared only for temporary use", Encryption),
    };

    private static MedicalRuleObservation O(string rule, string clause, PreflightOutcome outcome, string reason, string entry, string selector, string observed, string expected) =>
        new MedicalRuleObservation(rule, clause, outcome, reason, new MedicalEvidence(entry, selector, observed, expected));

    private static string Sample(IEnumerable<string> values)
    {
        var list = values.Take(4).ToList();
        return string.Join(", ", list.Take(3)) + (list.Count > 3 ? ", ..." : string.Empty);
    }

    // MH-001 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> DocType(MedicalRuleContext context)
    {
        const string rule = "MH-001", clause = "7.1(a)", expected = "DocType=\"OFD-H\" (ASCII hyphen-minus)";
        var root = context.Structure.Root?.Root;
        if (root == null) { yield return O(rule, clause, PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", Package, "OFD.xml absent or not well-formed", expected); yield break; }
        var value = (string?)root.Attribute("DocType");
        if (root.Attributes().Any(x => x.Name.LocalName == "DocType" && x.Name.Namespace != XNamespace.None))
            yield return O(rule, clause, PreflightOutcome.Unknown, "AmbiguousDocType", "OFD.xml", "/OFD/@DocType", "Qualified DocType attribute present", expected);
        else if (value == null) yield return O(rule, clause, PreflightOutcome.Fail, "DocTypeMissing", "OFD.xml", "/OFD/@DocType", "absent", expected);
        else if (MedicalDocTypes.IsMedical(value)) yield return O(rule, clause, PreflightOutcome.Pass, "MedicalDocType", "OFD.xml", "/OFD/@DocType", value, expected);
        else if (MedicalDocTypes.IsTypesetVariant(value))
            yield return O(rule, clause, PreflightOutcome.Fail, "TypesetDashVariant", "OFD.xml", "/OFD/@DocType", "OFD U+" + ((int)value[3]).ToString("X4") + " H", expected);
        else yield return O(rule, clause, PreflightOutcome.Fail, "NonMedicalDocType", "OFD.xml", "/OFD/@DocType", value, expected);
    }

    // MH-002 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> SingleDocument(MedicalRuleContext context)
    {
        const string rule = "MH-002";
        var s = context.Structure;
        if (s.Root?.Root == null)
        {
            yield return O(rule, "6.1.1", PreflightOutcome.Fail, "RootEntryMissing", "OFD.xml", Package, "OFD.xml absent or not well-formed", "OFD.xml main entry");
            yield return O(rule, "6.1.2(a)", PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", Package, "OFD.xml unavailable", "One DocBody at Doc_0");
            yield break;
        }
        if (s.Namespace == null || !Supported(s.Namespace.NamespaceName) || s.Root.Root.Name.LocalName != "OFD")
        {
            yield return O(rule, "6.1.1", PreflightOutcome.Fail, "UnrecognizedRoot", "OFD.xml", "/", s.Root.Root.Name.ToString(), "{OFD namespace}OFD");
            yield return O(rule, "6.1.2(a)", PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", "/", "Unrecognized root", "One DocBody at Doc_0");
            yield break;
        }
        var count = s.DocBodies.Count;
        if (count == 0)
        {
            yield return O(rule, "6.1.1", PreflightOutcome.Fail, "NoDocument", "OFD.xml", "/OFD/DocBody", "0 DocBody", "Exactly one DocBody");
            yield return O(rule, "6.1.2(a)", PreflightOutcome.Fail, "NoDocument", "OFD.xml", "/OFD/DocBody", "0 DocBody", "Exactly one DocBody");
            yield break;
        }
        yield return s.Document?.Root?.Name.LocalName == "Document"
            ? O(rule, "6.1.1", PreflightOutcome.Pass, "ContainerAndDocument", "OFD.xml", "/OFD/DocBody/DocRoot", s.DocumentPath!, "Document entry inside the container")
            : O(rule, "6.1.1", PreflightOutcome.Fail, "DocumentMissing", "OFD.xml", "/OFD/DocBody/DocRoot", s.DocumentPath ?? "unresolvable", "Document entry inside the container");
        if (count > 1)
            yield return O(rule, "6.1.2(a)", PreflightOutcome.Fail, "MultipleDocuments", "OFD.xml", "/OFD/DocBody", count + " DocBody", "Exactly one DocBody");
        else if (s.DocumentPath != "Doc_0/Document.xml")
            yield return O(rule, "6.1.2(a)", PreflightOutcome.Fail, "NonCanonicalDocRoot", "OFD.xml", "/OFD/DocBody/DocRoot", s.DocumentPath ?? "unresolvable", "Doc_0/Document.xml");
        else
            yield return O(rule, "6.1.2(a)", PreflightOutcome.Pass, "SingleDocumentDoc0", "OFD.xml", "/OFD/DocBody/DocRoot", s.DocumentPath, "Doc_0/Document.xml");
    }

    private static bool Supported(string ns) => ns == MedicalPackage.Namespace2016 || ns == MedicalPackage.ShortNamespace;

    // MH-003 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Namespaces(MedicalRuleContext context)
    {
        const string rule = "MH-003";
        var root = context.Structure.Root?.Root;
        if (root == null)
        {
            yield return O(rule, "6.2(c)", PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", Package, "OFD.xml unavailable", "GB/T 33190 namespace");
            yield return O(rule, "6.2(d)", PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", Package, "OFD.xml unavailable", "ofd prefix");
            yield break;
        }
        var version = (string?)root.Attribute("Version") ?? string.Empty;
        string? expected = version.StartsWith("1.", StringComparison.Ordinal) ? MedicalPackage.Namespace2016 : null;
        var wrongNamespace = new List<string>(); var undeclared = new List<string>(); var malformed = new List<string>();
        var unprefixed = new List<string>(); var qualified = new List<string>();
        foreach (var name in context.Package.EntryNames.Where(x => x.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).OrderBy(x => x, StringComparer.Ordinal))
        {
            context.Package.Visit();
            var document = context.Package.Xml(name)?.Root;
            if (document == null) { malformed.Add(name); continue; }
            var ns = document.Name.NamespaceName;
            if (!ns.StartsWith(MedicalPackage.ShortNamespace, StringComparison.Ordinal)) continue; // custom tag payloads etc.
            if (expected != null && ns != expected) wrongNamespace.Add(name);
            var declared = document.Attributes().Any(a => a.IsNamespaceDeclaration && a.Name.Namespace == XNamespace.Xmlns && a.Name.LocalName == "ofd" && a.Value == ns);
            if (!declared) undeclared.Add(name);
            var lexical = context.Package.Lexical(name, ns);
            if (lexical.UnprefixedOfdElements > 0) unprefixed.Add(name);
            if (lexical.QualifiedOfdAttributes > 0) qualified.Add(name);
        }
        if (expected == null)
            yield return O(rule, "6.2(c)", PreflightOutcome.Unknown, "EditionNotPinned", "OFD.xml", "/OFD/@Version", version, "Version 1.x with http://www.ofdspec.org/2016");
        else if (wrongNamespace.Count > 0)
            yield return O(rule, "6.2(c)", PreflightOutcome.Fail, "NamespaceNotStandard", wrongNamespace[0], "/*", wrongNamespace.Count + " entries: " + Sample(wrongNamespace), expected);
        else if (undeclared.Count > 0)
            yield return O(rule, "6.2(c)", PreflightOutcome.Fail, "OfdPrefixNotDeclared", undeclared[0], "/*/@xmlns:ofd", undeclared.Count + " entries: " + Sample(undeclared), "Root declares xmlns:ofd");
        else if (malformed.Count > 0)
            yield return O(rule, "6.2(c)", PreflightOutcome.Unknown, "XmlNotWellFormed", malformed[0], Package, malformed.Count + " entries: " + Sample(malformed), "Well-formed XML");
        else
            yield return O(rule, "6.2(c)", PreflightOutcome.Pass, "StandardNamespaceDeclared", "OFD.xml", "/OFD/@Version", version, expected);
        if (unprefixed.Count > 0)
            yield return O(rule, "6.2(d)", PreflightOutcome.Fail, "ElementsUnprefixed", unprefixed[0], Package, unprefixed.Count + " entries: " + Sample(unprefixed), "Elements use the ofd prefix");
        if (qualified.Count > 0)
            yield return O(rule, "6.2(d)", PreflightOutcome.Fail, "AttributesQualified", qualified[0], Package, qualified.Count + " entries: " + Sample(qualified), "Unqualified attributes");
        if (unprefixed.Count == 0 && qualified.Count == 0)
            yield return malformed.Count > 0
                ? O(rule, "6.2(d)", PreflightOutcome.Unknown, "XmlNotWellFormed", malformed[0], Package, malformed.Count + " entries", "Well-formed XML")
                : O(rule, "6.2(d)", PreflightOutcome.Pass, "PrefixedElements", "OFD.xml", Package, "All OFD elements prefixed, attributes unqualified", "ofd prefix");
    }

    // MH-004 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Deflate(MedicalRuleContext context)
    {
        const string rule = "MH-004", clause = "6.1.3(b)";
        var entries = context.Package.Directory;
        var encrypted = entries.Where(x => x.Encrypted).Select(x => x.Name).ToList();
        var unsupported = entries.Where(x => x.Method != 0 && x.Method != 8).Select(x => x.Name + " (method " + x.Method + ")").ToList();
        var stored = entries.Where(x => x.Method == 0 && x.UncompressedSize > 0 && x.Name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)).Select(x => x.Name).ToList();
        if (unsupported.Count > 0)
            yield return O(rule, clause, PreflightOutcome.Fail, "UnsupportedCompressionMethod", Package, "central directory", Sample(unsupported), "Deflate (8) or stored (0)");
        else if (stored.Count > 0)
            yield return O(rule, clause, PreflightOutcome.Fail, "CompressibleEntryStored", stored[0], "central directory", stored.Count + " XML entries stored: " + Sample(stored), "XML entries use Deflate");
        else if (encrypted.Count > 0)
            yield return O(rule, clause, PreflightOutcome.Unknown, "ZipEncryptedEntries", encrypted[0], "central directory", encrypted.Count + " encrypted entries", "Readable ZIP entries");
        else
            yield return O(rule, clause, PreflightOutcome.Pass, "CompressibleEntriesDeflated", Package, "central directory", entries.Count + " entries", "XML entries use Deflate");
    }

    // MH-005 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Linearized(MedicalRuleContext context)
    {
        const string rule = "MH-005", clause = "6.1.3(a)", expected = "OFD.xml first, document before pages, pages in order";
        var order = context.Package.PhysicalOrder.Select(x => x.Name).ToList();
        var s = context.Structure;
        if (s.DocumentPath == null || s.PagePaths.Count == 0)
        {
            yield return O(rule, clause, PreflightOutcome.Unknown, "StructureUnavailable", Package, "local headers", "Document or pages unresolved", expected);
            yield break;
        }
        int Index(string name) => order.FindIndex(x => string.Equals(x, name, StringComparison.OrdinalIgnoreCase));
        var pages = s.PagePaths.Select(Index).ToList();
        var linear = order.Count > 0 && order[0] == "OFD.xml" && Index(s.DocumentPath) >= 0 && pages.All(x => x > Index(s.DocumentPath)) &&
                     pages.Zip(pages.Skip(1), (a, b) => a < b).All(x => x);
        yield return linear
            ? O(rule, clause, PreflightOutcome.Pass, "EntryOrderLinearized", order[0], "local headers", "first: " + Sample(order), expected)
            : O(rule, clause, PreflightOutcome.Fail, "EntryOrderNotLinearized", order.FirstOrDefault() ?? Package, "local headers", "first: " + Sample(order), expected);
    }

    // MH-006 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Folders(MedicalRuleContext context)
    {
        const string rule = "MH-006";
        var s = context.Structure;
        if (s.DocumentPath == null)
        {
            foreach (var clause in new[] { "6.2(a)", "7.1(b)", "7.1(d)", "7.1(e)", "7.1(f)" })
                yield return O(rule, clause, PreflightOutcome.Unknown, "StructureUnavailable", "OFD.xml", Package, "Document unresolved", "Canonical folders");
            yield break;
        }
        var unsafeLocations = s.Locations.Where(x => x.Resolved == null).ToList();
        var bad = s.Locations.Where(x => x.Resolved != null && !Canonical(x)).ToList();
        yield return bad.Count > 0
            ? O(rule, "6.2(a)", PreflightOutcome.Fail, "NonCanonicalLocation", bad[0].DeclaredIn, bad[0].Selector, bad.Count + " locations: " + Sample(bad.Select(x => x.Resolved!)), "Doc_0 figure 3 folders")
            : unsafeLocations.Count > 0
                ? O(rule, "6.2(a)", PreflightOutcome.Unknown, "UnresolvableLocation", unsafeLocations[0].DeclaredIn, unsafeLocations[0].Selector, Sample(unsafeLocations.Select(x => x.Declared)), "Package-relative location")
                : O(rule, "6.2(a)", PreflightOutcome.Pass, "CanonicalFolders", s.DocumentPath, Package, s.Locations.Count + " locations", "Doc_0 figure 3 folders");

        var payloads = s.Locations.Where(x => x.Kind == MedicalLocationKind.ResourcePayload).ToList();
        var missing = payloads.Where(x => x.Resolved != null && !context.Package.Exists(x.Resolved)).ToList();
        var outside = payloads.Where(x => x.Resolved != null && context.Package.Exists(x.Resolved) && !Canonical(x)).ToList();
        if (payloads.Count == 0)
            yield return O(rule, "7.1(b)", PreflightOutcome.NotApplicable, "NoResourcePayloads", s.DocumentPath, Package, "No FontFile or MediaFile", "Resources inside the container");
        else if (missing.Count > 0)
            yield return O(rule, "7.1(b)", PreflightOutcome.Fail, "ResourceNotInContainer", missing[0].DeclaredIn, missing[0].Selector, missing.Count + " missing: " + Sample(missing.Select(x => x.Resolved!)), "Payload inside the container");
        else if (outside.Count > 0)
            yield return O(rule, "7.1(b)", PreflightOutcome.Fail, "ResourceOutsideResFolder", outside[0].DeclaredIn, outside[0].Selector, Sample(outside.Select(x => x.Resolved!)), "Doc_0/Res or Page_N/Res");
        else if (payloads.Any(x => x.Resolved == null))
            yield return O(rule, "7.1(b)", PreflightOutcome.Unknown, "UnresolvableLocation", payloads.First(x => x.Resolved == null).DeclaredIn, Package, "Unsafe payload location", "Package-relative location");
        else
            yield return O(rule, "7.1(b)", PreflightOutcome.Pass, "ResourcesInResFolders", s.DocumentPath, Package, payloads.Count + " payloads", "Resources inside Res folders");

        foreach (var (clause, kinds, folder) in new[]
        {
            ("7.1(d)", new[] { MedicalLocationKind.CustomTags, MedicalLocationKind.CustomTag }, "Doc_0/Tags/"),
            ("7.1(e)", new[] { MedicalLocationKind.Attachments, MedicalLocationKind.Attachment }, "Doc_0/Attachs/"),
            ("7.1(f)", new[] { MedicalLocationKind.Annotations, MedicalLocationKind.PageAnnotation }, "Doc_0/Annots/")
        })
        {
            var items = s.Locations.Where(x => kinds.Contains(x.Kind)).ToList();
            var wrong = items.Where(x => x.Resolved != null && !x.Resolved.StartsWith(folder, StringComparison.Ordinal)).ToList();
            if (items.Count == 0)
                yield return O(rule, clause, PreflightOutcome.NotApplicable, "NotDeclared", s.DocumentPath, Package, "None declared", folder);
            else if (wrong.Count > 0)
                yield return O(rule, clause, PreflightOutcome.Fail, "OutsideFolder", wrong[0].DeclaredIn, wrong[0].Selector, Sample(wrong.Select(x => x.Resolved!)), folder);
            else if (items.Any(x => x.Resolved == null))
                yield return O(rule, clause, PreflightOutcome.Unknown, "UnresolvableLocation", items.First(x => x.Resolved == null).DeclaredIn, Package, "Unsafe location", folder);
            else
                yield return O(rule, clause, PreflightOutcome.Pass, "InFolder", items[0].DeclaredIn, items[0].Selector, items.Count + " locations", folder);
        }
    }

    private static bool Canonical(MedicalLocation location)
    {
        var path = location.Resolved!;
        bool Direct(string folder) => path.StartsWith(folder, StringComparison.Ordinal) && path.IndexOf('/', folder.Length) < 0;
        switch (location.Kind)
        {
            case MedicalLocationKind.Document: return path == "Doc_0/Document.xml";
            case MedicalLocationKind.PublicRes:
            case MedicalLocationKind.DocumentRes:
            case MedicalLocationKind.Extensions: return Direct("Doc_0/");
            case MedicalLocationKind.Page:
            case MedicalLocationKind.PageRes: return PageFolder.IsMatch(path) && path.IndexOf('/', PageFolder.Match(path).Length) < 0;
            case MedicalLocationKind.Template: return Direct("Doc_0/Temps/");
            case MedicalLocationKind.Signatures: return Direct("Doc_0/Signs/");
            case MedicalLocationKind.Signature: return Regex.IsMatch(path, @"^Doc_0/Signs/Sign_\d+/[^/]+$");
            case MedicalLocationKind.Attachments: return Direct("Doc_0/Attachs/");
            case MedicalLocationKind.Attachment: return path.StartsWith("Doc_0/Attachs/", StringComparison.Ordinal);
            case MedicalLocationKind.CustomTags:
            case MedicalLocationKind.CustomTag: return path.StartsWith("Doc_0/Tags/", StringComparison.Ordinal);
            case MedicalLocationKind.Annotations:
            case MedicalLocationKind.PageAnnotation: return path.StartsWith("Doc_0/Annots/", StringComparison.Ordinal);
            case MedicalLocationKind.ResourcePayload:
                return path.StartsWith("Doc_0/Res/", StringComparison.Ordinal) || Regex.IsMatch(path, @"^Doc_0/Pages/Page_\d+/Res/");
            default: return false;
        }
    }

    // MH-007 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> ResourceTables(MedicalRuleContext context)
    {
        const string rule = "MH-007";
        var s = context.Structure;
        if (s.DocumentPath == null || s.Document == null)
        {
            yield return O(rule, "6.2(b)", PreflightOutcome.Unknown, "StructureUnavailable", "OFD.xml", Package, "Document unavailable", "Fonts and color spaces in PublicRes");
            yield return O(rule, "7.2(a)", PreflightOutcome.Unknown, "StructureUnavailable", "OFD.xml", Package, "Document unavailable", "Fonts in DocumentRes");
            yield break;
        }
        List<string> In(MedicalResourceScope scope, string local) =>
            s.Tables.Where(t => t.Scope == scope && t.Root.Descendants(t.Root.Name.Namespace + local).Any()).Select(t => t.Path).ToList();
        var fontsPage = In(MedicalResourceScope.Page, "Font"); var fontsDocument = In(MedicalResourceScope.Document, "Font");
        var fontsPublic = In(MedicalResourceScope.Public, "Font");
        var colorsPage = In(MedicalResourceScope.Page, "ColorSpace"); var colorsDocument = In(MedicalResourceScope.Document, "ColorSpace");
        var colorsPublic = In(MedicalResourceScope.Public, "ColorSpace");
        var unreadable = s.Unreadable.Count > 0;

        if (fontsPage.Count + colorsPage.Count > 0)
            yield return O(rule, "6.2(b)", PreflightOutcome.Fail, "DefinedInPageRes", fontsPage.Concat(colorsPage).First(), "Font|ColorSpace", Sample(fontsPage.Concat(colorsPage)), "PublicRes");
        else if (colorsDocument.Count > 0)
            yield return O(rule, "6.2(b)", PreflightOutcome.Fail, "ColorSpaceOutsidePublicRes", colorsDocument[0], "ColorSpace", Sample(colorsDocument), "PublicRes");
        else if (fontsDocument.Count > 0)
            yield return O(rule, "6.2(b)", PreflightOutcome.Unknown, "ResourceTableConflict", fontsDocument[0], "Font", "Fonts in DocumentRes (7.2(a))", "PublicRes (6.2 table)");
        else if (unreadable)
            yield return O(rule, "6.2(b)", PreflightOutcome.Unknown, "ResourceScopeIncomplete", s.Unreadable[0], Package, "Unreadable entry", "Complete resource scope");
        else if (fontsPublic.Count + colorsPublic.Count == 0)
            yield return O(rule, "6.2(b)", PreflightOutcome.NotApplicable, "NoFontsOrColorSpaces", s.DocumentPath, Package, "None defined", "PublicRes");
        else
            yield return O(rule, "6.2(b)", PreflightOutcome.Pass, "DefinedInPublicRes", fontsPublic.Concat(colorsPublic).First(), "Font|ColorSpace", Sample(fontsPublic.Concat(colorsPublic)), "PublicRes");

        if (fontsPage.Count > 0)
            yield return O(rule, "7.2(a)", PreflightOutcome.Fail, "FontInPageRes", fontsPage[0], "Font", Sample(fontsPage), "DocumentRes");
        else if (fontsPublic.Count > 0)
            yield return O(rule, "7.2(a)", PreflightOutcome.Unknown, "ResourceTableConflict", fontsPublic[0], "Font", "Fonts in PublicRes (6.2 table)", "DocumentRes (7.2(a))");
        else if (unreadable)
            yield return O(rule, "7.2(a)", PreflightOutcome.Unknown, "ResourceScopeIncomplete", s.Unreadable[0], Package, "Unreadable entry", "Complete resource scope");
        else if (fontsDocument.Count == 0)
            yield return O(rule, "7.2(a)", PreflightOutcome.NotApplicable, "NoFonts", s.DocumentPath, Package, "No fonts", "DocumentRes");
        else
            yield return O(rule, "7.2(a)", PreflightOutcome.Pass, "FontsInDocumentRes", fontsDocument[0], "Font", Sample(fontsDocument), "DocumentRes");
    }

    // MH-008 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Fonts(MedicalRuleContext context)
    {
        const string rule = "MH-008", clause = "7.2(b)";
        var archival = context.Archival;
        if (archival == null || archival.ExecutionState != ArchivalExecutionState.Complete)
        {
            yield return O(rule, clause, PreflightOutcome.Unknown, "ArchivalAssessmentIncomplete", Package, Package, archival?.ExecutionReason ?? "not run", "Complete AC-017/AC-018 evidence");
            yield break;
        }
        var observations = new[] { "6.2.6(b)(1)", "6.2.6(b)(2)" }.SelectMany(x => archival.GetClause(x).Observations)
            .Where(x => x.RuleId == "AC-017" || x.RuleId == "AC-018").ToList();
        if (observations.Count == 0)
        {
            yield return O(rule, clause, PreflightOutcome.Unknown, "NoFontEvidence", Package, Package, "AC-017/AC-018 produced no observation", "Embedded fonts");
            yield break;
        }
        foreach (var item in observations)
            yield return O(rule, clause, item.Outcome, item.RuleId + ":" + item.ReasonCode, item.Evidence.Entry, item.Evidence.Selector, item.Evidence.Observed, item.Evidence.Expected);
    }

    // MH-009 -------------------------------------------------------------------------------------------------
    internal static string? ImageFormat(byte[] head)
    {
        bool Starts(params byte[] prefix) => head.Length >= prefix.Length && prefix.Select((b, i) => head[i] == b).All(x => x);
        if (Starts(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A)) return "PNG";
        if (Starts(0xFF, 0xD8, 0xFF)) return "JPEG";
        if (Starts(0x00, 0x00, 0x00, 0x0C, 0x6A, 0x50, 0x20, 0x20, 0x0D, 0x0A, 0x87, 0x0A) || Starts(0xFF, 0x4F, 0xFF, 0x51)) return "JPEG2000";
        if (Starts(0x49, 0x49, 0x2A, 0x00) || Starts(0x4D, 0x4D, 0x00, 0x2A)) return "TIFF";
        if (Starts(0x42, 0x4D)) return "BMP";
        if (Starts(0x97, 0x4A, 0x42, 0x32, 0x0D, 0x0A, 0x1A, 0x0A)) return "JBIG2";
        return null;
    }

    internal static IEnumerable<MedicalRuleObservation> Images(MedicalRuleContext context)
    {
        const string rule = "MH-009";
        var images = context.Structure.Media.Where(x => string.Equals(x.Type, "Image", StringComparison.OrdinalIgnoreCase)).ToList();
        if (images.Count == 0)
        {
            var reason = context.Structure.Unreadable.Count > 0 ? PreflightOutcome.Unknown : PreflightOutcome.NotApplicable;
            yield return O(rule, "7.2(c)", reason, reason == PreflightOutcome.Unknown ? "ResourceScopeIncomplete" : "NoImages", Package, Package, "No image resources", "Image payloads in the file");
            yield return O(rule, "7.2(d)", reason, reason == PreflightOutcome.Unknown ? "ResourceScopeIncomplete" : "NoImages", Package, Package, "No image resources", "Permitted image formats");
            yield break;
        }
        var missing = images.Where(x => x.Resolved != null && !context.Package.Exists(x.Resolved)).ToList();
        var unresolved = images.Where(x => x.Resolved == null).ToList();
        yield return missing.Count > 0
            ? O(rule, "7.2(c)", PreflightOutcome.Fail, "ImageNotInFile", missing[0].Table.Path, missing[0].Selector, Sample(missing.Select(x => x.Resolved!)), "Image payload in the file")
            : unresolved.Count > 0
                ? O(rule, "7.2(c)", PreflightOutcome.Unknown, "UnresolvableLocation", unresolved[0].Table.Path, unresolved[0].Selector, unresolved[0].Declared, "Package-relative MediaFile")
                : O(rule, "7.2(c)", PreflightOutcome.Pass, "ImagesInFile", images[0].Table.Path, Package, images.Count + " images", "Image payloads in the file");
        var formats = new List<string>(); var rejected = new List<string>();
        foreach (var image in images.Where(x => x.Resolved != null && context.Package.Exists(x.Resolved)))
        {
            context.Package.Visit();
            var format = ImageFormat(context.Package.Head(image.Resolved!, 16) ?? Array.Empty<byte>());
            if (format == null) rejected.Add(image.Resolved!); else formats.Add(format);
        }
        yield return rejected.Count > 0
            ? O(rule, "7.2(d)", PreflightOutcome.Fail, "ImageFormatNotPermitted", rejected[0], "payload signature", Sample(rejected), "BMP, JPEG, JPEG2000, TIFF, PNG or JBIG2")
            : formats.Count == 0
                ? O(rule, "7.2(d)", PreflightOutcome.Unknown, "NoReadableImage", Package, Package, "No readable image payload", "Permitted image formats")
                : O(rule, "7.2(d)", PreflightOutcome.Pass, "PermittedImageFormats", Package, "payload signature", string.Join(", ", formats.Distinct()), "BMP, JPEG, JPEG2000, TIFF, PNG or JBIG2");
    }

    // MH-010 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> MediaAndSize(MedicalRuleContext context)
    {
        const string rule = "MH-010";
        var media = context.Structure.Media.Where(x => x.Resolved != null && context.Package.Exists(x.Resolved)).ToList();
        foreach (var (clause, type, name) in new[] { ("7.2(f)", "Audio", "MP3"), ("7.2(g)", "Video", "MP4") })
        {
            var items = media.Where(x => string.Equals(x.Type, type, StringComparison.OrdinalIgnoreCase)).ToList();
            if (items.Count == 0)
            {
                yield return O(rule, clause, PreflightOutcome.NotApplicable, "No" + type, Package, Package, "No embedded " + type.ToLowerInvariant(), name + " or another common format");
                continue;
            }
            var other = new List<string>();
            foreach (var item in items)
            {
                context.Package.Visit();
                var head = context.Package.Head(item.Resolved!, 12) ?? Array.Empty<byte>();
                var known = type == "Audio"
                    ? head.Length >= 3 && (head[0] == 0x49 && head[1] == 0x44 && head[2] == 0x33 || head[0] == 0xFF && (head[1] & 0xE0) == 0xE0)
                    : head.Length >= 8 && Encoding.ASCII.GetString(head, 4, 4) == "ftyp";
                if (!known) other.Add(item.Resolved!);
            }
            // "MP3 等" / "MP4 等" admits other common formats, so an unrecognized payload is Unknown, not Fail.
            yield return other.Count > 0
                ? O(rule, clause, PreflightOutcome.Unknown, "Other" + type + "Format", other[0], "payload signature", Sample(other), name + " or another common format")
                : O(rule, clause, PreflightOutcome.Pass, name + "Payloads", items[0].Resolved!, "payload signature", items.Count + " payloads", name);
        }
        var limit = context.Options.MaxEmbeddedObjectBytes;
        var mediaOver = media.Where(x => context.Package.Length(x.Resolved!) > limit).Select(x => x.Resolved!).ToList();
        yield return media.Count == 0
            ? O(rule, "7.2(e)", PreflightOutcome.NotApplicable, "NoEmbeddedObjects", Package, Package, "No embedded media", "At most " + limit + " bytes")
            : mediaOver.Count > 0
                ? O(rule, "7.2(e)", PreflightOutcome.Fail, "EmbeddedObjectOverLimit", mediaOver[0], "uncompressed size", Sample(mediaOver), "At most " + limit + " bytes")
                : O(rule, "7.2(e)", PreflightOutcome.Pass, "EmbeddedObjectsWithinLimit", Package, "uncompressed size", media.Count + " objects", "At most " + limit + " bytes");
        var attachments = context.Structure.InternalAttachments.Where(x => context.Package.Exists(x.Path)).ToList();
        var attachmentsOver = attachments.Where(x => context.Package.Length(x.Path) > limit).Select(x => x.Path).ToList();
        yield return attachments.Count == 0
            ? O(rule, "7.4(c)", PreflightOutcome.NotApplicable, "NoInternalAttachments", Package, Package, "No internal attachment payload", "At most " + limit + " bytes")
            : attachmentsOver.Count > 0
                ? O(rule, "7.4(c)", PreflightOutcome.Fail, "AttachmentOverLimit", attachmentsOver[0], "uncompressed size", Sample(attachmentsOver), "External attachment above " + limit + " bytes")
                : O(rule, "7.4(c)", PreflightOutcome.Pass, "AttachmentsWithinLimit", Package, "uncompressed size", attachments.Count + " attachments", "At most " + limit + " bytes");
    }

    // MH-011 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Metadata(MedicalRuleContext context)
    {
        const string rule = "MH-011";
        var s = context.Structure;
        var body = s.DocBodies.FirstOrDefault();
        if (body == null || s.Namespace == null)
        {
            foreach (var clause in new[] { "7.1(c)", "7.3(a)", "7.3(b)" })
                yield return O(rule, clause, PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", Package, "No DocBody", "DocInfo metadata");
            yield break;
        }
        var ns = s.Namespace;
        var info = body.Element(ns + "DocInfo");
        yield return info != null && info.HasElements
            ? O(rule, "7.1(c)", PreflightOutcome.Pass, "MetadataInDocInfo", "OFD.xml", "/OFD/DocBody/DocInfo", info.Elements().Count() + " fields", "DocInfo")
            : O(rule, "7.1(c)", PreflightOutcome.Fail, "DocInfoMissing", "OFD.xml", "/OFD/DocBody/DocInfo", "absent or empty", "DocInfo");
        string? Field(string name) => info?.Elements(ns + name).Select(x => x.Value.Trim()).FirstOrDefault(x => x.Length > 0);
        string? Custom(string? name) => name == null ? null : info?.Element(ns + "CustomDatas")?.Elements(ns + "CustomData")
            .Where(x => (string?)x.Attribute("Name") == name).Select(x => x.Value.Trim()).FirstOrDefault(x => x.Length > 0);
        var mapping = context.Mapping;
        foreach (var (clause, items) in new[]
        {
            ("7.3(a)", new (string Label, string? Value, bool Mapped)[]
            {
                ("Author", Field("Author"), true), ("CreationDate", Field("CreationDate"), true),
                ("Keywords", Field("Keywords"), true), ("Patient:" + (mapping?.PatientField ?? "<unmapped>"), Custom(mapping?.PatientField), mapping?.PatientField != null)
            }),
            ("7.3(b)", new (string Label, string? Value, bool Mapped)[]
            {
                ("DocID", Field("DocID"), true), ("FileType:" + (mapping?.FileTypeField ?? "<unmapped>"), Custom(mapping?.FileTypeField), mapping?.FileTypeField != null),
                ("SecretLevel:" + (mapping?.SecretLevelField ?? "<unmapped>"), Custom(mapping?.SecretLevelField), mapping?.SecretLevelField != null)
            })
        })
        {
            var absent = items.Where(x => x.Mapped && x.Value == null).Select(x => x.Label).ToList();
            var unmapped = items.Where(x => !x.Mapped).Select(x => x.Label).ToList();
            var expected = string.Join(", ", items.Select(x => x.Label));
            yield return absent.Count > 0
                ? O(rule, clause, PreflightOutcome.Fail, "MetadataMissing", "OFD.xml", "/OFD/DocBody/DocInfo", "missing: " + string.Join(", ", absent), expected)
                : unmapped.Count > 0
                    ? O(rule, clause, PreflightOutcome.Unknown, "FieldMappingNotProvided", "OFD.xml", "/OFD/DocBody/DocInfo", "unmapped: " + string.Join(", ", unmapped), expected)
                    : O(rule, clause, PreflightOutcome.Pass, "MetadataPresent", "OFD.xml", "/OFD/DocBody/DocInfo", "all present", expected);
        }
    }

    // MH-012 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> MachineReadable(MedicalRuleContext context)
    {
        const string rule = "MH-012";
        var s = context.Structure;
        var payloads = s.TagPayloads.Select(x => (Path: x, Kind: "xml"))
            .Concat(s.InternalAttachments.Select(x => (x.Path, Kind: Kind(x.Path, x.Format))).Where(x => x.Kind != null).Select(x => (x.Path, Kind: x.Kind!)))
            .Where(x => context.Package.Exists(x.Path)).Distinct().ToList();
        if (payloads.Count == 0)
        {
            yield return O(rule, "7.5(a)", PreflightOutcome.NotApplicable, "NoMachineReadablePayload", Package, Package, "No tag or XML/JSON attachment payload", "XML or JSON");
            yield return O(rule, "7.5(b)", PreflightOutcome.NotApplicable, "NoMachineReadablePayload", Package, Package, "No tag or XML/JSON attachment payload", "UTF-8");
            yield break;
        }
        var notUtf8 = new List<string>(); var notStructured = new List<string>();
        var strict = new UTF8Encoding(false, true);
        foreach (var (path, kind) in payloads)
        {
            context.Package.Visit();
            var bytes = context.Package.Bytes(path)!;
            string text;
            try { text = strict.GetString(bytes); }
            catch (ArgumentException) { notUtf8.Add(path); continue; }
            text = text.TrimStart('﻿');
            if (kind == "xml")
            {
                var document = context.Package.Xml(path);
                if (document == null) { notStructured.Add(path); continue; }
                var encoding = document.Declaration?.Encoding;
                if (encoding != null && !string.Equals(encoding, "UTF-8", StringComparison.OrdinalIgnoreCase)) notUtf8.Add(path);
            }
            else
            {
                var trimmed = text.Trim();
                if (!(trimmed.StartsWith("{", StringComparison.Ordinal) && trimmed.EndsWith("}", StringComparison.Ordinal) ||
                      trimmed.StartsWith("[", StringComparison.Ordinal) && trimmed.EndsWith("]", StringComparison.Ordinal)))
                    notStructured.Add(path);
            }
        }
        yield return notStructured.Count > 0
            ? O(rule, "7.5(a)", PreflightOutcome.Fail, "NotSemiStructured", notStructured[0], Package, Sample(notStructured), "Well-formed XML or JSON object/array")
            : O(rule, "7.5(a)", PreflightOutcome.Pass, "SemiStructuredText", payloads[0].Path, Package, payloads.Count + " payloads (JSON checked by shape only)", "XML or JSON");
        yield return notUtf8.Count > 0
            ? O(rule, "7.5(b)", PreflightOutcome.Fail, "NotUtf8", notUtf8[0], Package, Sample(notUtf8), "UTF-8")
            : O(rule, "7.5(b)", PreflightOutcome.Pass, "Utf8", payloads[0].Path, Package, payloads.Count + " payloads", "UTF-8");
    }

    private static string? Kind(string path, string? format)
    {
        var value = (format ?? string.Empty).Trim().TrimStart('.').ToLowerInvariant();
        if (value == "xml" || path.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)) return "xml";
        if (value == "json" || path.EndsWith(".json", StringComparison.OrdinalIgnoreCase)) return "json";
        return null;
    }

    // MH-013 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> SignatureCoverage(MedicalRuleContext context)
    {
        const string rule = "MH-013";
        var s = context.Structure;
        if (s.Root?.Root == null)
        {
            foreach (var clause in new[] { "6.3(e)", "6.3(g)", "8(c)", "7.3(d)" })
                yield return O(rule, clause, PreflightOutcome.Unknown, "RootUnavailable", "OFD.xml", Package, "OFD.xml unavailable", "Signature coverage");
            yield break;
        }
        if (s.SignaturePaths.Count == 0 && s.SignaturesPath == null)
        {
            foreach (var clause in new[] { "6.3(e)", "6.3(g)", "8(c)" })
                yield return O(rule, clause, PreflightOutcome.NotApplicable, "NoSignatures", "OFD.xml", "/OFD/DocBody/Signatures", "absent", "Signature coverage");
            yield return O(rule, "7.3(d)", PreflightOutcome.Fail, "NoSignature", "OFD.xml", "/OFD/DocBody/Signatures", "absent", "Signature or seal present");
            yield break;
        }
        var latest = s.SignaturePaths.LastOrDefault();
        var signature = latest == null ? null : context.Package.Xml(latest)?.Root;
        if (signature == null)
        {
            foreach (var clause in new[] { "6.3(e)", "6.3(g)", "8(c)", "7.3(d)" })
                yield return O(rule, clause, PreflightOutcome.Unknown, "SignatureUnreadable", latest ?? s.SignaturesPath ?? "OFD.xml", Package, "Signature file unavailable", "Readable Signature.xml");
            yield break;
        }
        var ns = signature.Name.Namespace;
        var protectedEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var unsafeRefs = new List<string>();
        foreach (var reference in signature.Descendants(ns + "Reference"))
        {
            context.Package.Visit();
            var fileRef = (string?)reference.Attribute("FileRef");
            if (fileRef != null && MedicalPackage.TryResolve("x", "/" + fileRef.TrimStart('/'), out var path)) protectedEntries.Add(path);
            else unsafeRefs.Add(fileRef ?? "<missing>");
        }
        var excluded = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (s.SignaturesPath != null) excluded.Add(s.SignaturesPath);
        var signDirectories = s.SignaturePaths.Select(x => x.Substring(0, x.LastIndexOf('/') + 1)).ToList();
        if (s.AnnotationsPath != null) excluded.Add(s.AnnotationsPath);
        foreach (var path in s.AnnotationPaths) excluded.Add(path);
        var required = context.Package.EntryNames
            .Where(x => !excluded.Contains(x) && !signDirectories.Any(d => d.Length > 0 && x.StartsWith(d, StringComparison.OrdinalIgnoreCase)))
            .OrderBy(x => x, StringComparer.Ordinal).ToList();
        var missingRefs = protectedEntries.Where(x => !context.Package.Exists(x)).ToList();
        var unprotected = required.Where(x => !protectedEntries.Contains(x)).ToList();
        var pages = s.PagePaths.Where(x => !protectedEntries.Contains(x)).ToList();
        var annotationIndexUnreadable = s.AnnotationsPath != null && s.Unreadable.Contains(s.AnnotationsPath);

        // FileRef resolves from the package root, as OfdSignatureVerifier and archival preflight do. While any reference
        // is unresolvable it may still name an unlisted page, so a missing page is Unknown rather than Fail.
        yield return pages.Count > 0 && unsafeRefs.Count > 0
            ? O(rule, "6.3(e)", PreflightOutcome.Unknown, "UnresolvableReference", latest!, "Reference/@FileRef", Sample(unsafeRefs), "Every page description referenced")
            : pages.Count > 0
            ? O(rule, "6.3(e)", PreflightOutcome.Fail, "PageContentNotProtected", latest!, "References", Sample(pages), "Every page description referenced")
            : s.PagePaths.Count == 0
                ? O(rule, "6.3(e)", PreflightOutcome.Unknown, "PagesUnresolved", latest!, "References", "No resolved pages", "Every page description referenced")
                : O(rule, "6.3(e)", PreflightOutcome.Pass, "PageContentProtected", latest!, "References", s.PagePaths.Count + " pages", "Every page description referenced");
        foreach (var clause in new[] { "8(c)", "6.3(g)" })
        {
            var reason = clause == "8(c)" ? "ContentNotProtected" : "UnsignedModification";
            if (unsafeRefs.Count > 0)
                yield return O(rule, clause, PreflightOutcome.Unknown, "UnresolvableReference", latest!, "Reference/@FileRef", Sample(unsafeRefs), "Package-relative FileRef");
            else if (missingRefs.Count > 0)
                yield return O(rule, clause, PreflightOutcome.Fail, "ReferencedEntryMissing", latest!, "Reference/@FileRef", Sample(missingRefs), "Referenced entries exist");
            else if (unprotected.Count > 0)
                yield return O(rule, clause, PreflightOutcome.Fail, reason, latest!, "References", unprotected.Count + " entries: " + Sample(unprotected), "All content except annotation and signature files");
            else if (annotationIndexUnreadable)
                yield return O(rule, clause, PreflightOutcome.Unknown, "AnnotationIndexUnreadable", s.AnnotationsPath!, Package, "Exclusions unknown", "Readable annotation index");
            else
                yield return O(rule, clause, PreflightOutcome.Pass, "StructuralCoverageComplete", latest!, "References", protectedEntries.Count + " references; digests not verified", "All content except annotation and signature files");
        }
        yield return O(rule, "7.3(d)", PreflightOutcome.Pass, "SignaturePresent", s.SignaturesPath ?? latest!, "Signature", s.SignaturePaths.Count + " signatures", "Signature or seal present");
    }

    // MH-014 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Algorithms(MedicalRuleContext context)
    {
        const string rule = "MH-014";
        var s = context.Structure;
        if (s.SignaturePaths.Count == 0)
        {
            var outcome = s.SignaturesPath != null && s.Unreadable.Contains(s.SignaturesPath) ? PreflightOutcome.Unknown : PreflightOutcome.NotApplicable;
            yield return O(rule, "6.3(b)", outcome, "NoSignatures", "OFD.xml", "/OFD/DocBody/Signatures", "none", "GB/T 33560 identifiers");
            yield return O(rule, "8(d)", outcome, "NoSignatures", "OFD.xml", "/OFD/DocBody/Signatures", "none", "SM2/SM3");
            yield break;
        }
        var values = new List<(string Path, string Kind, string? Value)>();
        foreach (var path in s.SignaturePaths)
        {
            context.Package.Visit();
            var root = context.Package.Xml(path)?.Root;
            if (root == null) { values.Add((path, "file", null)); continue; }
            var ns = root.Name.Namespace;
            values.Add((path, "SignatureMethod", root.Descendants(ns + "SignatureMethod").Select(x => x.Value.Trim()).FirstOrDefault()));
            values.Add((path, "CheckMethod", (string?)root.Descendants(ns + "References").FirstOrDefault()?.Attribute("CheckMethod")));
        }
        var nonSm = values.Where(x => x.Value != null && NonSm(x.Value)).ToList();
        var unknown = values.Where(x => x.Value == null || !Sm(x.Value) && !NonSm(x.Value)).ToList();
        var names = values.Where(x => x.Value != null && Sm(x.Value) && x.Value != Sm2Oid && x.Value != Sm3Oid).ToList();
        // SignatureMethod must name SM2 and CheckMethod SM3; an SM identifier in the wrong role is still wrong.
        var misrole = values.Where(x => x.Value != null && Sm(x.Value) && !RoleMatches(x.Kind, x.Value)).ToList();
        string Show((string Path, string Kind, string? Value) x) => x.Path + " " + x.Kind + "=" + (x.Value ?? "<absent>");
        yield return nonSm.Count > 0
            ? O(rule, "8(d)", PreflightOutcome.Fail, "NonSmAlgorithm", nonSm[0].Path, nonSm[0].Kind, Sample(nonSm.Select(Show)), "SM2 signature, SM3 digest")
            : misrole.Count > 0
            ? O(rule, "8(d)", PreflightOutcome.Fail, "AlgorithmRoleMismatch", misrole[0].Path, misrole[0].Kind, Sample(misrole.Select(Show)), "SignatureMethod SM2, CheckMethod SM3")
            : unknown.Count > 0
                ? O(rule, "8(d)", PreflightOutcome.Unknown, "AlgorithmUnrecognized", unknown[0].Path, unknown[0].Kind, Sample(unknown.Select(Show)), "SM2 signature, SM3 digest")
                : O(rule, "8(d)", PreflightOutcome.Pass, "SmAlgorithms", values[0].Path, "SignatureMethod|CheckMethod", values.Count + " identifiers", "SM2 signature, SM3 digest");
        yield return names.Count > 0
            ? O(rule, "6.3(b)", PreflightOutcome.Fail, "IdentifierNotOid", names[0].Path, names[0].Kind, Sample(names.Select(Show)), Sm2Oid + " / " + Sm3Oid)
            : misrole.Count > 0
            ? O(rule, "6.3(b)", PreflightOutcome.Fail, "AlgorithmRoleMismatch", misrole[0].Path, misrole[0].Kind, Sample(misrole.Select(Show)), "SignatureMethod " + Sm2Oid + ", CheckMethod " + Sm3Oid)
            : nonSm.Count + unknown.Count > 0
                ? O(rule, "6.3(b)", PreflightOutcome.Unknown, "IdentifierNotAssessed", values[0].Path, Package, "Non-SM or unrecognized identifiers (see 8(d))", Sm2Oid + " / " + Sm3Oid)
                : O(rule, "6.3(b)", PreflightOutcome.Pass, "Gbt33560Identifiers", values[0].Path, "SignatureMethod|CheckMethod", values.Count + " identifiers", Sm2Oid + " / " + Sm3Oid);
    }

    private static bool Sm(string value) => value == Sm2Oid || value == Sm3Oid ||
        Regex.IsMatch(value, @"^(sm2|sm3|sm2-?with-?sm3|sm3-?with-?sm2)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static bool RoleMatches(string kind, string value) => kind == "CheckMethod"
        ? value == Sm3Oid || string.Equals(value, "sm3", StringComparison.OrdinalIgnoreCase)
        : value == Sm2Oid || Regex.IsMatch(value, @"^(sm2|sm2-?with-?sm3|sm3-?with-?sm2)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    private static bool NonSm(string value) =>
        value.StartsWith("1.2.840.113549.", StringComparison.Ordinal) || value.StartsWith("2.16.840.1.101.3.4.", StringComparison.Ordinal) ||
        value.StartsWith("1.2.840.10045.", StringComparison.Ordinal) || value == "1.3.14.3.2.26" ||
        Regex.IsMatch(value, @"^(sha-?(1|224|256|384|512)|md5|rsa.*|ecdsa.*|.*with-?rsa.*)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

    // MH-015 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Permissions(MedicalRuleContext context)
    {
        const string rule = "MH-015", clause = "7.3(g)";
        var document = context.Structure.Document?.Root;
        if (document == null)
        {
            yield return O(rule, clause, PreflightOutcome.Unknown, "DocumentUnavailable", "OFD.xml", Package, "Document unavailable", "Annotation and signing not prohibited");
            yield break;
        }
        var ns = document.Name.Namespace;
        var permissions = document.Element(ns + "Permissions");
        var prohibited = new List<string>();
        foreach (var name in new[] { "Annot", "Signature" })
        {
            var value = permissions?.Element(ns + name)?.Value.Trim();
            if (string.Equals(value, "false", StringComparison.Ordinal) || value == "0") prohibited.Add(name);
        }
        yield return prohibited.Count > 0
            ? O(rule, clause, PreflightOutcome.Fail, "OperationProhibited", context.Structure.DocumentPath!, "Document/Permissions", string.Join(", ", prohibited) + "=false", "Annot and Signature not false")
            : O(rule, clause, PreflightOutcome.Pass, "NotProhibited", context.Structure.DocumentPath!, "Document/Permissions",
                permissions == null ? "no Permissions" : "Annot/Signature not false", "Annot and Signature not false");
    }

    // MH-016 -------------------------------------------------------------------------------------------------
    internal static IEnumerable<MedicalRuleObservation> Encryption(MedicalRuleContext context)
    {
        const string rule = "MH-016", clause = "8(b)";
        var encryption = context.Package.EntryNames.FirstOrDefault(x => x.Equals("Encryptions.xml", StringComparison.OrdinalIgnoreCase) ||
            x.EndsWith("/Encryptions.xml", StringComparison.OrdinalIgnoreCase));
        yield return encryption != null
            ? O(rule, clause, PreflightOutcome.Unknown, "TemporaryEncryptionDeclared", encryption, Package, "Encrypted package", "Decrypted before archiving or migration")
            : O(rule, clause, PreflightOutcome.NotApplicable, "NoEncryption", Package, Package, "No Encryptions.xml", "Decrypted before archiving or migration");
    }
}
