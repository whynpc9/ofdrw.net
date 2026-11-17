using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml.Linq;

namespace Ofdrw.Net.Archival.Internal;

/// <summary>What the scope inspection established about the package, for the report and for clause checks.</summary>
internal sealed class ArchivalScopeResult
{
    internal List<ArchivalScopeFinding> Findings { get; } = new List<ArchivalScopeFinding>();
    internal XElement? OfdRoot { get; set; }
    internal string? RawDocType { get; set; }
    internal bool ForeignDocType { get; set; }
    internal string? DocumentEntry { get; set; }
    internal XElement? Document { get; set; }
    /// <summary>Resolved page and template content entries with their parsed roots.</summary>
    internal List<KeyValuePair<string, XElement>> Pages { get; } = new List<KeyValuePair<string, XElement>>();
    internal string? BudgetExceeded { get; set; }

    internal ArchivalInputSupport Support =>
        BudgetExceeded != null ? ArchivalInputSupport.Unknown :
        Findings.Any(x => x.Support == ArchivalInputSupport.OutOfScope) ? ArchivalInputSupport.OutOfScope :
        Findings.Any(x => x.Support == ArchivalInputSupport.Unknown) ? ArchivalInputSupport.Unknown :
        ArchivalInputSupport.Supported;
}

/// <summary>
/// Decides whether a package lies inside <see cref="ArchivalInputProfile.SelfGeneratedSimple"/>. It only classifies
/// scope; it never produces clause outcomes, so an out-of-scope input cannot fabricate a standards failure.
/// </summary>
internal sealed class ArchivalScopeInspector
{
    private static readonly HashSet<string> DocBodyChildren = new HashSet<string>(StringComparer.Ordinal)
    { "DocInfo", "DocRoot", "Versions", "Signatures" };
    private static readonly HashSet<string> DocumentChildren = new HashSet<string>(StringComparer.Ordinal)
    { "CommonData", "Pages", "Outlines", "Permissions", "Actions", "VPreferences", "Bookmarks", "Annotations", "CustomTags", "Attachments", "Extensions" };
    private static readonly HashSet<string> PageChildren = new HashSet<string>(StringComparer.Ordinal)
    { "Area", "Template", "PageRes", "Content", "Actions" };
    private static readonly HashSet<string> LayerObjects = new HashSet<string>(StringComparer.Ordinal)
    { "TextObject", "PathObject", "ImageObject", "PageBlock", "CompositeObject" };

    private readonly ArchivalXmlReader xml;
    private readonly ArchivalScopeResult result = new ArchivalScopeResult();

    internal ArchivalScopeInspector(ArchivalXmlReader xml) { this.xml = xml; }

    internal ArchivalScopeResult Inspect()
    {
        try { InspectPackage(); }
        catch (ArchivalBudgetExceededException ex)
        {
            result.BudgetExceeded = ex.Budget;
            Add("ScopeInspectionBudgetExceeded", ArchivalInputSupport.Unknown, "<input>", "<package>", ex.Budget + " budget");
        }
        return result;
    }

    private void Add(string code, ArchivalInputSupport support, string entry, string selector, string observed,
        string expected = "Content accepted by " + ArchivalInputProfile.SelfGeneratedSimpleId)
    {
        if (result.Findings.Any(x => x.Code == code && x.Evidence.Entry == entry)) return; // one finding per code and entry
        result.Findings.Add(new ArchivalScopeFinding(code, support, new ArchivalEvidenceLocation(entry, selector, observed, expected)));
    }

    private void OutOfScope(string code, string entry, XElement node, string observed) =>
        Add(code, ArchivalInputSupport.OutOfScope, entry, ArchivalPackagePath.Selector(node), observed);

    private void Unknown(string code, string entry, XElement? node, string observed) =>
        Add(code, ArchivalInputSupport.Unknown, entry, node == null ? "<package>" : ArchivalPackagePath.Selector(node), observed);

    private void InspectPackage()
    {
        var archive = xml.Archive;
        if (archive.Contains("Encryptions.xml"))
            Add("EncryptionMarkerPresent", ArchivalInputSupport.OutOfScope, "Encryptions.xml", "<package>", "Root Encryptions.xml present");
        if (archive.EntryNames.Any(x => x.StartsWith("PasswordCrypto/", StringComparison.OrdinalIgnoreCase)))
            Add("EncryptionMarkerPresent", ArchivalInputSupport.OutOfScope, "PasswordCrypto/", "<package>", "PasswordCrypto entries present");
        var signatureDeclared = false;
        var attachmentDeclared = false;
        var root = xml.Read("OFD.xml")?.Root;
        if (root == null) { Unknown("RootScopeUnavailable", "OFD.xml", null, "OFD.xml missing or not well-formed"); return; }
        if (root.Name.LocalName != "OFD" || !ArchivalPackagePath.Supported(root.Name.Namespace))
        { Unknown("RootScopeUnavailable", "OFD.xml", root, "Unsupported root element or namespace"); return; }
        result.OfdRoot = root;
        result.RawDocType = (string?)root.Attribute("DocType");
        result.ForeignDocType = root.Attributes().Any(x => x.Name.LocalName == "DocType" && x.Name.Namespace != XNamespace.None);
        var ns = root.Name.Namespace;
        foreach (var child in root.Elements().Where(x => x.Name != ns + "DocBody"))
            Unknown("UnrecognizedRootContent", "OFD.xml", child, "Unrecognized OFD child");
        var bodies = root.Elements(ns + "DocBody").ToList();
        if (bodies.Count > 1) OutOfScope("MultipleDocuments", "OFD.xml", root, "DocBody count=" + bodies.Count);
        if (bodies.Count == 0) { Unknown("DocumentScopeUnavailable", "OFD.xml", root, "DocBody count=0"); return; }
        foreach (var body in bodies)
            foreach (var child in body.Elements())
            {
                if (child.Name.Namespace != ns || !DocBodyChildren.Contains(child.Name.LocalName))
                    Unknown("UnrecognizedDocBodyContent", "OFD.xml", child, "Unrecognized DocBody child");
                else if (child.Name.LocalName == "Versions") OutOfScope("VersionsDeclared", "OFD.xml", child, "Versions declared");
                else if (child.Name.LocalName == "Signatures")
                { signatureDeclared = true; OutOfScope("SignaturesDeclared", "OFD.xml", child, "Signatures declared"); }
            }
        if (bodies.Count == 1) InspectDocument(bodies[0], ns, ref attachmentDeclared);
        foreach (var entry in archive.EntryNames.OrderBy(x => x, StringComparer.Ordinal))
        {
            var name = entry.Substring(entry.LastIndexOf('/') + 1);
            if (!signatureDeclared && (name.Equals("Signatures.xml", StringComparison.OrdinalIgnoreCase) ||
                name.Equals("SignedValue.dat", StringComparison.OrdinalIgnoreCase)))
                Add("SignaturePayloadPresent", ArchivalInputSupport.OutOfScope, entry, "<package>", "Signature payload without declaration");
            if (!attachmentDeclared && name.Equals("Attachments.xml", StringComparison.OrdinalIgnoreCase))
                Add("AttachmentPayloadPresent", ArchivalInputSupport.OutOfScope, entry, "<package>", "Attachment index without declaration");
        }
    }

    private void InspectDocument(XElement body, XNamespace ns, ref bool attachmentDeclared)
    {
        var roots = body.Elements(ns + "DocRoot").ToList();
        if (roots.Count != 1 || roots[0].HasElements)
        { Unknown("DocumentScopeUnavailable", "OFD.xml", body, "DocRoot count=" + roots.Count); return; }
        var path = roots[0].Value.Trim();
        if (ArchivalPackagePath.Unsafe(path) || !ArchivalPackagePath.TryResolve("OFD.xml", path, out var entry))
        { Unknown("DocumentScopeUnavailable", "OFD.xml", roots[0], "DocRoot unresolved"); return; }
        xml.Visit();
        var document = xml.Read(entry)?.Root;
        if (document == null || document.Name.LocalName != "Document" || !ArchivalPackagePath.Supported(document.Name.Namespace))
        { Unknown("DocumentScopeUnavailable", entry, null, "Document missing, malformed or unsupported"); return; }
        result.DocumentEntry = entry;
        result.Document = document;
        var dns = document.Name.Namespace;
        foreach (var child in document.Elements())
        {
            if (child.Name.Namespace != dns || !DocumentChildren.Contains(child.Name.LocalName))
            { Unknown("UnrecognizedDocumentContent", entry, child, "Unrecognized Document child"); continue; }
            switch (child.Name.LocalName)
            {
                case "Attachments": attachmentDeclared = true; OutOfScope("AttachmentsDeclared", entry, child, "Attachments declared"); break;
                case "Annotations": OutOfScope("AnnotationsDeclared", entry, child, "Annotations declared"); break;
                case "Extensions": OutOfScope("ExtensionsDeclared", entry, child, "Extensions declared"); break;
                case "Actions": OutOfScope("DocumentActionsDeclared", entry, child, "Document actions declared"); break;
                case "Outlines":
                    var action = child.Descendants().FirstOrDefault(x => x.Name.LocalName == "Actions" || x.Name.LocalName == "Action");
                    if (action != null) OutOfScope("OutlineActionsDeclared", entry, action, "Outline actions declared");
                    break;
            }
        }
        var pages = document.Elements(dns + "Pages").SelectMany(x => x.Elements(dns + "Page"))
            .Concat(document.Elements(dns + "CommonData").SelectMany(x => x.Elements(dns + "TemplatePage"))).ToList();
        foreach (var page in pages) InspectPage(entry, page);
    }

    private void InspectPage(string documentEntry, XElement declaration)
    {
        xml.Visit();
        var baseLoc = (string?)declaration.Attribute("BaseLoc");
        if (ArchivalPackagePath.Unsafe(baseLoc) || !ArchivalPackagePath.TryResolve(documentEntry, baseLoc!, out var entry))
        { Unknown("PageContentUnavailable", documentEntry, declaration, "Page BaseLoc unresolved"); return; }
        var page = xml.Read(entry)?.Root;
        if (page == null || page.Name.LocalName != "Page" || !ArchivalPackagePath.Supported(page.Name.Namespace))
        { Unknown("PageContentUnavailable", entry, null, "Page missing, malformed or unsupported"); return; }
        result.Pages.Add(new KeyValuePair<string, XElement>(entry, page));
        var ns = page.Name.Namespace;
        var foreign = page.DescendantsAndSelf().FirstOrDefault(x => x.Name.Namespace != ns);
        if (foreign != null) Unknown("UnrecognizedPageContent", entry, foreign, "Foreign-namespace page content");
        foreach (var child in page.Elements().Where(x => x.Name.Namespace == ns && !PageChildren.Contains(x.Name.LocalName)))
            Unknown("UnrecognizedPageContent", entry, child, "Unrecognized Page child");
        var actions = page.Descendants(ns + "Actions").FirstOrDefault();
        if (actions != null)
            OutOfScope(actions.Parent == page ? "PageActionsDeclared" : "ObjectActionsDeclared", entry, actions, "Actions declared");
        foreach (var layer in page.Descendants(ns + "Layer"))
            foreach (var item in layer.Elements().Where(x => x.Name.Namespace == ns))
                if (!LayerObjects.Contains(item.Name.LocalName))
                    Unknown("UnrecognizedPageObject", entry, item, "Unrecognized layer object");
        var composite = page.Descendants(ns + "CompositeObject").FirstOrDefault();
        if (composite != null) OutOfScope("CompositeGraphicsPresent", entry, composite, "CompositeObject present");
    }
}
