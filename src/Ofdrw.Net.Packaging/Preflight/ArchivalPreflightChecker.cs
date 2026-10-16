using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.ExceptionServices;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.IO;
using Ofdrw.Net.Packaging.Archive;

namespace Ofdrw.Net.Packaging.Preflight;

/// <summary>Bounded, local, read-only OFD ZIP preflight. Does not fetch, render, decrypt, repair or verify signatures.</summary>
public sealed class ArchivalPreflightChecker
{
    /// <summary>Inspect a stable input snapshot, leaving the stream open. Seekable input must start at zero and is restored on every exit.</summary>
    /// <param name="input">Readable ZIP stream. Non-seekable input is consumed within the compressed input budget.</param>
    /// <param name="options">Local policy and budgets; copied before any asynchronous input operation.</param>
    /// <param name="cancellationToken">Cancellation propagates as OperationCanceledException, with stream restoration.</param>
    /// <returns>Four-state observations and explicit coverage limitations, never certification.</returns>
    /// <exception cref="ArgumentException">Input or options violate the borrowing contract.</exception>
    /// <exception cref="IOException">An ordinary input or position-restoration I/O operation fails.</exception>
    public async Task<ArchivalPreflightReport> CheckAsync(Stream input, ArchivalPreflightOptions options,
        CancellationToken cancellationToken = default)
    {
        if (input == null) throw new ArgumentNullException(nameof(input));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (!input.CanRead) throw new ArgumentException("Input must be readable.", nameof(input));
        if (input.CanSeek && input.Position != 0) throw new ArgumentException("Seekable input must start at zero.", nameof(input));
        var snapshot = options.Snapshot();
        var run = new Inspection(snapshot, cancellationToken);
        try
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                using var borrowed = new ReadOnlyBorrow(input);
                OfdPackageArchive archive;
                try { archive = await new OfdPackageLoader().LoadAsync(borrowed, snapshot.LoadOptions, cancellationToken).ConfigureAwait(false); }
                catch (InvalidDataException) when (borrowed.Failure != null)
                { borrowed.Failure.Throw(); throw; }
                run.Inspect(archive);
            }
            catch (InvalidDataException) { run.Interrupt("InputRejected"); }
            catch (BudgetExceededException) { run.Interrupt("InspectionBudgetExceeded"); }
            return run.Finish();
        }
        finally { if (input.CanSeek) input.Position = 0; }
    }

    private sealed class BudgetExceededException : Exception { }

    // ZipArchive may wrap an underlying IOException as InvalidDataException. Preserve the caller's ordinary I/O failure.
    private sealed class ReadOnlyBorrow : Stream
    {
        private readonly Stream source;
        internal ReadOnlyBorrow(Stream source) { this.source = source; }
        internal ExceptionDispatchInfo? Failure { get; private set; }
        private T Io<T>(Func<T> action)
        {
            try { return action(); }
            catch (IOException ex) { Failure = ExceptionDispatchInfo.Capture(ex); throw; }
        }
        public override bool CanRead => source.CanRead;
        public override bool CanSeek => source.CanSeek;
        public override bool CanWrite => false;
        public override long Length => Io(() => source.Length);
        public override long Position { get => Io(() => source.Position); set => Io(() => { source.Position = value; return 0; }); }
        public override int Read(byte[] buffer, int offset, int count) => Io(() => source.Read(buffer, offset, count));
        public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken)
        {
            try { return await source.ReadAsync(buffer, offset, count, cancellationToken).ConfigureAwait(false); }
            catch (IOException ex) { Failure = ExceptionDispatchInfo.Capture(ex); throw; }
        }
        public override long Seek(long offset, SeekOrigin origin)
        {
            // ZIP probing a too-short input is malformed data, rather than an underlying transport failure.
            var start = origin == SeekOrigin.Begin ? 0 : origin == SeekOrigin.Current ? Position : Length;
            if (offset < -start) throw new InvalidDataException("ZIP seek precedes input.");
            return Io(() => source.Seek(offset, origin));
        }
        public override void Flush() => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
    }

    private sealed class Inspection
    {
        private readonly ArchivalPreflightOptions options;
        private readonly CancellationToken token;
        private readonly List<PreflightRuleResult> results = new List<PreflightRuleResult>();
        private readonly HashSet<string> gaps = new HashSet<string>(StringComparer.Ordinal)
        {
            "Standard clauses, OFD-H/OFD conformance and archival certification are not assessed.",
            "Font format, glyph coverage, embedding rights and licensing are not assessed.",
            "Signature digests, cryptography, certificates, timestamps and legal validity are not assessed.",
            "Unknown XML, actions, nested payload contents, full resource-ID closure and XSD validation are not covered."
        };
        private readonly Dictionary<string, XDocument?> xml = new Dictionary<string, XDocument?>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> resourceEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> attachmentEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> signatureEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private OfdPackageArchive package = null!;
        private PreflightRunState state;
        private long xmlBytes;
        private int nodes, references, entryCount, documentCount;
        private bool scopeUnknown;

        internal Inspection(ArchivalPreflightOptions options, CancellationToken token) { this.options = options; this.token = token; }
        private static string Rule(int rule) => "APF-00" + rule;
        private void Check() => token.ThrowIfCancellationRequested();
        private void Add(int rule, PreflightOutcome outcome, string reason, string entry, XElement? node = null, string observed = "", string expected = "Known local declaration and existing package target")
        {
            Check();
            if (rule == 1 && !options.RequireSingleDocument) outcome = PreflightOutcome.NotApplicable;
            if (results.Count >= options.MaxResults - 6) throw new BudgetExceededException();
            results.Add(new PreflightRuleResult(Rule(rule), outcome, reason,
                new PreflightEvidence(Clean(entry), node == null ? "<package>" : Selector(node), Clean(observed), expected)));
        }
        private static string Clean(string value)
        {
            // Paths with a scheme or UNC are never echoed: credentials, query values and host names are unnecessary evidence.
            if (value.Contains(":") || value.Contains("?") || value.Contains("#") || value.Contains("%") || value.Contains("\\\\") || value.Contains("//")) return "<redacted ambiguous or external reference>";
            return value.Length > 512 ? value.Substring(0, 512) + "…" : value;
        }
        private static string Selector(XElement node)
        {
            var parts = new List<string>();
            foreach (var item in node.AncestorsAndSelf().Reverse())
            {
                var index = item.ElementsBeforeSelf(item.Name).Count() + 1;
                var ns = Supported(item) ? item.Name.NamespaceName : "<unsupported-namespace>";
                parts.Add("{" + ns + "}" + item.Name.LocalName + "[" + index.ToString(System.Globalization.CultureInfo.InvariantCulture) + "]");
            }
            var selector = "/" + string.Join("/", parts);
            return selector.Length <= 8192 ? selector : "<selector exceeds metadata limit>";
        }
        private static string Identity(XElement node)
        {
            var owner = node.AncestorsAndSelf().FirstOrDefault(x => x.Attribute("ID") != null);
            return owner == null ? string.Empty : "; ID=" + (string?)owner.Attribute("ID") + "; Name=" + (string?)owner.Attribute("Name");
        }
        private static bool Supported(XElement root) => root.Name.NamespaceName == "http://www.ofdspec.org" || root.Name.NamespaceName == "http://www.ofdspec.org/2016";
        private static IEnumerable<XElement> Children(XElement root, string name) => root.Elements(root.Name.Namespace + name);
        private static IEnumerable<XElement> Descendants(XElement root, string name) => root.Descendants(root.Name.Namespace + name);
        private void Uncertain(string reason, string entry, XElement? node = null)
        {
            scopeUnknown = true;
            gaps.Add(reason + ": " + Clean(entry));
            foreach (var rule in new[] { 2, 3, 5, 6 }) Add(rule, PreflightOutcome.Unknown, reason, entry, node);
            if (options.RequireNoEncryption) Add(4, PreflightOutcome.Unknown, "UnrecognizedControlOrEncryptionScope", entry, node);
        }
        private XDocument? ReadXml(string entry, int rule, string rootName, bool malformedUnknown = false)
        {
            Check();
            if (!xml.TryGetValue(entry, out var doc))
            {
                var bytes = package.GetBytes(entry);
                if (bytes.Length > options.MaxSingleXmlBytes || bytes.Length > options.MaxXmlBytes - xmlBytes) throw new BudgetExceededException();
                xmlBytes += bytes.Length;
                var settings = new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
                    MaxCharactersInDocument = options.MaxSingleXmlBytes, MaxCharactersFromEntities = 1
                };
                try
                {
                    // Count before building the tree: limits apply to allocation, not merely to traversal afterwards.
                    using (var source = new MemoryStream(bytes, false))
                    using (var reader = XmlReader.Create(source, settings))
                        while (reader.Read())
                        {
                            Check();
                            if (++nodes > options.MaxXmlNodes || reader.Depth > options.MaxXmlDepth) throw new BudgetExceededException();
                        }
                    using (var source = new MemoryStream(bytes, false))
                    using (var reader = XmlReader.Create(source, settings)) doc = XDocument.Load(reader);
                }
                catch (XmlException ex)
                {
                    // Unsupported declared encodings are evidence gaps; malformed local XML and prohibited DTD are diagnostics.
                    var encoding = ex.InnerException is ArgumentException || ex.InnerException is NotSupportedException;
                    Add(rule, malformedUnknown || encoding ? PreflightOutcome.Unknown : PreflightOutcome.Fail,
                        encoding ? "UnsupportedXmlEncoding" : "InvalidOrProhibitedXml", entry);
                    Uncertain("UnreadableControlXml", entry);
                    xml[entry] = null;
                    return null;
                }
                xml[entry] = doc;
            }
            if (doc?.Root == null) return null;
            if (!Supported(doc.Root) || doc.Root.Name.LocalName != rootName)
            {
                Add(rule, PreflightOutcome.Unknown, "UnsupportedRootOrNamespace", entry, doc.Root);
                Uncertain("UnsupportedControlScope", entry, doc.Root);
                if (options.RequireNoEncryption) Add(4, PreflightOutcome.Unknown, "UnrecognizedEncryptionExtension", entry, doc.Root);
                return null;
            }
            if (doc.Root.DescendantsAndSelf().Any(x => x.Name.Namespace != doc.Root.Name.Namespace ||
                x.Attributes().Any(a => !a.IsNamespaceDeclaration && a.Name.NamespaceName.Length != 0 && a.Name.Namespace != doc.Root.Name.Namespace)))
            {
                Uncertain("UnsupportedNamespace", entry, doc.Root);
                Add(1, PreflightOutcome.Unknown, "UnsupportedNamespace", entry, doc.Root);
                if (options.RequireNoEncryption) Add(4, PreflightOutcome.Unknown, "UnrecognizedEncryptionExtension", entry, doc.Root);
            }
            if (options.RequireNoEncryption && doc.Root.DescendantsAndSelf().Any(x => x.Name.LocalName.IndexOf("Encrypt", StringComparison.OrdinalIgnoreCase) >= 0 || x.Name.LocalName.IndexOf("Crypto", StringComparison.OrdinalIgnoreCase) >= 0))
                Add(4, PreflightOutcome.Unknown, "UnrecognizedEncryptionExtension", entry, doc.Root);
            if (doc.Root.DescendantsAndSelf().Any(x => new[] { "Extensions", "CustomTags", "Annotations", "CompositeGraphicUnits" }.Contains(x.Name.LocalName)))
                Uncertain("UnsupportedResourceScope", entry, doc.Root);
            ValidateShapes(entry, doc.Root);
            return doc;
        }
        private void ValidateShapes(string entry, XElement root)
        {
            // Only these containers own the finite fields. Unknown control shapes are coverage gaps, never empty success.
            var shapes = new Dictionary<string, string[]>(StringComparer.Ordinal)
            {
                ["OFD"] = new[] { "DocBody" }, ["DocBody"] = new[] { "DocInfo", "DocRoot", "Versions", "Signatures" },
                ["Document"] = new[] { "CommonData", "Pages", "Outlines", "Permissions", "Actions", "VPreferences", "Bookmarks", "Annotations", "CustomTags", "Attachments", "Extensions" },
                ["CommonData"] = new[] { "MaxUnitID", "PageArea", "PublicRes", "DocumentRes", "TemplatePage", "DefaultCS" },
                ["Pages"] = new[] { "Page" },
                ["Res"] = new[] { "Fonts", "MultiMedias", "ColorSpaces", "DrawParams", "CompositeGraphicUnits" },
                ["Fonts"] = new[] { "Font" }, ["Font"] = new[] { "FontFile" },
                ["MultiMedias"] = new[] { "MultiMedia" }, ["MultiMedia"] = new[] { "MediaFile" },
                ["Attachments"] = new[] { "Attachment" }, ["Attachment"] = new[] { "FileLoc" },
                ["Signatures"] = new[] { "Signature" }, ["Signature"] = new[] { "SignedInfo", "SignedValue" },
                ["SignedInfo"] = new[] { "Provider", "SignatureMethod", "SignatureDateTime", "References", "StampAnnot", "Seal" },
                ["References"] = new[] { "Reference" }
            };
            var unknown = root.DescendantsAndSelf().FirstOrDefault(node => node.HasElements &&
                shapes.TryGetValue(node.Name.LocalName, out var allowed) && node.Elements().Any(child =>
                    child.Name.Namespace != root.Name.Namespace || !allowed.Contains(child.Name.LocalName)));
            if (unknown != null) Uncertain("UnsupportedControlShape", entry, unknown);
            if (root.Name.LocalName == "Page")
            {
                var content = Children(root, "Content");
                var objectNames = new[] { "TextObject", "PathObject", "ImageObject", "CompositeObject" };
                foreach (var layer in content.SelectMany(node => Children(node, "Layer")))
                    if (layer.Elements().Any(node => node.Name.Namespace != root.Name.Namespace || !objectNames.Contains(node.Name.LocalName)))
                        Add(3, PreflightOutcome.Unknown, "UnsupportedPageObject", entry, layer);
                if (Descendants(root, "CompositeObject").Any()) Add(3, PreflightOutcome.Unknown, "UnsupportedCompositeFontScope", entry, root);
            }
        }
        private string? Target(int rule, string entry, XElement node, string? value, string? containing = null, bool nonemptyPayload = false)
        {
            Check();
            if (++references > options.MaxReferences) throw new BudgetExceededException();
            var attributeName = node.Name.LocalName == "Page" || node.Name.LocalName == "TemplatePage" || node.Name.LocalName == "Signature" ? "BaseLoc" :
                node.Name.LocalName == "Reference" ? "FileRef" : node.Name.LocalName == "Attachment" ? "FileLoc" : null;
            if (attributeName != null && node.Attributes().Any(x => x.Name.LocalName == attributeName && x.Name.NamespaceName.Length != 0))
            { Add(rule, PreflightOutcome.Unknown, "UnsupportedAttributeNamespace", entry, node); return null; }
            if (node.HasElements && (node.Name.LocalName == "DocRoot" || node.Name.LocalName == "PublicRes" || node.Name.LocalName == "DocumentRes" ||
                node.Name.LocalName == "FontFile" || node.Name.LocalName == "MediaFile" || node.Name.LocalName == "FileLoc" || node.Name.LocalName == "Attachments" || node.Name.LocalName == "Signatures" || node.Name.LocalName == "SignedValue"))
            { Add(rule, PreflightOutcome.Unknown, "UnsupportedReferenceShape", entry, node); return null; }
            if (string.IsNullOrWhiteSpace(value) && attributeName == "FileLoc" && node.Elements().Any(x => x.Name.LocalName == "FileLoc" && x.Name.Namespace != node.Name.Namespace))
            { Add(rule, PreflightOutcome.Unknown, "UnsupportedReferenceNamespace", entry, node); return null; }
            value = value?.Trim() ?? string.Empty;
            if (value.Length == 0) { Add(rule, PreflightOutcome.Fail, "EmptyReference", entry, node); return null; }
            if (value.Contains(":") || value.StartsWith("\\", StringComparison.Ordinal) || value.StartsWith("//", StringComparison.Ordinal))
            { Add(rule, PreflightOutcome.Fail, "ExternalReference", entry, node, "External location (redacted)"); return null; }
            if (value.IndexOfAny(new[] { '?', '#', '%', '\\' }) >= 0 || value.Any(char.IsControl))
            { Add(rule, PreflightOutcome.Unknown, "AmbiguousReference", entry, node, "Ambiguous location (redacted)"); return null; }
            string target;
            try { target = OfdPackagePath.Resolve(containing ?? entry, value); }
            catch (InvalidDataException) { Add(rule, PreflightOutcome.Fail, "RootEscape", entry, node); return null; }
            if (!package.TryGetBytes(target, out var payload))
            { Add(rule, PreflightOutcome.Fail, "MissingTarget", entry, node, "Target=" + target); return null; }
            if (nonemptyPayload && payload.Length == 0)
            { Add(rule, PreflightOutcome.Fail, "EmptyPayload", entry, node, "Target=" + target + "; length=0"); return null; }
            Add(rule, PreflightOutcome.Pass, "LocalPayloadPresent", entry, node, "Reference=" + value + "; containing=" + (containing ?? entry) + "; Target=" + target + "; length=" + payload.Length + Identity(node));
            return target;
        }
        internal void Inspect(OfdPackageArchive archive)
        {
            package = archive; entryCount = archive.EntryNames.Count;
            var encrypted = archive.Contains("Encryptions.xml");
            var orphanCrypto = archive.EntryNames.Any(x => x.StartsWith("PasswordCrypto/", StringComparison.OrdinalIgnoreCase));
            Add(4, !options.RequireNoEncryption ? PreflightOutcome.NotApplicable : encrypted ? PreflightOutcome.Fail :
                orphanCrypto ? PreflightOutcome.Unknown : PreflightOutcome.Pass,
                !options.RequireNoEncryption ? "PolicyDoesNotRequirePlainPackage" : encrypted ? "EncryptionDeclared" :
                orphanCrypto ? "OrphanEncryptionMarker" : "KnownEncryptionMarkersAbsent", encrypted ? "Encryptions.xml" : "<input>",
                observed: "Encryptions.xml=" + encrypted + "; PasswordCrypto/=" + orphanCrypto + "; entries=" + entryCount);
            if (encrypted)
            {
                foreach (var rule in new[] { 1, 2, 3, 5, 6 }) Add(rule, PreflightOutcome.Unknown, "EncryptionMayObscureContent", "Encryptions.xml");
                gaps.Add("Encrypted content was not decrypted or interpreted."); return;
            }
            if (orphanCrypto) gaps.Add("Orphan PasswordCrypto entries have unverified ownership.");
            if (!archive.Contains("OFD.xml"))
            {
                Add(1, options.RequireSingleDocument ? PreflightOutcome.Fail : PreflightOutcome.NotApplicable, "MissingOfdRoot", "OFD.xml");
                Uncertain("DocumentScopeUnavailable", "OFD.xml"); return;
            }
            var ofd = ReadXml("OFD.xml", 1, "OFD")?.Root;
            if (ofd == null) return;
            var bodies = Children(ofd, "DocBody").ToList(); documentCount = bodies.Count;
            Add(1, !options.RequireSingleDocument ? PreflightOutcome.NotApplicable : bodies.Count == 1 ? PreflightOutcome.Pass : bodies.Count == 0 && ofd.Elements().Any(x => x.Name.LocalName == "DocBody") ? PreflightOutcome.Unknown : PreflightOutcome.Fail,
                !options.RequireSingleDocument ? "PolicyAllowsMultipleDocuments" : "DocumentCount", "OFD.xml", ofd, "DocBody count=" + bodies.Count, "Exactly one recognized DocBody");
            if (bodies.Count == 0) Uncertain("DocumentScopeUnavailable", "OFD.xml", ofd);
            foreach (var body in bodies)
            {
                Check();
                var roots = Children(body, "DocRoot").ToList();
                if (roots.Count != 1)
                {
                    Add(1, options.RequireSingleDocument ? (roots.Count == 0 && body.Elements().Any(x => x.Name.LocalName == "DocRoot") ? PreflightOutcome.Unknown : PreflightOutcome.Fail) : PreflightOutcome.NotApplicable, "DocRootCount", "OFD.xml", body, "count=" + roots.Count);
                    Uncertain("AmbiguousDocumentOwnership", "OFD.xml", body);
                }
                foreach (var root in roots)
                {
                    var document = Target(1, "OFD.xml", root, root.Value, "");
                    if (document != null) InspectDocument(document);
                    else Uncertain("DocumentScopeUnavailable", "OFD.xml", root);
                }
                foreach (var signatures in Children(body, "Signatures")) InspectSignatures(signatures);
            }
            foreach (var entry in archive.EntryNames.OrderBy(x => x, StringComparer.Ordinal))
            {
                Check();
                if (entry.EndsWith("Attachments.xml", StringComparison.OrdinalIgnoreCase) && !attachmentEntries.Contains(entry))
                    Add(5, PreflightOutcome.Unknown, "OrphanAttachmentIndex", entry);
                if ((entry.EndsWith("Signatures.xml", StringComparison.OrdinalIgnoreCase) || entry.EndsWith("Signature.xml", StringComparison.OrdinalIgnoreCase) ||
                    entry.EndsWith("SignedValue.dat", StringComparison.OrdinalIgnoreCase)) && !signatureEntries.Contains(entry))
                    Add(6, PreflightOutcome.Unknown, "OrphanSignaturePayload", entry);
                if (new[] { "PublicRes.xml", "DocumentRes.xml", "PageRes.xml" }.Contains(entry.Substring(entry.LastIndexOf('/') + 1), StringComparer.OrdinalIgnoreCase) && !resourceEntries.Contains(entry))
                    Uncertain("OrphanResourceTable", entry);
            }
        }
        private void InspectDocument(string entry)
        {
            var doc = ReadXml(entry, 1, "Document")?.Root;
            if (doc == null) return;
            var fonts = new HashSet<string>(StringComparer.Ordinal);
            foreach (var common in Children(doc, "CommonData"))
                foreach (var res in common.Elements().Where(x => x.Name == doc.Name.Namespace + "PublicRes" || x.Name == doc.Name.Namespace + "DocumentRes"))
                {
                    var target = Target(2, entry, res, res.Value);
                    if (target != null) InspectResources(target, fonts);
                    else Add(3, PreflightOutcome.Unknown, "ResourceScopeUnavailable", entry, res);
                }
            var pages = Children(doc, "Pages").SelectMany(x => Children(x, "Page")).ToList();
            if (options.RequireSingleDocument) Add(1, pages.Count == 0 ? (doc.Elements().Any(x => x.Name.LocalName == "Pages" && x.Name.Namespace != doc.Name.Namespace) || Children(doc, "Pages").Any(x => x.Elements().Any(y => y.Name.LocalName == "Page" && y.Name.Namespace != doc.Name.Namespace)) ? PreflightOutcome.Unknown : PreflightOutcome.Fail) : PreflightOutcome.Pass, "PageCount", entry, doc, "count=" + pages.Count, "At least one existing page");
            foreach (var page in pages.Concat(Children(doc, "CommonData").SelectMany(x => Children(x, "TemplatePage"))))
            {
                var target = Target(2, entry, page, (string?)page.Attribute("BaseLoc"));
                if (target == null)
                {
                    if (options.RequireSingleDocument && page.Name.LocalName == "Page") Add(1, results[results.Count - 1].Outcome, "PageUnavailable", entry, page);
                    Add(3, PreflightOutcome.Unknown, "PageContentUnavailable", entry, page);
                    continue;
                }
                var localFonts = new HashSet<string>(fonts, StringComparer.Ordinal);
                var pageRes = OfdPackagePath.Resolve(target, "PageRes.xml");
                if (package.Contains(pageRes)) InspectResources(pageRes, localFonts);
                var content = ReadXml(target, 2, "Page")?.Root;
                if (content == null) { Add(3, PreflightOutcome.Unknown, "PageContentUnavailable", target); continue; }
                foreach (var text in Descendants(content, "TextObject"))
                {
                    var font = (string?)text.Attribute("Font");
                    if (font == null || !localFonts.Contains(font)) Add(3, PreflightOutcome.Unknown, "UnresolvedTextFont", target, text, "Font=" + font);
                }
            }
            foreach (var attachment in Children(doc, "Attachments")) InspectAttachments(entry, attachment);
        }
        private void InspectResources(string entry, HashSet<string> fonts)
        {
            resourceEntries.Add(entry);
            var res = ReadXml(entry, 2, "Res")?.Root;
            if (res == null) { Add(3, PreflightOutcome.Unknown, "ResourceScopeUnavailable", entry); return; }
            var baseLoc = (string?)res.Attribute("BaseLoc") ?? string.Empty;
            var foreignBase = res.Attributes().Any(x => x.Name.LocalName == "BaseLoc" && x.Name.NamespaceName.Length != 0);
            // A synthetic containing entry lets the shared resolver apply resource BaseLoc before a relative payload.
            string? resourceBase = null;
            if (foreignBase) Add(2, PreflightOutcome.Unknown, "UnsupportedAttributeNamespace", entry, res);
            else if (baseLoc.Contains(":") || baseLoc.StartsWith("\\", StringComparison.Ordinal) || baseLoc.StartsWith("//", StringComparison.Ordinal))
                Add(2, PreflightOutcome.Fail, "ExternalResourceBase", entry, res);
            else if (baseLoc.IndexOfAny(new[] { '?', '#', '%', '\\' }) >= 0) Add(2, PreflightOutcome.Unknown, "AmbiguousResourceBase", entry, res);
            else
            {
                try { resourceBase = OfdPackagePath.Resolve(entry, (baseLoc.Length == 0 ? "_payload_" : baseLoc.TrimEnd('/') + "/_payload_")); }
                catch (InvalidDataException) { Add(2, PreflightOutcome.Fail, "RootEscape", entry, res); }
            }
            var declarations = Children(res, "Fonts").SelectMany(x => Children(x, "Font")).ToList();
            foreach (var font in declarations)
            {
                var id = (string?)font.Attribute("ID");
                if (string.IsNullOrWhiteSpace(id) || !fonts.Add(id!)) Add(3, PreflightOutcome.Unknown, "AmbiguousFontId", entry, font, "Font ID=" + id);
                var files = Children(font, "FontFile").ToList();
                if (files.Count != 1)
                {
                    var unknownFile = files.Count > 1 || font.Elements().Any(x => x.Name.LocalName == "FontFile" && x.Name.Namespace != font.Name.Namespace);
                    Add(3, unknownFile ? PreflightOutcome.Unknown : PreflightOutcome.Fail, unknownFile ? "AmbiguousFontFile" : "FontPayloadNotDeclared", entry, font);
                }
                foreach (var file in files)
                {
                    if (resourceBase == null && !file.Value.StartsWith("/", StringComparison.Ordinal))
                    { Add(3, PreflightOutcome.Unknown, "ResourceBaseUnavailable", entry, file); continue; }
                    Target(2, entry, file, file.Value, resourceBase ?? entry);
                    Target(3, entry, file, file.Value, resourceBase ?? entry, true);
                }
            }
            foreach (var media in Children(res, "MultiMedias").SelectMany(x => Children(x, "MultiMedia")))
            {
                var files = Children(media, "MediaFile").ToList();
                if (files.Count == 0) Add(2, media.Elements().Any(x => x.Name.LocalName == "MediaFile" && x.Name.Namespace != media.Name.Namespace) ? PreflightOutcome.Unknown : PreflightOutcome.Fail, "MediaPayloadNotDeclared", entry, media);
                if (files.Count > 1) Add(2, PreflightOutcome.Unknown, "AmbiguousMediaFile", entry, media);
                foreach (var file in files)
                    if (resourceBase != null || file.Value.StartsWith("/", StringComparison.Ordinal)) Target(2, entry, file, file.Value, resourceBase ?? entry);
                    else Add(2, PreflightOutcome.Unknown, "ResourceBaseUnavailable", entry, file);
            }
        }
        private void InspectAttachments(string documentEntry, XElement declaration)
        {
            var target = Target(5, documentEntry, declaration, declaration.Value);
            if (target == null) return;
            attachmentEntries.Add(target);
            var index = ReadXml(target, 5, "Attachments")?.Root;
            if (index == null) return;
            var attachments = Children(index, "Attachment").ToList();
            if (attachments.Count == 0) Add(5, PreflightOutcome.Unknown, "EmptyAttachmentIndex", target, index);
            foreach (var attachment in attachments)
            {
                var external = (string?)attachment.Attribute("External");
                if (external != null && external != "0" && external != "1" && !string.Equals(external, "false", StringComparison.OrdinalIgnoreCase) && !string.Equals(external, "true", StringComparison.OrdinalIgnoreCase))
                    Add(5, PreflightOutcome.Unknown, "AmbiguousExternalFlag", target, attachment);
                if (string.Equals((string?)attachment.Attribute("External"), "true", StringComparison.OrdinalIgnoreCase) || (string?)attachment.Attribute("External") == "1")
                    Add(5, PreflightOutcome.Fail, "ExternalAttachment", target, attachment);
                var fields = Children(attachment, "FileLoc").ToList();
                var attribute = (string?)attachment.Attribute("FileLoc");
                if (fields.Count > 1 || (attribute != null && fields.Count > 0))
                { Add(5, PreflightOutcome.Unknown, "AmbiguousFileLoc", target, attachment); continue; }
                Target(5, target, fields.FirstOrDefault() ?? attachment, fields.FirstOrDefault()?.Value ?? attribute);
            }
        }
        private void InspectSignatures(XElement declaration)
        {
            var listEntry = Target(6, "OFD.xml", declaration, declaration.Value, "");
            if (listEntry == null) return;
            signatureEntries.Add(listEntry);
            var list = ReadXml(listEntry, 6, "Signatures", true)?.Root;
            if (list == null) return;
            var signatures = Children(list, "Signature").ToList();
            if (signatures.Count == 0) Add(6, list.Elements().Any(x => x.Name.LocalName == "Signature" && x.Name.Namespace != list.Name.Namespace) ? PreflightOutcome.Unknown : PreflightOutcome.Fail, "EmptySignatureList", listEntry, list);
            foreach (var signature in signatures)
            {
                var entry = Target(6, listEntry, signature, (string?)signature.Attribute("BaseLoc"));
                if (entry == null) continue;
                signatureEntries.Add(entry);
                var root = ReadXml(entry, 6, "Signature", true)?.Root;
                if (root == null) continue;
                var values = Children(root, "SignedValue").ToList();
                if (values.Count != 1) Add(6, PreflightOutcome.Unknown, "SignedValueCount", entry, root);
                if (values.Count == 0 && !root.Elements().Any(x => x.Name.LocalName == "SignedValue" && x.Name.Namespace != root.Name.Namespace)) Add(6, PreflightOutcome.Fail, "MissingSignedValue", entry, root);
                foreach (var value in values)
                {
                    var payload = Target(6, entry, value, value.Value);
                    if (payload != null) signatureEntries.Add(payload);
                }
                var referenceLists = Children(root, "SignedInfo").SelectMany(x => Children(x, "References")).ToList();
                if (referenceLists.Count != 1) Add(6, PreflightOutcome.Unknown, "ReferenceScopeUnavailable", entry, root);
                foreach (var refs in referenceLists)
                {
                    var declared = Children(refs, "Reference").ToList();
                    if (declared.Count == 0) Add(6, PreflightOutcome.Unknown, "EmptyReferenceSet", entry, refs);
                    foreach (var reference in declared)
                    {
                        if (reference.Elements().Any(x => x.Name != root.Name.Namespace + "CheckValue") || reference.Attributes().Any(x => !x.IsNamespaceDeclaration && x.Name != "FileRef"))
                            Add(6, PreflightOutcome.Unknown, "UnsupportedReferenceTransform", entry, reference);
                        Target(6, entry, reference, (string?)reference.Attribute("FileRef"), "");
                    }
                }
            }
        }
        internal void Interrupt(string reason)
        {
            state = PreflightRunState.Incomplete; gaps.Add(reason);
            // Reserved slots preserve every uncompleted rule, even when the result budget itself is exhausted.
            foreach (var rule in Enumerable.Range(1, 6))
                results.Add(new PreflightRuleResult(Rule(rule), PreflightOutcome.Unknown, reason,
                    new PreflightEvidence("<input>", "<package>", "Inspection stopped", "Complete bounded inspection")));
        }
        internal ArchivalPreflightReport Finish()
        {
            foreach (var rule in Enumerable.Range(1, 6))
                if (!results.Any(x => x.RuleId == Rule(rule)))
                    results.Add(new PreflightRuleResult(Rule(rule), scopeUnknown ? PreflightOutcome.Unknown : PreflightOutcome.NotApplicable,
                        scopeUnknown ? "ScopeUnavailable" : "NoApplicableDeclarations",
                        new PreflightEvidence("<input>", "<package>", "No recognized declarations", "Complete known object scope")));
            return new ArchivalPreflightReport(state, results, gaps, entryCount, documentCount, options);
        }
    }
}
