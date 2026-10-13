using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Models;

// A0 traverses the current model only. Raw hierarchy is not reparsed into a
// second object model. Existing reader/writer/clone entry points are unchanged.
internal static class OfdModelTraversal
{
    internal static OfdGraphicTraversalResult ReadGraphics(OfdDocumentPackage package,
        IReadOnlyList<OfdPage> selectedPages, OfdTraversalOptions options, CancellationToken cancellationToken = default)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (selectedPages is null) throw new ArgumentNullException(nameof(selectedPages));
        if (options is null) throw new ArgumentNullException(nameof(options));
        cancellationToken.ThrowIfCancellationRequested();
        var snapshot = options.Snapshot();
        var state = Preflight(package, snapshot.Limits, cancellationToken);
        if (selectedPages.Count > package.Pages.Count) throw new ArgumentException("Selected pages must be distinct package members.", nameof(selectedPages));
        var selected = new HashSet<OfdPage>(Identity<OfdPage>.Instance);
        var selectionOrder = new List<OfdPage>();
        foreach (var page in selectedPages)
        {
            cancellationToken.ThrowIfCancellationRequested();
            state.Budget.Reference();
            if (page is null || !state.Pages.ContainsKey(page) || !selected.Add(page))
                throw new ArgumentException("Selected pages must be distinct package members.", nameof(selectedPages));
            selectionOrder.Add(page);
        }

        var occurrences = new List<OfdGraphicOccurrence>();
        var diagnostics = new List<OfdTraversalDiagnostic>();
        foreach (var page in selectionOrder)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var ordinal = state.Pages[page];
            if (snapshot.IncludeTemplates)
                for (var i = 0; i < page.Templates.Count; i++)
                    if (IsBackground(page.Templates[i]) && page.Templates[i].Elements.Count > 0) Emit(page.Templates[i].Elements, page, ordinal, OfdContentScope.Template, i, TemplatePath(page.Templates[i]));
            Emit(page.Elements, page, ordinal, OfdContentScope.Body, 0, page.SourceEntryPath);
            if (snapshot.IncludeTemplates)
                for (var i = 0; i < page.Templates.Count; i++)
                    if (!IsBackground(page.Templates[i]) && page.Templates[i].Elements.Count > 0) Emit(page.Templates[i].Elements, page, ordinal, OfdContentScope.Template, i, TemplatePath(page.Templates[i]));
            if (snapshot.IncludeAnnotations)
                Emit(page.AnnotationAppearances, page, ordinal, OfdContentScope.Annotation, 0, null);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new OfdGraphicTraversalResult(occurrences, diagnostics, state.Budget);

        string? TemplatePath(OfdTemplateContent template)
        {
            var reference = template.BaseLocation;
            if (reference is null) return null;
            if (state.TemplatePaths.TryGetValue(reference, out var existing)) return existing;
            var documentPath = state.DocumentPath;
            var rootRelative = reference.StartsWith("/", StringComparison.Ordinal) || reference.StartsWith("\\", StringComparison.Ordinal);
            var directoryLength = rootRelative ? 0 : Math.Max(documentPath.LastIndexOf('/'), documentPath.LastIndexOf('\\')) + 1;
            // Upper-bound retained resolved text BEFORE the shared resolver's
            // string/segment allocations. Cache within this operation only.
            state.Budget.SourcePathCharacters((long)directoryLength + reference.Length + 1);
            ChargeSegments(reference);
            if (!rootRelative) ChargeSegments(documentPath);
            var resolved = OfdPackagePath.Resolve(documentPath, reference);
            state.TemplatePaths.Add(reference, resolved);
            return resolved;
        }

        void ChargeSegments(string value)
        {
            for (var i = 0; i < value.Length; i++)
            {
                if ((i & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                if (value[i] is '/' or '\\') state.Budget.Reference();
            }
        }

        void Emit(List<OfdElement> elements, OfdPage page, int pageOrdinal, OfdContentScope scope, int container, string? entryPath)
        {
            for (var index = 0; index < elements.Count; index++)
            {
                cancellationToken.ThrowIfCancellationRequested();
                var element = elements[index];
                var info = state.Elements[element];
                if (snapshot.VisibleOnly && !info.Visible) continue;
                var shared = info.Uses > 1 || scope == OfdContentScope.Template;
                var kind = !info.Primitive ? OfdTraversalDiagnosticKind.RawObject :
                    !info.Supported ? OfdTraversalDiagnosticKind.UnsupportedSource :
                    scope == OfdContentScope.Annotation ? OfdTraversalDiagnosticKind.DerivedProjection :
                    shared ? OfdTraversalDiagnosticKind.SharedDefinition : (OfdTraversalDiagnosticKind?)null;
                if (kind is OfdTraversalDiagnosticKind.RawObject or OfdTraversalDiagnosticKind.UnsupportedSource && snapshot.RequireSupportedSource)
                    throw new NotSupportedException("Selected OFD graphics contain unsupported source; no traversal result was returned.");
                IReadOnlyList<int>? path = null;
                if (kind.HasValue)
                {
                    state.Budget.Result(diagnostic: true);
                    path = Path(state.Budget, pageOrdinal, scope, container, index);
                    diagnostics.Add(new OfdTraversalDiagnostic(kind.Value, page, scope, path, entryPath, Message(kind.Value)));
                }
                if (!info.Primitive) continue;
                state.Budget.Result(diagnostic: false);
                path ??= Path(state.Budget, pageOrdinal, scope, container, index);
                occurrences.Add(new OfdGraphicOccurrence(element, page, scope, info.DefinitionPath, path,
                    shared, info.Supported, entryPath, info.SourceXml));
            }
        }
    }

    internal static void ValidateDefinitions(OfdDocumentPackage package, OfdModelLimits limits, CancellationToken cancellationToken = default)
    {
        if (package is null) throw new ArgumentNullException(nameof(package));
        if (limits is null) throw new ArgumentNullException(nameof(limits));
        cancellationToken.ThrowIfCancellationRequested();
        _ = Preflight(package, limits.Snapshot(), cancellationToken);
    }

    private static State Preflight(OfdDocumentPackage package, OfdModelLimits limits, CancellationToken token)
    {
        var state = new State(limits);
        state.Budget.Node(1);
        if (package.Options is null) throw new InvalidDataException("OFD model has no document options.");
        if (package.DocumentEntryPath is not null) GuardPath(package.DocumentEntryPath);
        else { GuardPath(package.Options.DocumentId); GuardPath("/"); GuardPath("Document.xml"); }
        state.DocumentPath = package.DocumentEntryPath ?? package.Options.DocumentId + "/Document.xml";
        for (var p = 0; p < package.Pages.Count; p++)
        {
            token.ThrowIfCancellationRequested();
            state.Budget.Reference(); state.Budget.Node(2);
            var page = package.Pages[p] ?? throw new InvalidDataException("OFD model contains a null page.");
            if (state.Pages.ContainsKey(page)) throw new InvalidDataException("OFD model contains duplicate page ownership.");
            state.Pages.Add(page, p);
            GuardPath(page.SourceEntryPath);
            Scan(page.Elements, p, OfdContentScope.Body, 0);
            for (var t = 0; t < page.Templates.Count; t++)
            {
                token.ThrowIfCancellationRequested();
                state.Budget.Reference(); state.Budget.Node(3);
                var template = page.Templates[t] ?? throw new InvalidDataException("OFD model contains a null template.");
                if (template.BaseLocation is not null) state.Budget.Reference();
                GuardPath(template.BaseLocation);
                Scan(template.Elements, p, OfdContentScope.Template, t);
            }
            Scan(page.AnnotationAppearances, p, OfdContentScope.Annotation, 0);
            ScanPreserved(page.PreservedPageElements);
        }
        ScanPreserved(package.PreservedCommonDataElements);
        ScanPreserved(package.PreservedDocumentElements);
        ScanPreserved(package.PreservedDocBodyElements);
        foreach (var font in package.Fonts) { token.ThrowIfCancellationRequested(); state.Budget.Reference(); state.Budget.Node(2); }
        foreach (var attachment in package.Attachments) { token.ThrowIfCancellationRequested(); state.Budget.Reference(); state.Budget.Node(2); }
        foreach (var entry in package.PreservedEntries) { token.ThrowIfCancellationRequested(); state.Budget.Reference(); state.Budget.Node(2); }
        foreach (var tag in package.CustomTags) { token.ThrowIfCancellationRequested(); state.Budget.Reference(); state.Budget.Node(2); }
        token.ThrowIfCancellationRequested();
        return state;

        void GuardPath(string? value)
        {
            token.ThrowIfCancellationRequested(); state.Budget.SourcePath(value);
            if (value is null) return;
            for (var i = 0; i < value.Length; i++)
            {
                if ((i & 1023) == 0) token.ThrowIfCancellationRequested();
                if (value[i] is '/' or '\\') state.Budget.Reference();
            }
        }

        void Scan(List<OfdElement> elements, int pageOrdinal, OfdContentScope scope, int container)
        {
            state.Budget.Node(3);
            for (var index = 0; index < elements.Count; index++)
            {
                token.ThrowIfCancellationRequested();
                state.Budget.Reference(); state.Budget.Node(4);
                var element = elements[index] ?? throw new InvalidDataException("OFD model contains a null graphic.");
                if (state.Elements.TryGetValue(element, out var existing)) { existing.Uses++; continue; }
                var source = Source(element);
                var primitive = element is OfdTextElement or OfdImageElement or OfdPathElement;
                var supported = primitive;
                var visible = true;
                if (source is not null)
                {
                    GuardXml(source, state.Budget, limits, token);
                    if (primitive && !string.IsNullOrWhiteSpace(source))
                    {
                        using var reader = XmlReader.Create(new StringReader(source), Settings(limits));
                        var xml = XElement.Load(reader, LoadOptions.PreserveWhitespace);
                        var expected = element is OfdTextElement ? "TextObject" : element is OfdImageElement ? "ImageObject" : "PathObject";
                        var knownRoot = xml.Name.LocalName == expected && (xml.Name.Namespace == XNamespace.None ||
                            xml.Name.NamespaceName == OfdConstants.Namespace || xml.Name.NamespaceName == OfdConstants.StandardNamespace);
                        supported &= knownRoot && OfdGraphicXmlContract.HasKnownChildren(xml) &&
                            xml.DescendantsAndSelf().Attributes().All(OfdGraphicXmlContract.IsKnownAttribute) && !OfdGraphicXmlContract.HasUnsupportedReferences(xml);
                        // Foreign/wrong-root lookalikes cannot suppress their own
                        // unsupported-source diagnostic by claiming invisibility.
                        if (knownRoot) visible = xml.Attribute("Visible")?.Value.ToLowerInvariant() is not ("false" or "0");
                    }
                }
                if (element.ClippingXml is not null)
                {
                    GuardXml(element.ClippingXml, state.Budget, limits, token);
                    if (!string.IsNullOrWhiteSpace(element.ClippingXml))
                    {
                        using var reader = XmlReader.Create(new StringReader(element.ClippingXml), Settings(limits));
                        var clip = XElement.Load(reader, LoadOptions.PreserveWhitespace);
                        supported &= clip.Name.LocalName == "Clips" && OfdGraphicXmlContract.HasKnownChildren(clip) &&
                            clip.DescendantsAndSelf().Attributes().All(OfdGraphicXmlContract.IsKnownAttribute) && !OfdGraphicXmlContract.HasUnsupportedReferences(clip);
                    }
                }
                if (element is OfdTextElement text)
                {
                    if (text.FontResourceId is not null) state.Budget.Reference();
                    foreach (var run in text.Runs) { token.ThrowIfCancellationRequested(); state.Budget.Reference(); state.Budget.Node(5); }
                }
                if (element is OfdImageElement image && !string.IsNullOrEmpty(image.ResourceId)) state.Budget.Reference();
                state.Elements.Add(element, new ElementInfo(primitive, supported, visible, source,
                    Path(state.Budget, pageOrdinal, scope, container, index)));
            }
        }

        void ScanPreserved(List<string> values)
        {
            foreach (var xml in values)
            {
                token.ThrowIfCancellationRequested(); state.Budget.Reference(); state.Budget.Node(3);
                if (xml is null) throw new InvalidDataException("OFD model contains null preserved XML.");
                GuardXml(xml, state.Budget, limits, token);
            }
        }
    }

    internal static void GuardXml(string xml, OfdTraversalBudget budget, OfdModelLimits limits, CancellationToken token)
    {
        // Count UTF-8 bytes without allocating a second full encoding buffer.
        long bytes = 0;
        var nextCancellationCheck = 0;
        for (var i = 0; i < xml.Length; i++)
        {
            CheckCancellationCheckpoint(i, ref nextCancellationCheck, token);
            var c = xml[i];
            if (c <= 0x7f) bytes++;
            else if (c <= 0x7ff) bytes += 2;
            else if (char.IsHighSurrogate(c) && i + 1 < xml.Length && char.IsLowSurrogate(xml[i + 1])) { bytes += 4; i++; }
            else bytes += 3;
            if (bytes > limits.MaxSourceXmlBytes - budget.SourceXmlBytes)
                throw OfdTraversalBudget.Limit("OFD traversal exceeds its source XML byte budget.");
        }
        token.ThrowIfCancellationRequested();
        budget.XmlBytes(bytes);
        if (string.IsNullOrWhiteSpace(xml)) return;
        try
        {
            using var reader = XmlReader.Create(new StringReader(xml), Settings(limits));
            while (reader.Read())
            {
                token.ThrowIfCancellationRequested(); budget.Node(reader.Depth + 1);
                for (var a = 0; a < reader.AttributeCount; a++) { token.ThrowIfCancellationRequested(); budget.Reference(); }
            }
            token.ThrowIfCancellationRequested();
        }
        catch (XmlException exception) { throw new InvalidDataException("OFD model contains malformed or prohibited source XML.", exception); }
    }

    // Advancing by Unicode scalars may cross a checkpoint without visiting its
    // exact index. A threshold also works for a final partial scan segment.
    internal static void CheckCancellationCheckpoint(int position, ref int nextCheck, CancellationToken token)
    {
        if (position < nextCheck) return;
        token.ThrowIfCancellationRequested();
        nextCheck = position > int.MaxValue - 1024 ? int.MaxValue : position + 1024;
    }

    private static XmlReaderSettings Settings(OfdModelLimits limits) => new()
    {
        DtdProcessing = DtdProcessing.Prohibit, XmlResolver = null,
        MaxCharactersInDocument = limits.MaxSourceXmlBytes
    };

    private static IReadOnlyList<int> Path(OfdTraversalBudget budget, int page, OfdContentScope scope, int container, int index)
    {
        budget.Path(4);
        return Array.AsReadOnly(new[] { page, (int)scope, container, index });
    }

    private static string? Source(OfdElement element) => element switch
    {
        OfdTextElement text => text.SourceXml, OfdImageElement image => image.SourceXml,
        OfdPathElement path => path.SourceXml, OfdRawElement raw => raw.Xml, _ => null
    };

    private static bool IsBackground(OfdTemplateContent template) => string.Equals(template.ZOrder, "Background", StringComparison.OrdinalIgnoreCase);
    private static string Message(OfdTraversalDiagnosticKind kind) => kind switch
    {
        OfdTraversalDiagnosticKind.RawObject => "Raw or unknown graphic hierarchy is not expanded.",
        OfdTraversalDiagnosticKind.UnsupportedSource => "Primitive source contains unsupported XML; geometry is not certified.",
        OfdTraversalDiagnosticKind.DerivedProjection => "Legacy annotation appearance is a derived page-space projection, not editable source.",
        _ => "Shared or template source does not support an occurrence-local edit."
    };

    private sealed class State(OfdModelLimits limits)
    {
        internal readonly OfdTraversalBudget Budget = new(limits);
        internal readonly Dictionary<OfdPage, int> Pages = new(Identity<OfdPage>.Instance);
        internal readonly Dictionary<OfdElement, ElementInfo> Elements = new(Identity<OfdElement>.Instance);
        internal readonly Dictionary<string, string> TemplatePaths = new(StringComparer.Ordinal);
        internal string DocumentPath = string.Empty;
    }

    private sealed class ElementInfo(bool primitive, bool supported, bool visible, string? sourceXml, IReadOnlyList<int> path)
    {
        internal readonly bool Primitive = primitive, Supported = supported, Visible = visible;
        internal readonly string? SourceXml = sourceXml;
        internal readonly IReadOnlyList<int> DefinitionPath = path;
        internal int Uses = 1;
    }

    private sealed class Identity<T> : IEqualityComparer<T> where T : class
    {
        internal static readonly Identity<T> Instance = new();
        public bool Equals(T? x, T? y) => ReferenceEquals(x, y);
        public int GetHashCode(T obj) => RuntimeHelpers.GetHashCode(obj);
    }
}
