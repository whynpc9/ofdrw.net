using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Keywords.Internal;

// Immutable source attributes preserve absence, expanded names and order.
// Issued only by the existing admitted DOM; never a public DTO or credential.
internal sealed class OfdKeywordInputAttribute(string ns, string name, string value)
{
    internal string Namespace { get; } = ns;
    internal string Name { get; } = name;
    internal string Value { get; } = value;
}

internal static class OfdKeywordInputValues
{
    internal static IReadOnlyList<OfdKeywordInputAttribute> Attributes(XElement node, OfdKeywordLocatorBudget budget)
    {
        var count = 0;
        foreach (var a in node.Attributes()) { budget.Field(); if (!a.IsNamespaceDeclaration) count++; }
        budget.Copy(checked(256L + 128L * count));
        var values = new OfdKeywordInputAttribute[count]; var index = 0;
        foreach (var a in node.Attributes())
        {
            budget.Field(); if (a.IsNamespaceDeclaration) continue;
            budget.Retain(a.Name.NamespaceName); budget.Retain(a.Name.LocalName); budget.Retain(a.Value);
            values[index++] = new OfdKeywordInputAttribute(a.Name.NamespaceName, a.Name.LocalName, a.Value);
        }
        return Array.AsReadOnly(values);
    }
}

internal sealed class OfdKeywordSlotInput(IReadOnlyList<OfdKeywordInputAttribute> attributes)
{
    internal IReadOnlyList<OfdKeywordInputAttribute> Attributes { get; } = attributes;
}
internal sealed class OfdKeywordRootInput(string ns, IReadOnlyList<OfdKeywordInputAttribute> attributes,
    IReadOnlyList<string> children, IReadOnlyList<OfdKeywordInputAttribute>? fill, IReadOnlyList<OfdKeywordSlotInput> slots)
{
    internal string Namespace { get; } = ns;
    internal IReadOnlyList<OfdKeywordInputAttribute> Attributes { get; } = attributes;
    internal IReadOnlyList<string> Children { get; } = children;
    internal IReadOnlyList<OfdKeywordInputAttribute>? Fill { get; } = fill;
    internal IReadOnlyList<OfdKeywordSlotInput> Slots { get; } = slots;
}
internal sealed class OfdKeywordLayerInput(string ns, IReadOnlyList<OfdKeywordInputAttribute> attributes, int objects)
{
    internal string Namespace { get; } = ns;
    internal IReadOnlyList<OfdKeywordInputAttribute> Attributes { get; } = attributes;
    internal int Objects { get; } = objects;
}
internal sealed class OfdKeywordPageInput(int position, string id, string entry, string ns, bool ownArea,
    string area, IReadOnlyList<OfdKeywordInputAttribute> reference, IReadOnlyList<OfdKeywordLayerInput> layers)
{
    internal int Position { get; } = position;
    internal string Id { get; } = id;
    internal string Entry { get; } = entry;
    internal string Namespace { get; } = ns;
    internal bool OwnArea { get; } = ownArea;
    internal string PhysicalBox { get; } = area;
    internal IReadOnlyList<OfdKeywordInputAttribute> Reference { get; } = reference;
    internal IReadOnlyList<OfdKeywordLayerInput> Layers { get; } = layers;
}
internal sealed class OfdKeywordFontInput(string ns, string resourceEntry, string entry,
    IReadOnlyList<OfdKeywordInputAttribute> resourceAttributes, IReadOnlyList<OfdKeywordInputAttribute> attributes,
    string fontFile, string? fontName, string? familyName, string? charset, bool bold, bool italic,
    int ordinal, int elementPreorder, string originEntry, string declarationKind, string declarationValue)
{
    internal string Namespace { get; } = ns;
    internal string ResourceEntry { get; } = resourceEntry;
    internal string Entry { get; } = entry;
    internal IReadOnlyList<OfdKeywordInputAttribute> ResourceAttributes { get; } = resourceAttributes;
    internal IReadOnlyList<OfdKeywordInputAttribute> Attributes { get; } = attributes;
    internal string FontFile { get; } = fontFile;
    internal string? FontName { get; } = fontName;
    internal string? FamilyName { get; } = familyName;
    internal string? Charset { get; } = charset;
    internal bool Bold { get; } = bold;
    internal bool Italic { get; } = italic;
    internal int Ordinal { get; } = ordinal;
    internal int ElementPreorder { get; } = elementPreorder;
    internal string OriginEntry { get; } = originEntry;
    internal string DeclarationKind { get; } = declarationKind;
    internal string DeclarationValue { get; } = declarationValue;
}

// This is an accounted temporary actual-readback capture, not an expectation,
// successful Writer receipt, provider qualification or publication capability.
internal sealed class OfdKeywordInputCapture : IDisposable
{
    private readonly OfdKeywordOperation operation;
    private readonly OfdKeywordSourceSnapshot snapshot;
    internal OfdKeywordLocatorBudget Budget { get; }
    internal OfdKeywordSourceAdmission Admission { get; }
    internal OfdGraphicTraversalResult Traversal { get; }
    internal IReadOnlyList<OfdKeywordTextSource> Sources { get; }
    private OfdKeywordInputCapture(OfdKeywordOperation operation, OfdKeywordSourceSnapshot snapshot,
        OfdKeywordLocatorBudget budget, OfdKeywordSourceAdmission admission,
        OfdGraphicTraversalResult traversal, OfdKeywordTextSource[] sources)
    {
        this.operation = operation; this.snapshot = snapshot; Budget = budget; Admission = admission;
        Traversal = traversal; Sources = Array.AsReadOnly(sources);
    }

    // Shared strict source kernel. Public analysis retains its existing admission
    // and projection behavior; the private entry adds complete typed evidence.
    internal static OfdKeywordTextSource CaptureSource(OfdGraphicOccurrence occurrence, int layer, string path,
        OfdKeywordLocatorBudget budget, int reservedUnits, OfdKeywordSourceAdmission admission)
        => OfdStrictTextGeometry.Capture(occurrence, layer, path, budget, reservedUnits, admission.HasOriginalEnvelope);

    internal static OfdKeywordInputCapture CaptureActual(OfdDocumentPackage source, int[] selected,
        OfdControlledEditContext context, CancellationToken token)
    {
        OfdKeywordOperation? operation = null; OfdKeywordSourceSnapshot? snapshot = null;
        try
        {
            context.AssertActive(); token.ThrowIfCancellationRequested();
            var limits = new OfdKeywordLocatorLimits();
            context.Material.ChargeCopy(checked(8192L + 64L * selected.Length), token);
            operation = new OfdKeywordOperation(new OfdKeywordSearchOptions(), token, context.Material);
            var budget = new OfdKeywordLocatorBudget(limits, operation);
            var pages = new OfdPage[selected.Length]; var seen = new HashSet<int>();
            for (var i = 0; i < selected.Length; i++)
            {
                budget.Check(); var index = selected[i];
                if (index < 0 || index >= source.Pages.Count || !seen.Add(index)) throw new InvalidDataException("PrivateInputPageSelection");
                pages[i] = source.Pages[index];
            }
            budget.Copy(checked(256L * limits.MaxNodes + 32L * limits.MaxReferences + 256L * limits.MaxOccurrences +
                12L * context.Material.Structure.SourceXmlBytes));
            var traversal = OfdModelTraversal.ReadGraphics(source, pages, new OfdTraversalOptions
            {
                IncludeTemplates = false, IncludeAnnotations = false, VisibleOnly = false, RequireSupportedSource = false,
                Limits = new OfdModelLimits { MaxNodeCount = limits.MaxNodes, MaxReferenceCount = limits.MaxReferences,
                    MaxExpandedOccurrenceCount = limits.MaxOccurrences, MaxDepth = limits.MaxDepth, MaxSourceXmlBytes = limits.MaxXmlBytes }
            }, token);
            if (traversal.Diagnostics.Any(d => d.Kind is OfdTraversalDiagnosticKind.RawObject or OfdTraversalDiagnosticKind.UnsupportedSource))
                throw new NotSupportedException("UnprovedPrivateInputNeighbors");
            snapshot = new OfdKeywordSourceSnapshot(source, budget);
            var admission = new OfdKeywordSourceAdmission(source, snapshot, budget); admission.PrepareDocument();
            budget.Copy(checked(1024L + 256L * selected.Length + 32L * traversal.Occurrences.Count));
            var certificates = new Dictionary<int, Dictionary<int, (int Layer, string Path)>>();
            foreach (var page in selected)
            {
                budget.Check();
                if (source.Pages[page].Templates.Count != 0 || source.Pages[page].AnnotationAppearances.Count != 0)
                    throw new NotSupportedException("UnprovedPrivateInputOverlays");
                certificates.Add(page, admission.Page(source.Pages[page], page));
            }
            var values = new OfdKeywordTextSource[traversal.Occurrences.Count]; var cursor = 0;
            foreach (var occurrence in traversal.Occurrences)
            {
                budget.Check();
                if (occurrence.Element is not OfdTextElement text || occurrence.IsInComposite || !occurrence.IsEditableSource ||
                    !certificates.TryGetValue(occurrence.OccurrencePath[0], out var certificate) ||
                    !certificate.TryGetValue(occurrence.OccurrencePath[3], out var location) || snapshot.HasSharedRun(text, budget))
                    throw new NotSupportedException("UnprovedPrivateInputNeighbors");
                long units = text.Text?.Length ?? 0;
                long runUnits = 0; foreach (var run in text.Runs) { budget.Check(); runUnits = checked(runUnits + (run.Text?.Length ?? 0)); }
                units = Math.Max(units, runUnits);
                if (units > int.MaxValue) throw new InvalidDataException("PrivateInputTextCapacity");
                budget.Text((int)units);
                var captured = OfdStrictTextGeometry.CapturePrivate(occurrence, location.Layer, location.Path, budget, (int)units, admission.HasOriginalEnvelope);
                if (!captured.Visible || !captured.Editable || !captured.HasXml) throw new NotSupportedException("UnprovedPrivateInputNeighbors");
                admission.CaptureFontInput(captured.FontId, text); values[cursor++] = captured;
            }
            snapshot.Seal(budget); snapshot.ValidateUnchanged(source, budget);
            context.AssertActive(); token.ThrowIfCancellationRequested();
            return new OfdKeywordInputCapture(operation, snapshot, budget, admission, traversal, values);
        }
        catch
        {
            context.Fail(); try { snapshot?.Dispose(); } catch { } try { operation?.Dispose(); } catch { } throw;
        }
    }
    public void Dispose() { snapshot.Dispose(); operation.Dispose(); }
}
