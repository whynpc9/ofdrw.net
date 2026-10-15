using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

internal enum OfdContentScope { Body, Template, Annotation }
internal enum OfdSourceAuthority { ModelDefinition, SharedTemplateSource, DerivedAnnotationProjection }

internal sealed class OfdTraversalOptions
{
    internal bool IncludeTemplates { get; set; }
    internal bool IncludeAnnotations { get; set; }
    internal bool VisibleOnly { get; set; }
    internal bool RequireSupportedSource { get; set; }
    internal OfdModelLimits Limits { get; set; } = new();

    internal OfdTraversalOptions Snapshot() => new()
    {
        IncludeTemplates = IncludeTemplates, IncludeAnnotations = IncludeAnnotations,
        VisibleOnly = VisibleOnly, RequireSupportedSource = RequireSupportedSource,
        Limits = (Limits ?? throw new System.ArgumentException("Traversal limits are required.")).Snapshot()
    };
}

// Paths and source strings describe this operation's model snapshot only. This
// does not certify glyph geometry or permit concurrent mutation or stale edits.
internal sealed class OfdGraphicOccurrence
{
    internal OfdGraphicOccurrence(OfdElement element, OfdPage page, OfdContentScope scope,
        IReadOnlyList<int> definitionPath, IReadOnlyList<int> occurrencePath, bool shared,
        bool supported, string? entryPath, string? sourceXml)
    {
        Element = element; Page = page; RootScope = scope;
        DefinitionPath = definitionPath; OccurrencePath = occurrencePath;
        IsSharedDefinition = shared;
        SourceAuthority = scope == OfdContentScope.Template ? OfdSourceAuthority.SharedTemplateSource :
            scope == OfdContentScope.Annotation ? OfdSourceAuthority.DerivedAnnotationProjection : OfdSourceAuthority.ModelDefinition;
        IsEditableSource = SourceAuthority == OfdSourceAuthority.ModelDefinition && !shared && supported;
        SourceEntryPath = entryPath; SourceXml = sourceXml;
    }

    internal OfdElement Element { get; }
    internal OfdPage Page { get; }
    internal OfdContentScope RootScope { get; }
    internal OfdSourceAuthority SourceAuthority { get; }
    internal IReadOnlyList<int> DefinitionPath { get; }
    internal IReadOnlyList<int> OccurrencePath { get; }
    internal IReadOnlyList<double> ParentToPage { get; } = System.Array.AsReadOnly(new double[] { 1, 0, 0, 1, 0, 0 });
    internal bool IsInComposite => false; // Raw composites are diagnosed, never expanded in A0.
    internal bool IsEditableSource { get; }
    internal bool IsSharedDefinition { get; }
    internal string? SourceEntryPath { get; }
    internal string? SourceXml { get; }
}
