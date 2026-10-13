using System.Collections.Generic;

namespace Ofdrw.Net.Core.Models;

internal enum OfdTraversalDiagnosticKind { RawObject, UnsupportedSource, SharedDefinition, DerivedProjection }

internal sealed class OfdTraversalDiagnostic
{
    internal OfdTraversalDiagnostic(OfdTraversalDiagnosticKind kind, OfdPage page,
        OfdContentScope scope, IReadOnlyList<int> path, string? entryPath, string message)
    {
        Kind = kind; Page = page; RootScope = scope; OccurrencePath = path;
        SourceEntryPath = entryPath; Message = message;
    }
    internal OfdTraversalDiagnosticKind Kind { get; }
    internal OfdPage Page { get; }
    internal OfdContentScope RootScope { get; }
    internal IReadOnlyList<int> OccurrencePath { get; }
    internal string? SourceEntryPath { get; }
    internal string Message { get; }
}

internal sealed class OfdGraphicTraversalResult
{
    internal OfdGraphicTraversalResult(List<OfdGraphicOccurrence> occurrences,
        List<OfdTraversalDiagnostic> diagnostics, OfdTraversalBudget budget)
    {
        Occurrences = occurrences.AsReadOnly(); Diagnostics = diagnostics.AsReadOnly();
        NodeCount = budget.Nodes; ReferenceCount = budget.References;
        SourceXmlBytes = budget.SourceXmlBytes; PathValueCount = budget.PathValues;
    }
    internal IReadOnlyList<OfdGraphicOccurrence> Occurrences { get; }
    internal IReadOnlyList<OfdTraversalDiagnostic> Diagnostics { get; }
    internal long NodeCount { get; }
    internal long ReferenceCount { get; }
    internal long SourceXmlBytes { get; }
    internal long PathValueCount { get; }
}
