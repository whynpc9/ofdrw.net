using System;
using System.IO;

namespace Ofdrw.Net.Core.Models;

// Internal until a caller-facing read/write option needs this contract.
internal sealed class OfdModelLimits
{
    internal int MaxDepth { get; set; } = 64;
    internal int MaxNodeCount { get; set; } = 100_000;
    internal int MaxReferenceCount { get; set; } = 200_000;
    internal int MaxExpandedOccurrenceCount { get; set; } = 100_000;
    internal long MaxSourceXmlBytes { get; set; } = 128L * 1024 * 1024;
    internal long MaxMaterializedBytes { get; set; } = 512L * 1024 * 1024;

    internal OfdModelLimits Snapshot()
    {
        var copy = new OfdModelLimits
        {
            MaxDepth = MaxDepth, MaxNodeCount = MaxNodeCount,
            MaxReferenceCount = MaxReferenceCount,
            MaxExpandedOccurrenceCount = MaxExpandedOccurrenceCount,
            MaxSourceXmlBytes = MaxSourceXmlBytes, MaxMaterializedBytes = MaxMaterializedBytes
        };
        if (copy.MaxDepth <= 0 || copy.MaxNodeCount <= 0 || copy.MaxReferenceCount <= 0 ||
            copy.MaxExpandedOccurrenceCount <= 0 || copy.MaxSourceXmlBytes <= 0 || copy.MaxMaterializedBytes <= 0)
            throw new ArgumentException("Traversal limits must be positive.");
        return copy;
    }
}

internal sealed class OfdTraversalBudget(OfdModelLimits limits)
{
    internal long Nodes { get; private set; }
    internal long References { get; private set; }
    internal long SourceXmlBytes { get; private set; }
    internal long Occurrences { get; private set; }
    internal long Diagnostics { get; private set; }
    internal long PathValues { get; private set; }
    private long sourcePathCharacters;

    internal void Node(int depth)
    {
        Depth(depth);
        if (Nodes >= limits.MaxNodeCount) Fail("node");
        Nodes++;
    }

    internal void Depth(int depth)
    {
        if (depth > limits.MaxDepth) Fail("depth");
    }

    internal void Reference()
    {
        if (References >= limits.MaxReferenceCount) Fail("reference");
        References++;
    }

    internal void NodeWork(long count, int depth)
    {
        Depth(depth);
        if (count < 0 || count > limits.MaxNodeCount - Nodes) Fail("node");
        Nodes += count;
    }

    internal void XmlBytes(long bytes)
    {
        if (bytes < 0 || bytes > limits.MaxSourceXmlBytes - SourceXmlBytes) Fail("source XML byte");
        SourceXmlBytes += bytes;
    }

    internal void Result(bool diagnostic)
    {
        if (Occurrences + Diagnostics >= limits.MaxExpandedOccurrenceCount) Fail("occurrence/diagnostic");
        if (diagnostic) Diagnostics++; else Occurrences++;
    }

    internal void Path(int length)
    {
        Depth(length);
        var maximum = (long)limits.MaxExpandedOccurrenceCount * limits.MaxDepth * 2;
        if (length > maximum - PathValues) Fail("path");
        PathValues += length;
    }

    internal void SourcePath(string? value)
    {
        if (value is null) return;
        SourcePathCharacters(value.Length);
    }

    internal void SourcePathCharacters(long count)
    {
        // Bound both the text and the segment array the existing path resolver
        // may allocate. This is metadata work, separate from XML byte storage.
        if (count < 0 || count > limits.MaxSourceXmlBytes - sourcePathCharacters) Fail("source path character");
        sourcePathCharacters += count;
    }

    private static readonly object LimitMarker = new();
    internal static InvalidDataException Limit(string message)
    {
        var exception = new InvalidDataException(message); exception.Data[LimitMarker] = true; return exception;
    }
    internal static bool IsLimit(Exception exception) => exception is InvalidDataException && exception.Data.Contains(LimitMarker);
    private static void Fail(string kind) => throw Limit($"OFD traversal exceeds its {kind} budget.");
}
