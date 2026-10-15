using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal sealed class OfdReadbackRootAddress
{
    internal OfdOwnedTextRootKey Key { get; }
    internal int PageIndex { get; }
    internal string EntryPath { get; }
    internal int LayerOrdinal { get; }
    internal int RootOrdinal { get; }
    internal string OutputObjectId { get; }
    internal string OutputPageId { get; }
    internal string? OutputLayerId { get; }
    internal string OutputFontResourceId { get; }
    internal OfdOwnedFontDeclaration FontDeclaration { get; }
    internal string? SourcePageBaseLocRaw { get; }
    internal string? OutputPageBaseLocRaw { get; }
    internal string SourcePageReferenceExpandedName { get; }
    internal string OutputPageReferenceExpandedName { get; }
    internal string SourcePageReferenceOwnerEntry { get; }
    internal string OutputPageReferenceOwnerEntry { get; }
    private OfdReadbackRootAddress(OfdOwnedTextRootKey key, int pageIndex, string entryPath, int layerOrdinal,
        int rootOrdinal, string objectId, string pageId, string? layerId, string fontId, OfdOwnedFontDeclaration font,
        string? sourceBaseLoc, string? outputBaseLoc, string sourcePageQName, string outputPageQName, string sourceDoc, string outputDoc)
    {
        Key = key; PageIndex = pageIndex; EntryPath = entryPath; LayerOrdinal = layerOrdinal; RootOrdinal = rootOrdinal;
        OutputObjectId = objectId; OutputPageId = pageId; OutputLayerId = layerId; OutputFontResourceId = fontId; FontDeclaration = font;
        SourcePageBaseLocRaw = sourceBaseLoc; OutputPageBaseLocRaw = outputBaseLoc;
        SourcePageReferenceExpandedName = sourcePageQName; OutputPageReferenceExpandedName = outputPageQName;
        SourcePageReferenceOwnerEntry = sourceDoc; OutputPageReferenceOwnerEntry = outputDoc;
    }
    internal static OfdReadbackRootAddress FromActualPlan(OfdOwnedTextRootKey key, int pageIndex, string entryPath,
        int layerOrdinal, int rootOrdinal, string objectId, string pageId, string? layerId, string fontId, OfdOwnedFontDeclaration font,
        string? sourceBaseLoc, string? outputBaseLoc, string sourcePageQName, string outputPageQName, string sourceDoc, string outputDoc)
        => new(key, pageIndex, entryPath, layerOrdinal, rootOrdinal, objectId, pageId, layerId, fontId, font,
            sourceBaseLoc, outputBaseLoc, sourcePageQName, outputPageQName, sourceDoc, outputDoc);
}
