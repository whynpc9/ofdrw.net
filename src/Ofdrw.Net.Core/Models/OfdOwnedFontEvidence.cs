using System;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal sealed class OfdOwnedFontDeclaration
{
    internal sealed class Facts
    {
        internal string DeclarationOriginEntry { get; }
        internal string DeclarationKind { get; }
        internal string DeclarationValue { get; }
        internal string? RawLinkValue { get; }
        internal string? DeclarationExpandedName { get; }
        internal bool DeclarationLinkPresent { get; }
        internal string ResourceEntry { get; }
        internal string FontEntry { get; }
        internal string FontExpandedName { get; }
        internal int FontOrdinalWithinFonts { get; }
        internal int FontElementPreorderOrdinal { get; }
        internal string? FontFileRaw { get; }
        internal string? ResBaseLocRaw { get; }
        internal Facts(string declarationOriginEntry, string declarationKind, string declarationValue, string? rawLinkValue, string? declarationExpandedName, bool declarationLinkPresent, string resourceEntry, string fontEntry, string fontExpandedName, int fontOrdinalWithinFonts, int fontElementPreorderOrdinal, string? fontFileRaw, string? resBaseLocRaw)
        {
            DeclarationOriginEntry = declarationOriginEntry;
            DeclarationKind = declarationKind;
            DeclarationValue = declarationValue;
            RawLinkValue = rawLinkValue;
            DeclarationExpandedName = declarationExpandedName;
            DeclarationLinkPresent = declarationLinkPresent;
            ResourceEntry = resourceEntry;
            FontEntry = fontEntry;
            FontExpandedName = fontExpandedName;
            FontOrdinalWithinFonts = fontOrdinalWithinFonts;
            FontElementPreorderOrdinal = fontElementPreorderOrdinal;
            FontFileRaw = fontFileRaw;
            ResBaseLocRaw = resBaseLocRaw;
        }
    }
    private readonly Facts source, output;
    internal string? SourceFontResourceId { get; }
    internal string OutputFontResourceId { get; }
    private OfdOwnedFontDeclaration(string sourceId, string outputId, Facts source, Facts output)
    { SourceFontResourceId = sourceId; OutputFontResourceId = outputId; this.source = source; this.output = output; }
    internal static OfdOwnedFontDeclaration FromActualPlan(string sourceId, string outputId, Facts source, Facts output)
        => new(sourceId, outputId, source, output);
    internal string? SourceDeclarationOriginEntry => source.DeclarationOriginEntry;
    internal string? SourceDeclarationKind => source.DeclarationKind;
    internal string? SourceDeclarationValue => source.DeclarationValue;
    internal string? SourceRawLinkValue => source.RawLinkValue;
    internal string? SourceDeclarationExpandedName => source.DeclarationExpandedName;
    internal bool SourceDeclarationLinkPresent => source.DeclarationLinkPresent;
    internal string? SourceResourceEntry => source.ResourceEntry;
    internal string? SourceFontEntry => source.FontEntry;
    internal string? SourceFontExpandedName => source.FontExpandedName;
    internal int? SourceFontOrdinalWithinFonts => source.FontOrdinalWithinFonts;
    internal int? SourceFontElementPreorderOrdinal => source.FontElementPreorderOrdinal;
    internal string? SourceFontFileRaw => source.FontFileRaw;
    internal string? SourceResBaseLocRaw => source.ResBaseLocRaw;
    internal string OutputDeclarationOriginEntry => output.DeclarationOriginEntry;
    internal string OutputDeclarationKind => output.DeclarationKind;
    internal string OutputDeclarationValue => output.DeclarationValue;
    internal string? OutputRawLinkValue => output.RawLinkValue;
    internal string? OutputDeclarationExpandedName => output.DeclarationExpandedName;
    internal bool OutputDeclarationLinkPresent => output.DeclarationLinkPresent;
    internal string OutputResourceEntry => output.ResourceEntry;
    internal string OutputFontEntry => output.FontEntry;
    internal string OutputFontExpandedName => output.FontExpandedName;
    internal int OutputFontOrdinalWithinFonts => output.FontOrdinalWithinFonts;
    internal int OutputFontElementPreorderOrdinal => output.FontElementPreorderOrdinal;
    internal string? OutputFontFileRaw => output.FontFileRaw;
    internal string? OutputResBaseLocRaw => output.ResBaseLocRaw;
}

internal sealed class OfdOwnedVerifiedFontSummary
{
    internal string Sha256Hex { get; }
    internal long ByteLength { get; }
    internal string ResolvedPayloadSha256Hex { get; }
    internal long ResolvedPayloadByteLength { get; }
    internal string ContentProfile => "raw-embedded-font-bytes/sha256/v1";
    internal string FontName { get; }
    internal string? FamilyName { get; }
    internal string? Charset { get; }
    internal bool Bold { get; }
    internal bool Italic { get; }
    internal string? ModelFileName { get; }
    internal OfdOwnedFontDeclaration Declaration { get; }
    private OfdOwnedVerifiedFontSummary(OfdFontResource actual, OfdObservedPayloadIdentity model,
        OfdObservedPayloadIdentity payload, OfdOwnedFontDeclaration declaration)
    {
        Sha256Hex = model.Sha256Hex; ByteLength = model.ByteLength;
        ResolvedPayloadSha256Hex = payload.Sha256Hex; ResolvedPayloadByteLength = payload.ByteLength;
        FontName = actual.FontName; FamilyName = actual.FamilyName; Charset = actual.Charset;
        Bold = actual.Bold; Italic = actual.Italic; ModelFileName = actual.FileName; Declaration = declaration;
    }
    internal static OfdOwnedVerifiedFontSummary FromActualCompoundPass(OfdFontResource actual,
        byte[] actualResolvedPayload, OfdDocumentPackage actualR, OfdControlledEditContext context,
        OfdCurrentPayloadEvidence current, OfdOwnedFontDeclaration declaration, CancellationToken token)
    {
        context.Material.ChargeCopy(1024, token);
        var model = current.RequireObserved(actualR, context, actual.Data, token);
        var payload = current.RequireObserved(actualR, context, actualResolvedPayload, token);
        return new(actual, model, payload, declaration);
    }
}
