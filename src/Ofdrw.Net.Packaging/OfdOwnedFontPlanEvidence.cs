using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging;

internal static class OfdOwnedFontPlanEvidence
{
    internal static OfdOwnedFontDeclaration.Facts CaptureSource(OfdDocumentPackage package,
        OfdFontResource font, OfdModelMaterialization material, CancellationToken token)
    {
        if (package.SourceIdentity is null || !package.SourceIdentity.TryOwner(font, out var owner) ||
            owner.Kind.LocalName != "Font" || owner.Id != font.Id)
            throw new InvalidDataException("Private font has no actual Reader source owner.");
        material.ChargeCopy(checked(256L + 32L * package.Pages.Count), token);
        return Capture(package.DocumentEntryPath!, package.Pages.Select(page => page.SourceEntryPath!).ToArray(),
            package.PreservedEntries, owner.Entry, font.Id, owner.Slot, material, token);
    }
    internal static OfdOwnedFontDeclaration.Facts CaptureOutput(string documentEntry,
        string[] pageEntries, IReadOnlyDictionary<string, byte[]> entries, OfdOwnedFontDeclaration.Facts source,
        string fontId, OfdModelMaterialization material, CancellationToken token)
        => Capture(documentEntry, pageEntries, entries, source.ResourceEntry, fontId, source.FontElementPreorderOrdinal, material, token);

    private static OfdOwnedFontDeclaration.Facts Capture(string documentEntry, string[] pages,
        IReadOnlyDictionary<string, byte[]> entries, string resourceEntry, string fontId, int expectedSlot,
        OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(4096L + 256L * pages.Length, token);
        if (!entries.TryGetValue(documentEntry, out var docBytes) || !entries.TryGetValue(resourceEntry, out var resBytes))
            throw new InvalidDataException("Actual font declaration entries are absent.");
        var document = material.ReadXml(documentEntry, docBytes, token);
        var resources = material.ReadXml(resourceEntry, resBytes, token);
        var root = resources.Root ?? throw new InvalidDataException("Missing resource root.");
        var ns = root.Name.Namespace;
        if (root.Name != ns + "Res") throw new InvalidDataException("Unsupported font resource root.");
        var matches = root.Elements(ns + "Fonts").Elements(ns + "Font").Where(node => node.Attribute("ID")?.Value == fontId).ToArray();
        if (matches.Length != 1) throw new InvalidDataException("Font lacks a unique direct Fonts owner.");
        var font = matches[0]; var fullSlot = 0;
        foreach (var node in root.DescendantsAndSelf())
        {
            material.ChargeWork(16, token);
            if (ReferenceEquals(node, font)) break;
            fullSlot++;
        }
        if (fullSlot != expectedSlot) throw new InvalidDataException("Font preorder owner relocated outside the identity profile.");
        var ordinal = 0;
        foreach (var node in root.Elements(ns + "Fonts").Elements(ns + "Font"))
        { material.ChargeWork(16, token); if (ReferenceEquals(node, font)) break; ordinal++; }
        var fontFiles = font.Elements(ns + "FontFile").ToArray();
        if (fontFiles.Length != 1 || font.Attribute("FontFile") is not null || string.IsNullOrWhiteSpace(fontFiles[0].Value))
            throw new InvalidDataException("Private embedded font requires one actual FontFile child.");
        var raw = fontFiles[0].Value; var baseLoc = root.Attribute("BaseLoc")?.Value;
        material.ReservePathComposition(baseLoc ?? string.Empty, raw, token);
        var payload = material.ResolvePath(resourceEntry, string.IsNullOrEmpty(baseLoc) || raw.StartsWith("/", StringComparison.Ordinal)
            ? raw : baseLoc!.TrimEnd('/') + "/" + raw, token);
        if (!entries.ContainsKey(payload)) throw new InvalidDataException("Resolved actual font payload is absent.");
        string? origin = null, kind = null, value = null, qname = null; bool present = false;
        var docNs = document.Root!.Name.Namespace;
        foreach (var link in document.Root.Element(docNs + "CommonData")?.Elements() ?? Enumerable.Empty<XElement>())
        {
            if (link.Name != docNs + "PublicRes" && link.Name != docNs + "DocumentRes") continue;
            if (material.ResolvePath(documentEntry, link.Value, token) != resourceEntry) continue;
            if (origin is not null) throw new InvalidDataException("Font has ambiguous declaration origins.");
            origin = documentEntry; kind = link.Name.LocalName; value = link.Value; qname = link.Name.ToString(); present = true;
        }
        if (origin is null)
            foreach (var page in pages)
            {
                material.ReservePathComposition(page, "PageRes.xml", token);
                var fixedPath = material.ResolvePath(page, "PageRes.xml", token);
                if (fixedPath != resourceEntry) continue;
                if (origin is not null) throw new InvalidDataException("Font has ambiguous fixed PageRes origins.");
                origin = page; kind = "FixedPageRes"; value = "PageRes.xml";
            }
        if (origin is null) throw new InvalidDataException("Font resource origin is outside the supported document/page profile.");
        return new(origin, kind!, value!, present ? value : null, qname, present, resourceEntry, payload,
            font.Name.ToString(), ordinal, fullSlot, raw, baseLoc);
    }
    internal static void RequireIdentityMapping(OfdOwnedFontDeclaration.Facts source, OfdOwnedFontDeclaration.Facts output)
    {
        if (source.DeclarationOriginEntry != output.DeclarationOriginEntry || source.DeclarationKind != output.DeclarationKind ||
            source.DeclarationValue != output.DeclarationValue || source.RawLinkValue != output.RawLinkValue ||
            source.DeclarationExpandedName != output.DeclarationExpandedName || source.DeclarationLinkPresent != output.DeclarationLinkPresent ||
            source.ResourceEntry != output.ResourceEntry || source.FontEntry != output.FontEntry || source.FontExpandedName != output.FontExpandedName ||
            source.FontOrdinalWithinFonts != output.FontOrdinalWithinFonts || source.FontElementPreorderOrdinal != output.FontElementPreorderOrdinal ||
            source.FontFileRaw != output.FontFileRaw || source.ResBaseLocRaw != output.ResBaseLocRaw)
            throw new InvalidDataException("Actual font mapping requires unsupported origin/ID/path/literal relocation.");
    }
}
