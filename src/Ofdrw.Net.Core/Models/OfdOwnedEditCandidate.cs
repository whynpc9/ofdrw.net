using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// This certificate can only be issued by the complete ownership-transfer
// factory. IDs, XML strings, equal options or a caller-supplied map cannot issue it.
internal sealed partial class OfdOwnedEditCandidate : IDisposable
{
    private readonly OfdDocumentPackage source;
    private readonly OfdSerializedGuard sourceGuard;
    private readonly Dictionary<object, object> mapping;
    private int state; // editable -> claimed -> terminal (failure also consumes)
    internal OfdDocumentPackage Package { get; }
    internal OfdControlledEditContext Context { get; }
    private OfdOwnedEditCandidate(OfdDocumentPackage source, OfdDocumentPackage package,
        Dictionary<object, object> mapping, OfdSerializedGuard guard, OfdControlledEditContext context)
    { this.source = source; Package = package; this.mapping = mapping; sourceGuard = guard; Context = context; }

    internal static OfdOwnedEditCandidate CloneForControlledEdit(OfdDocumentPackage source,
        OfdSerializedGuard guard, OfdControlledEditContext context, CancellationToken token = default)
    {
        context.AssertActive();
        try { context.ClaimClone(); return CloneCore(source, guard, context, token); }
        catch { context.Fail(); throw; }
    }

    private static OfdOwnedEditCandidate CloneCore(OfdDocumentPackage source,
        OfdSerializedGuard guard, OfdControlledEditContext context, CancellationToken token)
    {
        context.AssertActive(); guard.AssertUnchanged(source, context, token);
        var material = context.Material;
        material.GuardPreservedXml(source, token); source.SourceIdentity?.Reconcile(source, material, token);
        OfdNavigationXml.Validate(source, source.Pages, token, material);
        material.ChargeCopy(4096L + (long)(source.Fonts.Count + source.Attachments.Count) * 256, token);
        var mapping = new Dictionary<object, object>(OfdReferenceIdentity.Instance);
        var copy = new OfdDocumentPackage
        {
            DocumentEntryPath = source.DocumentEntryPath, PublicResourceLocation = source.PublicResourceLocation,
            DocumentResourceLocation = source.DocumentResourceLocation,
            Options = new OfdDocumentOptions
            {
                DocType = source.Options.DocType, Namespace = source.Options.Namespace, DocumentId = source.Options.DocumentId,
                EnableDeflateCompression = source.Options.EnableDeflateCompression,
                DefaultPageWidthMillimeters = source.Options.DefaultPageWidthMillimeters,
                DefaultPageHeightMillimeters = source.Options.DefaultPageHeightMillimeters,
                Metadata = new OfdMetadata
                {
                    Title = source.Options.Metadata.Title, Author = source.Options.Metadata.Author,
                    Subject = source.Options.Metadata.Subject, Keywords = source.Options.Metadata.Keywords,
                    Creator = source.Options.Metadata.Creator, CreationDate = source.Options.Metadata.CreationDate,
                    ModificationDate = source.Options.Metadata.ModificationDate
                }
            }
        };
        Map(source, copy); Map(source.Options, copy.Options); Map(source.Options.Metadata, copy.Options.Metadata);
        Map(source.Pages, copy.Pages); Map(source.Fonts, copy.Fonts); Map(source.Attachments, copy.Attachments);
        Map(source.CustomTags, copy.CustomTags); Map(source.PreservedEntries, copy.PreservedEntries);
        Map(source.PreservedCommonDataElements, copy.PreservedCommonDataElements);
        Map(source.PreservedDocumentElements, copy.PreservedDocumentElements); Map(source.PreservedDocBodyElements, copy.PreservedDocBodyElements);
        foreach (var page in source.Pages) { token.ThrowIfCancellationRequested(); copy.Pages.Add(Page(page)); }
        // Original-page identities are evidence even after removal. Their
        // detached clones never become new active pages or borrow source data.
        foreach (var page in source.SourceIdentity?.OriginalPages ?? Array.Empty<OfdPage>())
            if (!mapping.ContainsKey(page)) _ = Page(page);
        foreach (var font in source.Fonts)
        {
            token.ThrowIfCancellationRequested(); var clone = new OfdFontResource
            {
                Id = font.Id, FontName = font.FontName, FamilyName = font.FamilyName, Charset = font.Charset,
                Bold = font.Bold, Italic = font.Italic, FileName = font.FileName, Data = Bytes(font.Data)
            };
            Map(font, clone); copy.Fonts.Add(clone);
        }
        foreach (var attachment in source.Attachments)
        {
            token.ThrowIfCancellationRequested(); var clone = new OfdAttachment
            {
                Id = attachment.Id, Name = attachment.Name, MediaType = attachment.MediaType, Format = attachment.Format,
                CreationDate = attachment.CreationDate, ModificationDate = attachment.ModificationDate,
                SizeKilobytes = attachment.SizeKilobytes, Visible = attachment.Visible, Usage = attachment.Usage,
                IsExternal = attachment.IsExternal, ExternalPath = attachment.ExternalPath, Data = Bytes(attachment.Data)
            };
            Map(attachment, clone); copy.Attachments.Add(clone);
        }
        foreach (var entry in source.PreservedEntries) { material.ChargeCopy(128, token); copy.PreservedEntries.Add(entry.Key, Bytes(entry.Value)); }
        foreach (var tag in source.CustomTags) { material.ChargeCopy(128, token); copy.CustomTags.Add(tag.Key, tag.Value); }
        Strings(source.PreservedCommonDataElements, copy.PreservedCommonDataElements);
        Strings(source.PreservedDocumentElements, copy.PreservedDocumentElements); Strings(source.PreservedDocBodyElements, copy.PreservedDocBodyElements);
        OfdNavigationCloner.Clone(source, copy, mapping, material, token, requireClosedMap: true);
        copy.SourceEnvelope = source.SourceEnvelope?.Clone(mapping, material, token, requireClosedMap: true);
        copy.SourceIdentity = source.SourceIdentity?.CloneFor(copy, mapping, material, token, requireClosedMap: true);
        guard.AssertUnchanged(source, context, token);
        // A complete second inventory admits the resulting graph before the
        // private certificate is issued. No model or mutable array is published
        // on a partial factory failure.
        material.InspectStorage(copy, token);
        return new OfdOwnedEditCandidate(source, copy, mapping, guard, context);

        void Map(object original, object clone)
        { material.ChargeCopy(128, token); if (mapping.ContainsKey(original)) throw new InvalidDataException("Ambiguous clone ownership."); mapping.Add(original, clone); }
        byte[] Bytes(byte[] original)
        { material.ChargeCopy(original.LongLength, token); var clone = new byte[original.Length]; Array.Copy(original, clone, original.Length); return clone; }
        void Strings(List<string> original, List<string> clone)
        { material.ChargeCopy((long)original.Count * IntPtr.Size * 2, token); foreach (var value in original) { token.ThrowIfCancellationRequested(); clone.Add(value); } }
        OfdPage Page(OfdPage page)
        {
            // Cover owner/list/map growth before the existing cloner allocates.
            material.ChargeCopy(512, token);
            return OfdModelCloner.ClonePage(page, keepSourcePath: true, clonePayloads: true, material: material, cancellationToken: token, mapping: mapping);
        }
    }

    internal T GetMapped<T>(T original) where T : class
    {
        AssertEditable();
        if (!mapping.TryGetValue(original, out var mapped) || mapped is not T typed)
            throw new InvalidOperationException("Object does not belong to this owned clone map.");
        return typed;
    }

    private void AssertEditable()
    { Context.AssertActive(); if (Volatile.Read(ref state) != 0) throw new InvalidOperationException("Candidate is no longer editable."); }

    internal OfdSerializedGuard ClaimWrite(CancellationToken token)
    {
        Context.AssertActive();
        if (Interlocked.CompareExchange(ref state, 1, 0) != 0) throw new InvalidOperationException("Owned candidate has already been consumed.");
        sourceGuard.AssertUnchanged(source, Context, token);
        return Context.CaptureSource(Package, token);
    }

    internal void AssertSourceUnchanged(CancellationToken token) => sourceGuard.AssertUnchanged(source, Context, token);
    internal void AssertClaimedActive()
    { Context.AssertActive(); if (Volatile.Read(ref state) != 1) throw new InvalidOperationException("Owned candidate authority has been disposed."); }
    internal OfdSerializedGuard CaptureCandidate(CancellationToken token = default)
    { AssertEditable(); return Context.CaptureSource(Package, token); }
    public void Dispose() { Interlocked.Exchange(ref state, 2); mapping.Clear(); }
}
