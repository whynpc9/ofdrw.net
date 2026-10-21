using System;
using System.Collections.Generic;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

internal static class OfdNavigationCloner
{
    internal static void Clone(OfdDocumentPackage source, OfdDocumentPackage destination,
        IDictionary<object, object> mapping, OfdModelMaterialization material, CancellationToken token, bool requireClosedMap = false)
    {
        var sources = new Dictionary<OfdNavigationSource, OfdNavigationSource>();
        Ownership(source.OutlineOwnership, destination.OutlineOwnership, 1);
        Ownership(source.BookmarkOwnership, destination.BookmarkOwnership, 1);
        Ownership(source.ActionOwnership, destination.ActionOwnership, 1);

        object Map(object original, int depth)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference(); material.Structure.Depth(depth);
            if (mapping.TryGetValue(original, out var mapped)) return mapped;
            if (original is not IOfdNavigationNode node)
            {
                if (requireClosedMap) throw new InvalidOperationException("Navigation owner mapping is incomplete.");
                return original; // selected-page clones may retain excluded baseline identities
            }
            material.Structure.Node(depth);
            material.ChargeCopy(256, token);
            IOfdNavigationNode clone = original switch
            {
                OfdOutlineCollection => new OfdOutlineCollection(), OfdBookmarkCollection => new OfdBookmarkCollection(), OfdActionCollection => new OfdActionCollection(),
                OfdOutline outline => new OfdOutline { Title = outline.Title, Expanded = outline.Expanded, Count = outline.Count },
                OfdBookmark bookmark => new OfdBookmark { Name = bookmark.Name },
                OfdDestination destination => new OfdDestination { Type = destination.Type, Left = destination.Left, Top = destination.Top, Right = destination.Right, Bottom = destination.Bottom, Zoom = destination.Zoom },
                OfdGoToAction go => new OfdGoToAction { Event = go.Event }, OfdRawAction raw => new OfdRawAction { Xml = raw.Xml, Event = raw.Event },
                OfdUriAction uri => new OfdUriAction { Event = uri.Event, Uri = uri.Uri, Base = uri.Base, Target = uri.Target },
                OfdAttachmentAction attachment => new OfdAttachmentAction { Event = attachment.Event, AttachId = attachment.AttachId, NewWindow = attachment.NewWindow },
                OfdSoundAction sound => new OfdSoundAction { Event = sound.Event, ResourceId = sound.ResourceId, Volume = sound.Volume, Repeat = sound.Repeat, Synchronous = sound.Synchronous },
                OfdMovieAction movie => new OfdMovieAction { Event = movie.Event, ResourceId = movie.ResourceId, Operator = movie.Operator },
                _ => throw new NotSupportedException("Unknown navigation clone contract.")
            };
            mapping.Add(original, clone);
            if (original is OfdOutlineCollection oc) MapList(oc.Items, ((OfdOutlineCollection)clone).Items);
            if (original is OfdBookmarkCollection bc) MapList(bc.Items, ((OfdBookmarkCollection)clone).Items);
            if (original is OfdActionCollection ac) MapList(ac.Items, ((OfdActionCollection)clone).Items);
            if (original is OfdOutline oo) MapList(oo.Children, ((OfdOutline)clone).Children);
            // Establish the source memo before traversing source slots.
            CopySource(node.Source, clone.Source, depth);
            switch (original)
            {
                case OfdOutlineCollection outlines: foreach (var item in outlines.Items) ((OfdOutlineCollection)clone).Items.Add((OfdOutline)Map(item, depth + 1)); break;
                case OfdBookmarkCollection bookmarks: foreach (var item in bookmarks.Items) ((OfdBookmarkCollection)clone).Items.Add((OfdBookmark)Map(item, depth + 1)); break;
                case OfdActionCollection actions: foreach (var item in actions.Items) ((OfdActionCollection)clone).Items.Add((OfdAction)Map(item, depth + 1)); break;
                case OfdOutline outline:
                    foreach (var child in outline.Children) ((OfdOutline)clone).Children.Add((OfdOutline)Map(child, depth + 1));
                    Ownership(outline.ActionOwnership, ((OfdOutline)clone).ActionOwnership, depth + 1); break;
                case OfdBookmark bookmark: ((OfdBookmark)clone).Destination = (OfdDestination)Map(bookmark.Destination, depth + 1); break;
                case OfdDestination destination: ((OfdDestination)clone).Page = (OfdPage)Map(destination.Page, depth + 1); break;
                case OfdGoToAction go:
                    ((OfdGoToAction)clone).Destination = go.Destination is null ? null : (OfdDestination)Map(go.Destination, depth + 1);
                    ((OfdGoToAction)clone).Bookmark = go.Bookmark is null ? null : (OfdBookmark)Map(go.Bookmark, depth + 1); break;
            }
            return clone;
        }
        void MapList<T>(List<T> original, List<T> clone)
        { material.ChargeCopy(128L + (long)original.Count * IntPtr.Size * 2, token); mapping.Add(original, clone); }
        OfdNavigationSource? Source(OfdNavigationSource? original, int depth)
        {
            if (original is null) return null;
            if (sources.TryGetValue(original, out var clone)) return clone;
            material.ChargeCopy(128, token);
            clone = new OfdNavigationSource(); CopySource(original, clone, depth); return clone;
        }
        void CopySource(OfdNavigationSource original, OfdNavigationSource clone, int depth)
        {
            if (sources.TryGetValue(original, out var existing))
            {
                material.ChargeCopy((long)existing.Slots.Count * 128, token);
                clone.Xml = existing.Xml; clone.OriginalName = existing.OriginalName; clone.HasOpaqueLocalContent = existing.HasOpaqueLocalContent; clone.DocumentOwner = existing.DocumentOwner;
                foreach (var slot in existing.Slots) { material.Structure.Reference(); clone.Slots.Add(slot.Key, slot.Value); }
                return;
            }
            material.ChargeCopy(128L + (long)original.Slots.Count * 128, token); sources.Add(original, clone);
            clone.Xml = original.Xml; clone.OriginalName = original.OriginalName; clone.HasOpaqueLocalContent = original.HasOpaqueLocalContent;
            if (requireClosedMap && original.DocumentOwner is not null && !ReferenceEquals(original.DocumentOwner, source))
                throw new System.IO.InvalidDataException("Foreign navigation document owner.");
            clone.DocumentOwner = ReferenceEquals(original.DocumentOwner, source) ? destination : original.DocumentOwner;
            foreach (var slot in original.Slots)
            { token.ThrowIfCancellationRequested(); material.Structure.Reference(); clone.Slots.Add(slot.Key, Map(slot.Value, depth + 1)); }
        }
        void Ownership<T>(OfdNavigationOwnership<T> original, OfdNavigationOwnership<T> clone, int depth) where T : class
        {
            var value = original.Value is null ? null : (T)Map(original.Value, depth);
            var envelope = Source(original.OriginalSource, depth);
            var snapshot = original.BaselineSnapshot?.Remap(value => Map(value, depth), material, token);
            clone.RestoreClonedState(value, original.OriginalXml, envelope, snapshot);
        }
    }
}
