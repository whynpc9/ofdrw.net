using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.CompilerServices;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Immutable scalar/reference tokens, not a hash of mutable IDs or page indexes.
// This detects edits through List<T>, child properties and shared references.
internal sealed class OfdNavigationSnapshot
{
    private readonly List<Token> tokens;
    private readonly List<BookmarkTarget> bookmarkTargets;
    private readonly bool opaque;
    private OfdNavigationSnapshot(List<Token> tokens, List<BookmarkTarget> bookmarkTargets, bool opaque)
    { this.tokens = tokens; this.bookmarkTargets = bookmarkTargets; this.opaque = opaque; }
    internal readonly struct BookmarkTarget(OfdBookmark bookmark, string name)
    {
        internal readonly OfdBookmark Bookmark = bookmark;
        internal readonly string Name = name;
    }
    internal IEnumerable<BookmarkTarget> BookmarkTargets => bookmarkTargets;
    internal IEnumerable<OfdPage> ReferencedPages
    {
        get { foreach (var token in tokens) if (token.Reference && token.Value is OfdPage page) yield return page; }
    }
    internal IEnumerable<OfdBookmark> ReferencedBookmarks
    {
        get { foreach (var token in tokens) if (token.Reference && token.Value is OfdBookmark bookmark) yield return bookmark; }
    }
    internal bool HasOpaqueSource => opaque;
    internal IEnumerable<object> ReferencedNodes
    { get { foreach (var item in tokens) if (item.Reference && item.Value is IOfdNavigationNode) yield return item.Value; } }

    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("snapshotTokens", tokens, true); sink.Item("count", tokens.Count);
        foreach (var item in tokens) sink.Item("snapshotValue", item.Value, item.Reference);
        sink.Item("bookmarkTargets", bookmarkTargets, true); sink.Item("count", bookmarkTargets.Count);
        foreach (var target in bookmarkTargets) { sink.Item("bookmark", target.Bookmark, true); sink.Item("name", target.Name); }
        sink.Item("opaque", opaque);
    }

    internal void InspectMaterial(OfdModelMaterialization material, bool output, CancellationToken token, bool guardXml = false)
    {
        foreach (var item in tokens)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (guardXml && item.Reference && item.Value is IOfdNavigationNode node)
            {
                if (node.Source.Xml is string source) material.GuardFragmentXml(source, token);
                if (node is OfdRawAction raw) material.GuardFragmentXml(raw.Xml, token);
            }
            if (item.Value is string text)
            {
                material.StoreString(text);
                if (output) { material.ChargeOutput(6L * text.Length); material.ChargeCopy(2L * text.Length, token); }
            }
            else if (output && item.Value is not null && !item.Reference) material.ChargeOutput(326);
        }
    }

    internal OfdNavigationSnapshot Remap(Func<object, object> map, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(128L + (long)(tokens.Count + bookmarkTargets.Count) * 32, token);
        var copied = new List<Token>(tokens.Count); var targets = new List<BookmarkTarget>(bookmarkTargets.Count);
        foreach (var item in tokens)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            copied.Add(new Token(item.Reference && item.Value is not null ? map(item.Value) : item.Value, item.Reference));
        }
        foreach (var target in bookmarkTargets)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            targets.Add(new BookmarkTarget((OfdBookmark)map(target.Bookmark), target.Name));
        }
        return new OfdNavigationSnapshot(copied, targets, opaque);
    }

    internal bool Matches(object value)
    {
        var current = Capture(value);
        if (tokens.Count != current.tokens.Count) return false;
        for (var i = 0; i < tokens.Count; i++)
        {
            var a = tokens[i]; var b = current.tokens[i];
            if (a.Reference != b.Reference || (a.Reference ? !ReferenceEquals(a.Value, b.Value) : !Equals(a.Value, b.Value))) return false;
        }
        return true;
    }

    internal static OfdNavigationSnapshot Capture(object value, OfdModelMaterialization? material = null, CancellationToken cancellationToken = default)
    {
        var limits = new OfdModelLimits(); var budget = material?.Structure ?? new OfdTraversalBudget(limits);
        material?.ReserveReadModel(1024, cancellationToken);
        var result = new List<Token>(); var targets = new List<BookmarkTarget>(); var opaque = false;
        var active = new HashSet<object>(Identity.Instance);
        var stack = new Stack<Frame>(); stack.Push(new Frame(value, 1, false));
        while (stack.Count > 0)
        {
            cancellationToken.ThrowIfCancellationRequested(); var frame = stack.Pop();
            if (frame.Exit) { active.Remove(frame.Value); continue; }
            budget.Node(frame.Depth); budget.Reference();
            if (!active.Add(frame.Value)) throw new InvalidDataException("Navigation contains an ownership cycle.");
            stack.Push(new Frame(frame.Value, frame.Depth, true));
            Add(frame.Value, reference: true);
            if (frame.Value is IOfdNavigationNode source) { Add(source.Source.Xml); opaque |= source.Source.HasOpaqueLocalContent; }
            switch (frame.Value)
            {
                case OfdOutlineCollection outlines: Push(outlines.Items, frame.Depth); break;
                case OfdBookmarkCollection bookmarks: Push(bookmarks.Items, frame.Depth); break;
                case OfdActionCollection actions: Push(actions.Items, frame.Depth); break;
                case OfdOutline outline:
                    Add(outline.Title); Add(outline.Expanded); Add(outline.Count);
                    Push(outline.Children, frame.Depth);
                    var owned = outline.ActionOwnership;
                    Add(owned.Value is not null || owned.OriginalXml is not null);
                    if (owned.Value is not null) PushOne(owned.Value, frame.Depth);
                    else if (owned.OriginalXml is not null && owned.BaselineSnapshot is not null)
                        {
                            foreach (var token in owned.BaselineSnapshot.tokens) Add(token.Value, token.Reference);
                            targets.AddRange(owned.BaselineSnapshot.bookmarkTargets); opaque |= owned.BaselineSnapshot.opaque;
                        }
                    break;
                case OfdBookmark bookmark:
                    Add(bookmark.Name); PushOne(bookmark.Destination, frame.Depth); break;
                case OfdDestination destination:
                    Add(destination.Page, reference: true); Add(destination.Type);
                    Add(destination.Left); Add(destination.Top); Add(destination.Right); Add(destination.Bottom); Add(destination.Zoom);
                    break;
                case OfdGoToAction go:
                    Add(go.Event); Add(go.Bookmark, reference: true); Add(go.Bookmark?.Name);
                    if (go.Bookmark is not null) targets.Add(new BookmarkTarget(go.Bookmark, go.Bookmark.Name));
                    Add(go.Destination is not null);
                    if (go.Destination is not null) PushOne(go.Destination, frame.Depth);
                    break;
                case OfdRawAction raw: Add(raw.Event); Add(raw.Xml); break;
                default: throw new NotSupportedException("Unsupported typed navigation node.");
            }
        }
        return new OfdNavigationSnapshot(result, targets, opaque);

        void Add(object? item, bool reference = false)
        {
            budget.Reference();
            if (item is string text) { if (material is null) budget.XmlBytes(2L * text.Length); else material.StoreString(text); }
            material?.ReserveReadModel(128, cancellationToken);
            result.Add(new Token(item, reference));
        }
        void PushOne(object? item, int depth)
        {
            if (item is null) throw new InvalidDataException("Navigation contains a null owned child.");
            budget.Reference(); budget.Depth(depth + 1);
            material?.ReserveReadModel(128, cancellationToken);
            stack.Push(new Frame(item, depth + 1, false));
        }
        void Push<T>(List<T> items, int depth) where T : class
        {
            Add(items.Count);
            for (var i = items.Count - 1; i >= 0; i--) PushOne(items[i], depth);
        }
    }

    private readonly struct Token(object? value, bool reference)
    {
        internal readonly object? Value = value;
        internal readonly bool Reference = reference;
    }
    private readonly struct Frame(object value, int depth, bool exit)
    {
        internal readonly object Value = value;
        internal readonly int Depth = depth;
        internal readonly bool Exit = exit;
    }
    private sealed class Identity : IEqualityComparer<object>
    {
        internal static readonly Identity Instance = new();
        public new bool Equals(object? a, object? b) => ReferenceEquals(a, b);
        public int GetHashCode(object value) => RuntimeHelpers.GetHashCode(value);
    }
}
