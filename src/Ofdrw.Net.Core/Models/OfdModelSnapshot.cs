using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Ofdrw.Net.Core.Models;

// Immutable values captured through the same field table as storage/guard/output.
// Payloads are digested; mutable arrays are never retained as baseline values.
internal sealed class OfdModelSnapshot
{
    private readonly List<Entry> entries;
    private readonly struct Entry(string role, object? value, bool identity)
    { internal readonly string Role = role; internal readonly object? Value = value; internal readonly bool Identity = identity; }
    private OfdModelSnapshot(List<Entry> entries) => this.entries = entries;
    internal static OfdModelSnapshot Capture(object scope, OfdModelMaterialization? material = null, CancellationToken token = default)
    {
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        material.ChargeCopy(256, token); var entries = new List<Entry>();
        OfdSerializedInventory.Walk(null!, new Sink(material, token, entries, false), material, token, scope: scope);
        return new OfdModelSnapshot(entries);
    }
    internal bool Matches(object scope, OfdModelMaterialization? material = null, CancellationToken token = default)
    {
        material ??= new OfdModelMaterialization(new OfdModelLimits());
        var sink = new Sink(material, token, entries, true);
        OfdSerializedInventory.Walk(null!, sink, material, token, scope: scope); return sink.Equal && sink.Count == entries.Count;
    }
    internal OfdModelSnapshot Remap(IDictionary<object, object> map, OfdModelMaterialization material, CancellationToken token)
    {
        material.ChargeCopy(128L + 32L * entries.Count, token); var result = new List<Entry>(entries.Count);
        foreach (var entry in entries)
        {
            token.ThrowIfCancellationRequested();
            if (entry.Identity && entry.Value is not null)
            {
                if (!map.TryGetValue(entry.Value, out var replacement)) throw new InvalidDataException("Annotation baseline clone map is not closed.");
                result.Add(new Entry(entry.Role, replacement, true));
            }
            else result.Add(entry);
        }
        return new OfdModelSnapshot(result);
    }
    internal void VisitInventory(OfdSerializedInventory.ISink sink)
    {
        sink.Item("snapshot", entries, true); sink.Item("count", entries.Count);
        foreach (var entry in entries) { sink.Item("role", entry.Role); sink.Item("value", entry.Value, entry.Identity); }
    }
    private sealed class Sink(OfdModelMaterialization material, CancellationToken token, List<Entry> entries, bool comparing) : OfdSerializedInventory.ISink
    {
        internal int Count; internal bool Equal = true;
        private readonly Dictionary<byte[], byte[]> hashes = new();
        public void Item(string role, object? value, bool identity = false)
        {
            token.ThrowIfCancellationRequested(); material.ChargeWork(value is string text ? text.Length * 2L : 16, token);
            if (value is double number) value = BitConverter.DoubleToInt64Bits(number);
            if (value is DateTime date) value = (date.Ticks, date.Kind);
            if (!comparing) { material.ChargeCopy(128, token); entries.Add(new Entry(role, value, identity)); }
            else if (Count >= entries.Count) Equal = false;
            else { var expected = entries[Count]; Equal &= role == expected.Role && identity == expected.Identity && (identity ? ReferenceEquals(value, expected.Value) : Equals(value, expected.Value)); }
            Count++;
        }
        public void Bytes(string role, byte[] value, object? owner = null)
        {
            if (value is null) throw new InvalidDataException("Null annotation payload.");
            Item(role, value.LongLength);
            if (!hashes.TryGetValue(value, out var hash)) { material.ChargeCopy(64, token); hashes.Add(value, hash = material.HashBytes(value, token)); }
            foreach (var part in hash) Item("digest", part);
        }
        public void Matrix(double[]? value)
        { Item("matrixLength", value?.Length); if (value is not null) foreach (var part in value) Item("number", part); }
        public void Xml(string? value) { if (value is not null) material.GuardFragmentXml(value, token); }
        public void Path(string? value) { }
    }
}
