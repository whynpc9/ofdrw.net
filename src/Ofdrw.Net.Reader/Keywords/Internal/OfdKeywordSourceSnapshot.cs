using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Reader.Keywords.Internal;

// Holds original authority and mutation-verification evidence, never a second graphics model.
internal sealed class OfdKeywordSourceSnapshot : IDisposable
{
    private readonly OfdDocumentPackage owner;
    private readonly List<Action<OfdKeywordLocatorBudget>> checks = new();
    private readonly Dictionary<string, byte[]> controls = new(StringComparer.Ordinal);
    private readonly List<(OfdFontResource Source, OfdGlyphFontSnapshot Font, string? Entry)> bindings = new();
    private readonly List<(OfdFontResource Source, string Id, bool Bold, bool Italic, string Digest, string? Entry)> retainedBindings = new();
    private readonly OfdKeywordLocatorLimits limits;
    private readonly HashSet<string> entryKeys;
    private readonly Dictionary<OfdTextRun, int> runOwners = new();
    private readonly SHA256 fingerprint = SHA256.Create();
    private readonly byte[] hashBuffer = new byte[8192];
    private int hashUsed;
    private bool sealedSnapshot;
    private long controlCopyBytes;

    internal OfdKeywordSourceSnapshot(OfdDocumentPackage source, OfdKeywordLocatorBudget budget)
    {
        owner = source; limits = budget.Limits;
        entryKeys = new HashSet<string>(StringComparer.Ordinal);
        try
        {
        foreach (var key in source.PreservedEntries.Keys) { budget.Reference(); entryKeys.Add(key); }
        checks.Add(b =>
        {
            b.Field(); Require(source.PreservedEntries.Count == entryKeys.Count);
            foreach (var key in source.PreservedEntries.Keys) { b.Field(); Require(entryKeys.Contains(key)); }
        });
        String(() => source.DocumentEntryPath, budget);
        var options = source.Options;
        checks.Add(b => { b.Field(); Require(ReferenceEquals(source.Options, options)); });
        String(() => options.Namespace, budget); String(() => options.DocumentId, budget);
        var pages = List(source.Pages, budget);
        foreach (var page in pages)
        {
            budget.Check(); String(() => page.SourceEntryPath, budget); String(() => page.Id, budget);
            Number(() => page.Index, budget); Number(() => page.XMillimeters, budget); Number(() => page.YMillimeters, budget);
            Number(() => page.WidthMillimeters, budget); Number(() => page.HeightMillimeters, budget);
            CaptureElements(List(page.Elements, budget), budget);
            CaptureElements(List(page.AnnotationAppearances, budget), budget);
            foreach (var template in List(page.Templates, budget))
            {
                String(() => template.TemplateId, budget); String(() => template.BaseLocation, budget); String(() => template.ZOrder, budget);
                CaptureElements(List(template.Elements, budget), budget);
            }
            Strings(page.PreservedPageElements, budget);
        }
        Fonts = List(source.Fonts, budget);
        foreach (var font in Fonts)
        {
            String(() => font.Id, budget); String(() => font.FontName, budget); String(() => font.FamilyName, budget);
            String(() => font.FileName, budget); String(() => font.Charset, budget);
            Number(() => font.Bold ? 1 : 0, budget); Number(() => font.Italic ? 1 : 0, budget);
        }
        Strings(source.PreservedCommonDataElements, budget); Strings(source.PreservedDocumentElements, budget);
        Strings(source.PreservedDocBodyElements, budget);
        }
        catch { fingerprint.Dispose(); throw; }
    }
    internal OfdFontResource[] Fonts { get; }
    internal string Version { get; private set; } = string.Empty;

    private void CaptureElements(OfdElement[] elements, OfdKeywordLocatorBudget budget)
    {
        foreach (var element in elements)
        {
            if (element is null) throw new InvalidDataException("Null source graphic.");
            String(() => element.ObjectId, budget); String(() => element.LayerId, budget); String(() => element.LayerType, budget);
            String(() => element.ClippingXml, budget);
            Number(() => element.XMillimeters, budget); Number(() => element.YMillimeters, budget);
            Number(() => element.WidthMillimeters, budget); Number(() => element.HeightMillimeters, budget);
            if (element is OfdTextElement text)
            {
                String(() => text.SourceXml, budget); String(() => text.Text, budget); String(() => text.FontName, budget);
                String(() => text.FontResourceId, budget); Number(() => text.FontSizeMillimeters, budget);
                Number(() => text.Weight, budget); Number(() => text.Italic ? 1 : 0, budget);
                Matrix(() => text.Transform, budget);
                var color = text.FillColor;
                checks.Add(b => { b.Field(); Require(ReferenceEquals(color, text.FillColor)); });
                foreach (var run in List(text.Runs, budget))
                {
                    runOwners.TryGetValue(run, out var uses); runOwners[run] = uses + 1;
                    String(() => run.Text, budget); String(() => run.DeltaX, budget); String(() => run.DeltaY, budget);
                    Number(() => run.XMillimeters, budget); Number(() => run.YMillimeters, budget);
                }
            }
            else if (element is OfdImageElement image) { String(() => image.SourceXml, budget); Matrix(() => image.Transform, budget); }
            else if (element is OfdPathElement path) { String(() => path.SourceXml, budget); Matrix(() => path.Transform, budget); }
            else if (element is OfdRawElement raw) { String(() => raw.Xml, budget); String(() => raw.LocalName, budget); }
            // A0 diagnoses unknown definitions as Raw; their base ownership and
            // fields are retained here without attempting another model parser.
        }
    }
    private T[] List<T>(List<T> list, OfdKeywordLocatorBudget budget)
    {
        if (list.Count > limits.MaxReferences) throw new InvalidDataException("Source list capacity exceeded.");
        budget.ReserveReferences(list.Count);
        var copy = new T[list.Count];
        Add(list.Count.ToString(CultureInfo.InvariantCulture), budget);
        for (var i = 0; i < copy.Length; i++) { budget.Check(); copy[i] = list[i]; }
        checks.Add(b =>
        {
            b.Field(); Require(list.Count == copy.Length);
            for (var i = 0; i < copy.Length; i++) { b.Field(); Require(ReferenceEquals(list[i], copy[i])); }
        });
        return copy;
    }
    private void Strings(List<string> list, OfdKeywordLocatorBudget budget)
    {
        var copy = List(list, budget);
        foreach (var value in copy) Add(value, budget);
    }
    private void String(Func<string?> read, OfdKeywordLocatorBudget budget)
    {
        var copy = read(); Add(copy, budget);
        checks.Add(b =>
        {
            b.Field(); var current = read();
            if (ReferenceEquals(current, copy)) return;
            b.Characters(Math.Max(copy?.Length ?? 0, current?.Length ?? 0));
            Require(string.Equals(current, copy, StringComparison.Ordinal));
        });
    }
    private void Number(Func<double> read, OfdKeywordLocatorBudget budget)
    {
        var copy = read(); Add(copy.ToString("R", CultureInfo.InvariantCulture), budget);
        checks.Add(b => { b.Field(); Require(read().Equals(copy)); });
    }
    private void Matrix(Func<double[]?> read, OfdKeywordLocatorBudget budget)
    {
        var value = read();
        if (value is { Length: > 6 }) throw new InvalidDataException("Invalid source transform length.");
        var copy = value is null ? null : (double[])value.Clone();
        Add(copy is null ? null : string.Join(" ", copy.Select(n => n.ToString("R", CultureInfo.InvariantCulture))), budget);
        checks.Add(b =>
        {
            b.Field(); var current = read();
            Require((current is null) == (copy is null)); if (current is null) return;
            Require(current.Length == copy!.Length);
            for (var i = 0; i < current.Length; i++) { b.Field(); Require(current[i].Equals(copy[i])); }
        });
    }

    internal byte[] ControlEntry(string path, OfdKeywordLocatorBudget budget)
    {
        if (controls.TryGetValue(path, out var cached)) return cached;
        var original = ExactEntry(path);
        if (original.Length > limits.MaxSingleXmlBytes || original.Length > limits.MaxXmlBytes - controlCopyBytes)
            throw new InvalidDataException("Source control entry copy capacity exceeded.");
        controlCopyBytes += original.Length;
        var copy = new byte[original.Length];
        for (var start = 0; start < copy.Length; start += 8192)
        {
            budget.Check(); Buffer.BlockCopy(original, start, copy, start, Math.Min(8192, copy.Length - start));
        }
        controls.Add(path, copy); Add(path, budget);
        Add(copy.Length.ToString(CultureInfo.InvariantCulture), budget);
        for (var i = 0; i < copy.Length; i++) { if ((i & 255) == 0) budget.Check(); Put(copy[i]); }
        checks.Add(b => CompareBytes(ExactEntry(path), copy, b));
        return copy;
    }
    internal byte[] ExactEntry(string path)
    {
        // Package's dictionary is case-insensitive; OFD resource paths are not.
        if (!entryKeys.Contains(path) ||
            !owner.PreservedEntries.TryGetValue(path, out var bytes) || bytes is null)
            throw new NotSupportedException("MissingExactSourceEntry");
        return bytes;
    }
    internal bool HasExactEntry(string path) => entryKeys.Contains(path);
    internal bool HasSharedRun(OfdTextElement source, OfdKeywordLocatorBudget budget)
    {
        foreach (var run in source.Runs) { budget.Check(); if (runOwners[run] > 1) return true; }
        return false;
    }
    internal void Bind(OfdFontResource source, OfdGlyphFontSnapshot font, string? entry, OfdKeywordLocatorBudget budget)
    {
        if (bindings.Any(binding => ReferenceEquals(binding.Source, source))) return;
        bindings.Add((source, font, entry));
        retainedBindings.Add((source, font.ResourceId, font.Bold, font.Italic, font.ContentSha256, entry));
        Add(font.ResourceId, budget); Add(font.ContentSha256, budget); Add(entry, budget);
    }
    internal void ValidateFontEntries(OfdKeywordLocatorBudget budget)
    {
        foreach (var binding in bindings)
            if (binding.Entry is not null) CompareBytes(ExactEntry(binding.Entry), binding.Font.Data, budget);
    }
    internal void Seal(OfdKeywordLocatorBudget budget)
    {
        Flush(); fingerprint.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        Version = "q2:" + Guid.NewGuid().ToString("N") + ":" + BitConverter.ToString(fingerprint.Hash!).Replace("-", "").ToLowerInvariant();
        fingerprint.Dispose(); sealedSnapshot = true; budget.Check();
    }
    internal void ValidateUnchanged(OfdDocumentPackage package, OfdKeywordLocatorBudget budget)
    {
        budget.Check(); Require(ReferenceEquals(package, owner));
        try { foreach (var check in checks) { budget.Check(); check(budget); } }
        catch (NotSupportedException exception) { throw new InvalidDataException("Keyword source entry changed after capture.", exception); }
        budget.Check();
    }
    // Future editing must bind the same package instance and revalidate font content,
    // never accept SourceVersion or ObjectId supplied by a caller as proof.
    internal void ValidateUnchanged(OfdDocumentPackage package)
    {
        using var operation = new OfdKeywordOperation(); var budget = new OfdKeywordLocatorBudget(limits, operation);
        ValidateUnchanged(package, budget);
        foreach (var binding in retainedBindings)
        {
            Require(binding.Source.Id == binding.Id && binding.Source.Bold == binding.Bold && binding.Source.Italic == binding.Italic);
            Require(Digest(binding.Source.Data, budget) == binding.Digest);
            if (binding.Entry is not null) Require(Digest(ExactEntry(binding.Entry), budget) == binding.Digest);
        }
    }
    private static string Digest(byte[] data, OfdKeywordLocatorBudget budget)
    {
        budget.Bytes(data.Length); using var sha = SHA256.Create();
        for (var start = 0; start < data.Length; start += 8192)
        { budget.Check(); var count = Math.Min(8192, data.Length - start); sha.TransformBlock(data, start, count, null, 0); }
        sha.TransformFinalBlock(Array.Empty<byte>(), 0, 0);
        return BitConverter.ToString(sha.Hash!).Replace("-", "").ToLowerInvariant();
    }
    private static void CompareBytes(byte[] current, byte[] copy, OfdKeywordLocatorBudget budget)
    {
        budget.Bytes(copy.Length); Require(current.Length == copy.Length);
        for (var i = 0; i < copy.Length; i++) { if ((i & 255) == 0) budget.Check(); Require(current[i] == copy[i]); }
    }
    private void Add(string? value, OfdKeywordLocatorBudget budget)
    {
        if (sealedSnapshot) throw new InvalidOperationException("Snapshot is sealed.");
        if (value is { Length: > 0 } && value.Length > Math.Max(limits.MaxSingleXmlBytes, 16_000_000))
            throw new InvalidDataException("Source fingerprint text capacity exceeded.");
        budget.Field(); budget.Characters(value?.Length ?? 0);
        var length = value?.Length ?? -1;
        for (var shift = 0; shift < 32; shift += 8) Put((byte)(length >> shift));
        if (value is null) return;
        for (var i = 0; i < value.Length; i++)
        { if ((i & 255) == 0) budget.Check(); Put((byte)value[i]); Put((byte)(value[i] >> 8)); }
    }
    private void Put(byte value) { hashBuffer[hashUsed++] = value; if (hashUsed == hashBuffer.Length) Flush(); }
    private void Flush() { if (hashUsed > 0) fingerprint.TransformBlock(hashBuffer, 0, hashUsed, hashBuffer, 0); hashUsed = 0; }
    public void Dispose() => fingerprint.Dispose();
    private static void Require(bool condition) { if (!condition) throw new InvalidDataException("Keyword source changed or belongs to another snapshot."); }
}
