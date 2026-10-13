using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Threading;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

// RawOfdText is already OFD encoded by the caller. This layer only performs
// XML escaping; node count does not constrain the number of Unicode scalars.
internal readonly struct OfdTextSlotEdit(int ordinal, string rawOfdText, double x, double y, string? deltaX = null, string? deltaY = null)
{
    internal int Ordinal { get; } = ordinal;
    internal string RawOfdText { get; } = rawOfdText;
    internal double X { get; } = x;
    internal double Y { get; } = y;
    internal string? DeltaX { get; } = deltaX;
    internal string? DeltaY { get; } = deltaY;
}

internal sealed partial class OfdOwnedEditCandidate
{
    private readonly HashSet<OfdTextElement> patchedText = new();

    internal void ApplyTextSlots(OfdTextElement mappedText, OfdTextSlotEdit[] edits, CancellationToken token = default)
    {
        AssertEditable();
        if (Interlocked.CompareExchange(ref state, 3, 0) != 0) throw new InvalidOperationException("Candidate is busy.");
        try { ApplyTextSlotsCore(mappedText, edits, token); Volatile.Write(ref state, 0); }
        catch { Volatile.Write(ref state, 2); Context.Fail(); throw; }
    }

    private void ApplyTextSlotsCore(OfdTextElement mappedText, OfdTextSlotEdit[] edits, CancellationToken token)
    {
        token.ThrowIfCancellationRequested();
        if (mappedText is null || edits is null) throw new ArgumentNullException(mappedText is null ? nameof(mappedText) : nameof(edits));
        if (patchedText.Contains(mappedText)) throw new InvalidOperationException("Retained text slots have already been patched.");
        var material = Context.Material;
        // Read only the private owned candidate here. Whole-source guards run
        // at callback/clone/write boundaries, not once per glyph or TextObject.
        OfdTextElement? original = null;
        foreach (var pair in mapping)
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (ReferenceEquals(pair.Value, mappedText)) { original = pair.Key as OfdTextElement; break; }
        }
        if (original is null || Package.SourceIdentity is null || !Package.SourceIdentity.TryOwner(mappedText, out var owner) ||
            owner.Kind != XNamespace.Get(Package.Options.Namespace) + "TextObject" || !owner.ActiveDefinition)
            throw new NotSupportedException("Text slot editing requires an authenticated imported TextObject root.");
        if (string.IsNullOrWhiteSpace(mappedText.SourceXml))
            throw new InvalidDataException("Retained TextObject source is stale.");
        material.GuardFragmentXml(mappedText.SourceXml!, token);
        material.ChargeCopy(4L * mappedText.SourceXml!.Length, token);
        var root = OfdXmlEnvelope.Parse(mappedText.SourceXml!); var ns = root.Name.Namespace;
        Package.SourceIdentity.AssertFragment(mappedText, root, true, material, token);
        if (!Package.PreservedEntries.TryGetValue(owner.Entry, out var bytes)) throw new InvalidDataException("Text source entry is missing.");
        var document = material.ReadXml(owner.Entry, bytes, token); XElement? sourceRoot = null; var ordinal = 0;
        foreach (var node in document.Root!.DescendantsAndSelf())
        { token.ThrowIfCancellationRequested(); material.Structure.Reference(); if (ordinal++ == owner.Slot) { sourceRoot = node; break; } }
        var capturedSource = sourceRoot is null ? null : OfdXmlEnvelope.Capture(sourceRoot, material, token);
        material.ChargeWork(2L * ((capturedSource?.Length ?? 0) + (long)mappedText.SourceXml.Length), token);
        if (capturedSource != mappedText.SourceXml)
            throw new InvalidDataException("Text root no longer agrees with its authenticated source slot.");
        if (root.Nodes().Any(node => node is XText text && !string.IsNullOrWhiteSpace(text.Value)))
            throw new NotSupportedException("Mixed TextObject text cannot be patched.");
        var slots = new List<XElement>();
        foreach (var child in root.Elements())
        {
            token.ThrowIfCancellationRequested(); material.Structure.Reference();
            if (child.Name.LocalName != "TextCode") continue;
            if (child.Name != ns + "TextCode") throw new NotSupportedException("Foreign TextCode lookalike cannot be patched.");
            if (child.Attributes().Any(attribute => !attribute.IsNamespaceDeclaration &&
                (attribute.Name.Namespace != XNamespace.None || attribute.Name.LocalName is not ("X" or "Y" or "DeltaX" or "DeltaY"))) ||
                child.Nodes().Any(node => node is not XText || node is XCData))
                throw new NotSupportedException("TextCode slot has unsupported attributes or mixed content.");
            material.ChargeCopy(32, token); slots.Add(child);
        }
        if (slots.Count == 0 || edits.Length != slots.Count || mappedText.Runs.Count != slots.Count)
            throw new InvalidDataException("TextCode slots, edits and typed runs must have identical counts.");
        material.ChargeCopy(8L * mappedText.SourceXml!.Length + (long)slots.Count * 128, token);
        material.ChargeWork(2L * mappedText.SourceXml.Length, token);
        ValidateProjection(mappedText, root, slots, owner.Id, Package, token);

        // Freeze request values before building anything. No caller collection
        // is re-enumerated during mutation, and every slot must occur once in order.
        material.ChargeCopy((long)edits.Length * 128, token);
        var frozen = new OfdTextSlotEdit[edits.Length]; long totalText = 0;
        for (var index = 0; index < frozen.Length; index++)
        {
            token.ThrowIfCancellationRequested(); var edit = edits[index];
            if (edit.Ordinal != index || edit.RawOfdText is null || !Finite(edit.X) || !Finite(edit.Y))
                throw new InvalidDataException("Text slot edits require ordered ordinals, raw text and finite explicit anchors.");
            material.StoreString(edit.RawOfdText); material.StoreString(edit.DeltaX); material.StoreString(edit.DeltaY);
            totalText = checked(totalText + edit.RawOfdText.Length);
            material.ChargeCopy(2L * edit.RawOfdText.Length + 2L * (edit.DeltaX?.Length ?? 0) + 2L * (edit.DeltaY?.Length ?? 0) + 2048, token);
            frozen[index] = edit;
        }
        if (totalText > int.MaxValue) throw OfdTraversalBudget.Limit("Text projection is too large.");
        material.ChargeCopy(totalText * 2, token);
        var texts = new string[frozen.Length]; var runs = new OfdTextRun[frozen.Length];
        for (var index = 0; index < frozen.Length; index++)
        {
            token.ThrowIfCancellationRequested(); var edit = frozen[index]; var slot = slots[index];
            slot.Value = edit.RawOfdText;
            slot.SetAttributeValue("X", edit.X.ToString("R", CultureInfo.InvariantCulture));
            slot.SetAttributeValue("Y", edit.Y.ToString("R", CultureInfo.InvariantCulture));
            slot.SetAttributeValue("DeltaX", edit.DeltaX); slot.SetAttributeValue("DeltaY", edit.DeltaY);
            texts[index] = edit.RawOfdText;
            runs[index] = new OfdTextRun { Text = edit.RawOfdText, XMillimeters = edit.X, YMillimeters = edit.Y, DeltaX = edit.DeltaX, DeltaY = edit.DeltaY };
        }
        Package.SourceIdentity.AssertFragment(mappedText, root, true, material, token);
        var xml = OfdXmlEnvelope.Capture(root, material, token); var textValue = string.Concat(texts);
        token.ThrowIfCancellationRequested();
        // Commit only after the complete root and every request have passed.
        // Existing run owners are retained, preserving the factory's owner map.
        for (var index = 0; index < runs.Length; index++)
        {
            var target = mappedText.Runs[index]; var run = runs[index];
            target.Text = run.Text; target.XMillimeters = run.XMillimeters; target.YMillimeters = run.YMillimeters; target.DeltaX = run.DeltaX; target.DeltaY = run.DeltaY;
        }
        mappedText.SourceXml = xml; mappedText.Text = textValue; patchedText.Add(mappedText);
    }

    private static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
    private static double Number(string? value, double fallback)
    {
        if (value is null) return fallback;
        if (!double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out var number) || !Finite(number))
            throw new InvalidDataException("Text projection requires finite invariant numbers.");
        return number;
    }

    private static void ValidateProjection(OfdTextElement text, XElement root, List<XElement> slots, string id, OfdDocumentPackage package, CancellationToken token)
    {
        if (text.ObjectId != id || root.Attribute("ID")?.Value != id || root.Attribute("Font")?.Value != text.FontResourceId ||
            Number(root.Attribute("Size")?.Value, 4) != text.FontSizeMillimeters ||
            Number(root.Attribute("Weight")?.Value, 400) != text.Weight ||
            (root.Attribute("Italic")?.Value is "true" or "1") != text.Italic)
            throw new InvalidDataException("Text root identity/style projection is inconsistent.");
        var font = package.Fonts.SingleOrDefault(value => value.Id == text.FontResourceId);
        if (font is null || text.FontName != font.FontName) throw new InvalidDataException("Text font binding projection is inconsistent.");
        if (!OfdBoxParser.TryParse(root.Attribute("Boundary")?.Value, out var box) ||
            box.x != text.XMillimeters || box.y != text.YMillimeters || box.w != text.WidthMillimeters || box.h != text.HeightMillimeters)
            throw new InvalidDataException("Text boundary projection is inconsistent.");
        var matrix = root.Attribute("CTM")?.Value;
        var numbers = matrix?.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (numbers is null ? text.Transform is not null : text.Transform is null || numbers.Length != 6 || text.Transform.Length != 6)
            throw new InvalidDataException("Text transform projection is inconsistent.");
        if (numbers is not null) for (var i = 0; i < 6; i++) if (Number(numbers[i], 0) != text.Transform![i]) throw new InvalidDataException("Text transform projection is inconsistent.");
        var fill = root.Elements().FirstOrDefault(node => node.Name.LocalName == "FillColor");
        var components = fill?.Attribute("Value")?.Value.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (fill is null ? text.FillColor.Red != 0 || text.FillColor.Green != 0 || text.FillColor.Blue != 0 || text.FillColor.Alpha != 255 :
            components is null || components.Length != 3 || Number(components[0], 0) != text.FillColor.Red || Number(components[1], 0) != text.FillColor.Green ||
            Number(components[2], 0) != text.FillColor.Blue || Number(fill.Attribute("Alpha")?.Value, 255) != text.FillColor.Alpha)
            throw new InvalidDataException("Text color projection is inconsistent.");
        var clips = root.Element(root.Name.Namespace + "Clips")?.ToString(SaveOptions.DisableFormatting);
        if (clips != text.ClippingXml) throw new InvalidDataException("Text clipping projection is inconsistent.");
        var values = new string[slots.Count];
        for (var index = 0; index < slots.Count; index++)
        {
            token.ThrowIfCancellationRequested();
            var slot = slots[index]; var run = text.Runs[index]; values[index] = slot.Value;
            if (run.Text != slot.Value || run.XMillimeters != Number(slot.Attribute("X")?.Value, 0) ||
                run.YMillimeters != Number(slot.Attribute("Y")?.Value, text.FontSizeMillimeters) ||
                run.DeltaX != slot.Attribute("DeltaX")?.Value || run.DeltaY != slot.Attribute("DeltaY")?.Value)
                throw new InvalidDataException("Typed text runs disagree with retained TextCode slots.");
        }
        if (string.Concat(values) != text.Text) throw new InvalidDataException("Typed text disagrees with retained TextCode slots.");
    }
}
