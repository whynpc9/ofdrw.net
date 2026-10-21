using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

/// <summary>Document usage declared in DocInfo. An omitted value has the standard's Normal default.</summary>
public enum OfdDocUsage { Normal, EBook, ENewsPaper, EMagzine }

// The XML backing is deliberately private: typed edits retain extension nodes,
// attributes, namespace bindings and the original order without exposing a
// second mutable XML ownership path to callers.
public abstract class OfdStandardNode
{
    private readonly XElement root;
    internal string? OriginalXml { get; private set; }
    internal object? DocumentOwner { get; private set; }
    internal XElement Root => root;
    protected OfdStandardNode(string local) { root = new XElement(local); }
    internal OfdStandardNode(XElement source) { root = source; }
    internal void Bind(object owner, string xml) { OriginalXml = xml; DocumentOwner = owner; }
    internal bool IsDirty => OriginalXml is not null && !XNode.DeepEquals(root, OfdXmlEnvelope.Parse(OriginalXml));
    internal XElement Emit(XNamespace ns, object owner)
    {
        if (DocumentOwner is not null && !ReferenceEquals(DocumentOwner, owner))
            throw new NotSupportedException("Imported standard object belongs to another document.");
        var copy = new XElement(root);
        if (copy.Name.Namespace == XNamespace.None)
        {
            foreach (var element in copy.DescendantsAndSelf().Where(e => e.Name.Namespace == XNamespace.None))
                element.Name = ns + element.Name.LocalName;
        }
        if (copy.Name.Namespace != ns) throw new NotSupportedException("Cannot migrate standard object namespace.");
        return copy;
    }
    internal string InventoryXml => root.ToString(SaveOptions.DisableFormatting);
    protected string? Text(string name) => root.Element(root.Name.Namespace + name)?.Value;
    protected void Text(string name, string? value)
    {
        var nodes = root.Elements(root.Name.Namespace + name).ToArray();
        if (nodes.Length > 1) throw new InvalidDataException("Duplicate standard object field.");
        if (nodes.Length == 0) { if (value is not null) OfdXmlEnvelope.Singleton(root, name, new XElement(root.Name.Namespace + name, value)); return; }
        var node = nodes[0];
        if (node.Nodes().Any(n => n is not XText))
            throw new NotSupportedException("Cannot overwrite mixed standard field content.");
        if (value is null)
        {
            if (node.HasAttributes || node.Nodes().Any(n => n is not XText))
                throw new NotSupportedException("Cannot remove a field carrying unknown content.");
            node.Remove();
        }
        else if (node.Value != value) { node.Nodes().OfType<XText>().Remove(); node.AddFirst(new XText(value)); }
    }
    protected string? Attribute(string name) => root.Attribute(name)?.Value;
    protected void Attribute(string name, string? value) => root.SetAttributeValue(name, value);
    protected bool? Boolean(string name) => TryBoolean(Text(name));
    protected void Boolean(string name, bool? value) => Text(name, value?.ToString().ToLowerInvariant());
    internal static bool? TryBoolean(string? text) => text switch { "true" or "1" => true, "false" or "0" => false, _ => null };
    internal static bool WellFormed(XElement root, params string[] singletonFields)
    {
        var ns = root.Name.Namespace;
        return singletonFields.All(name => root.Elements(ns + name).Take(2).Count() <= 1);
    }
    internal static bool ValidBoolean(XElement root, params string[] names)
    {
        var ns = root.Name.Namespace;
        return names.All(name => root.Element(ns + name) is not XElement e || TryBoolean(e.Value).HasValue);
    }
}

/// <summary>An ordered DocInfo custom metadata entry. Name is required by the OFD schema.</summary>
public sealed class OfdCustomData : OfdStandardNode
{
    public OfdCustomData(string name, string value) : base("CustomData") { Name = name; Value = value; }
    internal OfdCustomData(XElement source) : base(source) { }
    public string Name { get => Attribute("Name") ?? string.Empty; set => Attribute("Name", value); }
    public string Value
    {
        get => Root.Value;
        set
        {
            if (Root.Nodes().Any(n => n is not XText))
                throw new NotSupportedException("Cannot overwrite mixed custom metadata content.");
            Root.Nodes().OfType<XText>().Remove(); Root.AddFirst(new XText(value));
        }
    }
}

/// <summary>Ordered DocInfo custom metadata. Unknown children remain in their source positions.</summary>
public sealed class OfdCustomDatas : OfdStandardNode
{
    public OfdCustomDatas() : base("CustomDatas") { }
    internal OfdCustomDatas(XElement source) : base(source) { }
    public IReadOnlyList<OfdCustomData> Items => Root.Elements(Root.Name.Namespace + "CustomData").Select(e => new OfdCustomData(e)).ToArray();
    public OfdCustomData Add(string name, string value)
    {
        if (string.IsNullOrEmpty(name)) throw new ArgumentException("CustomData Name is required.", nameof(name));
        var item = new OfdCustomData(name, value);
        var attached = item.Emit(Root.Name.Namespace, this); Root.Add(attached); return new OfdCustomData(attached);
    }
    public void RemoveAt(int index) => Root.Elements(Root.Name.Namespace + "CustomData").ElementAt(index).Remove();
    public void Clear() { foreach (var item in Root.Elements(Root.Name.Namespace + "CustomData").ToArray()) item.Remove(); }
}

/// <summary>Document permission declarations. This model does not enforce access restrictions.</summary>
public sealed class OfdPermissions : OfdStandardNode
{
    public OfdPermissions() : base("Permissions") { }
    internal OfdPermissions(XElement source) : base(source) { }
    public bool? Edit { get => Boolean("Edit"); set => Boolean("Edit", value); }
    public bool? Annot { get => Boolean("Annot"); set => Boolean("Annot", value); }
    public bool? Export { get => Boolean("Export"); set => Boolean("Export", value); }
    public bool? Signature { get => Boolean("Signature"); set => Boolean("Signature", value); }
    public bool? Watermark { get => Boolean("Watermark"); set => Boolean("Watermark", value); }
    public bool? PrintScreen { get => Boolean("PrintScreen"); set => Boolean("PrintScreen", value); }
    public OfdPrintPermission? Print
    {
        get => Root.Element(Root.Name.Namespace + "Print") is XElement e ? new OfdPrintPermission(e) : null;
        set => OfdXmlEnvelope.Singleton(Root, "Print", value?.Emit(Root.Name.Namespace, this));
    }
    public OfdValidPeriod? ValidPeriod
    {
        get => Root.Element(Root.Name.Namespace + "ValidPeriod") is XElement e ? new OfdValidPeriod(e) : null;
        set => OfdXmlEnvelope.Singleton(Root, "ValidPeriod", value?.Emit(Root.Name.Namespace, this));
    }
    internal static bool Recognized(XElement root) => WellFormed(root, "Edit", "Annot", "Export", "Signature", "Watermark", "PrintScreen", "Print", "ValidPeriod") &&
        ValidBoolean(root, "Edit", "Annot", "Export", "Signature", "Watermark", "PrintScreen") &&
        (root.Element(root.Name.Namespace + "Print") is not XElement print ||
            (print.Attribute("Printable") is XAttribute p && TryBoolean(p.Value).HasValue &&
             (print.Attribute("Copies") is not XAttribute c || int.TryParse(c.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _)))) &&
        (root.Element(root.Name.Namespace + "ValidPeriod") is not XElement period ||
            ValidDate(period.Attribute("StartDate")?.Value) && ValidDate(period.Attribute("EndDate")?.Value));
    private static bool ValidDate(string? value)
    {
        if (value is null) return true;
        try { _ = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind); return true; }
        catch (FormatException) { return false; }
    }
}

public sealed class OfdPrintPermission : OfdStandardNode
{
    public OfdPrintPermission() : base("Print") { }
    internal OfdPrintPermission(XElement source) : base(source) { }
    public bool? Printable { get => TryBoolean(Attribute("Printable")); set => Attribute("Printable", value?.ToString().ToLowerInvariant()); }
    public int? Copies { get => int.TryParse(Attribute("Copies"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null; set => Attribute("Copies", value?.ToString(CultureInfo.InvariantCulture)); }
}

public sealed class OfdValidPeriod : OfdStandardNode
{
    public OfdValidPeriod() : base("ValidPeriod") { }
    internal OfdValidPeriod(XElement source) : base(source) { }
    /// <summary>XML dateTime lexical value; retained exactly to preserve timezone omission or offset.</summary>
    public string? StartDate { get => Attribute("StartDate"); set { CheckDate(value); Attribute("StartDate", value); } }
    /// <summary>XML dateTime lexical value; retained exactly to preserve timezone omission or offset.</summary>
    public string? EndDate { get => Attribute("EndDate"); set { CheckDate(value); Attribute("EndDate", value); } }
    private static void CheckDate(string? value) { if (value is not null) _ = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind); }
}

public enum OfdPageMode { None, FullScreen, UseOutlines, UseThumbs, UseCustomTags, UseLayers, UseAttatchs, UseBookmarks }
public enum OfdPageLayout { OnePage, OneColumn, TwoPageL, TwoColumnL, TwoPageR, TwoColumnR }
public enum OfdTabDisplay { DocTitle, FileName }
public enum OfdZoomMode { Default, FitHeight, FitWidth, FitRect }

/// <summary>Viewer preference declarations; this library does not apply them to a UI.</summary>
public sealed class OfdViewingPreferences : OfdStandardNode
{
    public OfdViewingPreferences() : base("VPreferences") { }
    internal OfdViewingPreferences(XElement source) : base(source) { }
    public OfdPageMode? PageMode { get => EnumValue<OfdPageMode>("PageMode"); set => Text("PageMode", value?.ToString()); }
    public OfdPageLayout? PageLayout { get => EnumValue<OfdPageLayout>("PageLayout"); set => Text("PageLayout", value?.ToString()); }
    public OfdTabDisplay? TabDisplay { get => EnumValue<OfdTabDisplay>("TabDisplay"); set => Text("TabDisplay", value?.ToString()); }
    public bool? HideToolbar { get => Boolean("HideToolbar"); set => Boolean("HideToolbar", value); }
    public bool? HideMenubar { get => Boolean("HideMenubar"); set => Boolean("HideMenubar", value); }
    public bool? HideWindowUI { get => Boolean("HideWindowUI"); set => Boolean("HideWindowUI", value); }
    public OfdZoomMode? ZoomMode { get => EnumValue<OfdZoomMode>("ZoomMode"); set { if (value.HasValue) Text("Zoom", null); Text("ZoomMode", value?.ToString()); } }
    public double? Zoom { get => double.TryParse(Text("Zoom"), NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : null; set { if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value))) throw new ArgumentOutOfRangeException(nameof(value)); if (value.HasValue) Text("ZoomMode", null); Text("Zoom", value?.ToString("R", CultureInfo.InvariantCulture)); } }
    private T? EnumValue<T>(string name) where T : struct, Enum => Enum.TryParse<T>(Text(name), false, out var value) && Enum.IsDefined(typeof(T), value) ? value : null;
    internal static bool Recognized(XElement root)
    {
        if (!WellFormed(root, "PageMode", "PageLayout", "TabDisplay", "HideToolbar", "HideMenubar", "HideWindowUI", "ZoomMode", "Zoom")) return false;
        if (!ValidBoolean(root, "HideToolbar", "HideMenubar", "HideWindowUI")) return false;
        var model = new OfdViewingPreferences(root);
        return (model.Text("PageMode") is null || model.PageMode.HasValue) &&
               (model.Text("PageLayout") is null || model.PageLayout.HasValue) &&
               (model.Text("TabDisplay") is null || model.TabDisplay.HasValue) &&
               (model.Text("ZoomMode") is null || model.ZoomMode.HasValue) &&
               (model.Text("Zoom") is null || model.Zoom.HasValue && !double.IsNaN(model.Zoom.Value) && !double.IsInfinity(model.Zoom.Value)) &&
               !(model.ZoomMode.HasValue && model.Zoom.HasValue);
    }
}

/// <summary>A version descriptor in the selected DocBody.</summary>
public sealed class OfdVersion : OfdStandardNode
{
    public OfdVersion(string id, int index, string baseLocation) : base("Version") { Id = id; Index = index; BaseLocation = baseLocation; }
    internal OfdVersion(XElement source) : base(source) { }
    public string Id { get => Attribute("ID") ?? string.Empty; set => Attribute("ID", value); }
    public int Index { get => int.TryParse(Attribute("Index"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0; set => Attribute("Index", value.ToString(CultureInfo.InvariantCulture)); }
    public bool? Current { get => TryBoolean(Attribute("Current")); set => Attribute("Current", value?.ToString().ToLowerInvariant()); }
    public string BaseLocation { get => Attribute("BaseLoc") ?? string.Empty; set => Attribute("BaseLoc", value); }
}

/// <summary>Ordered version descriptors. Referenced version files are preserved, never synthesized.</summary>
public sealed class OfdVersions : OfdStandardNode
{
    public OfdVersions() : base("Versions") { }
    internal OfdVersions(XElement source) : base(source) { }
    public IReadOnlyList<OfdVersion> Items => Root.Elements(Root.Name.Namespace + "Version").Select(e => new OfdVersion(e)).ToArray();
    public OfdVersion Add(string id, int index, string baseLocation)
    {
        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(baseLocation)) throw new ArgumentException("Version ID and BaseLoc are required.");
        var item = new OfdVersion(id, index, baseLocation);
        var attached = item.Emit(Root.Name.Namespace, this); Root.Add(attached); return new OfdVersion(attached);
    }
    public void RemoveAt(int index) => Root.Elements(Root.Name.Namespace + "Version").ElementAt(index).Remove();
    public void Clear() { foreach (var item in Root.Elements(Root.Name.Namespace + "Version").ToArray()) item.Remove(); }
    internal static bool Recognized(XElement root) => root.Elements(root.Name.Namespace + "Version").Any() &&
        root.Elements(root.Name.Namespace + "Version").All(e =>
        e.Attribute("ID") is XAttribute id && !string.IsNullOrWhiteSpace(id.Value) &&
        e.Attribute("Index") is XAttribute index && int.TryParse(index.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) &&
        e.Attribute("BaseLoc") is XAttribute path && !string.IsNullOrWhiteSpace(path.Value) &&
        (e.Attribute("Current") is not XAttribute current || TryBoolean(current.Value).HasValue));
}

internal static class OfdStandardXml
{
    internal static void ReadMetadata(OfdDocumentPackage package, XElement info, OfdModelMaterialization material,
        System.Threading.CancellationToken token)
    {
        var ns = info.Name.Namespace; var metadata = package.Options.Metadata;
        if (info.Elements(ns + "DocUsage").Count() == 1)
        {
            var value = info.Element(ns + "DocUsage")!.Value;
            if (Enum.TryParse<OfdDocUsage>(value, false, out var usage) && Enum.IsDefined(typeof(OfdDocUsage), usage)) metadata.DocUsage = usage;
        }
        metadata.Cover = Scalar(info, "Cover"); metadata.CreatorVersion = Scalar(info, "CreatorVersion");
        if (info.Elements(ns + "CustomDatas").Count() == 1)
        {
            var element = info.Element(ns + "CustomDatas")!;
            if (element.Elements(ns + "CustomData").All(e => e.Attribute("Name") is not null && !e.HasElements))
            {
                var xml = OfdXmlEnvelope.Capture(element, material, token);
                var model = new OfdCustomDatas(OfdXmlEnvelope.Parse(xml)); model.Bind(package, xml); metadata.CustomDatas = model;
            }
        }
    }

    internal static void ReadDocument(OfdDocumentPackage package, XElement? body, XElement document,
        OfdModelMaterialization material, System.Threading.CancellationToken token)
    {
        var ns = document.Name.Namespace;
        if (document.Elements(ns + "Permissions").Count() == 1 &&
            document.Element(ns + "Permissions") is XElement permissions && OfdPermissions.Recognized(permissions))
            package.Permissions = Bind(new OfdPermissions(OfdXmlEnvelope.Parse(OfdXmlEnvelope.Capture(permissions, material, token))), permissions);
        if (document.Elements(ns + "VPreferences").Count() == 1 &&
            document.Element(ns + "VPreferences") is XElement preferences && OfdViewingPreferences.Recognized(preferences))
            package.ViewingPreferences = Bind(new OfdViewingPreferences(OfdXmlEnvelope.Parse(OfdXmlEnvelope.Capture(preferences, material, token))), preferences);
        if (body?.Elements(body.Name.Namespace + "Versions").Count() == 1 &&
            body.Element(body.Name.Namespace + "Versions") is XElement versions && OfdVersions.Recognized(versions))
            package.Versions = Bind(new OfdVersions(OfdXmlEnvelope.Parse(OfdXmlEnvelope.Capture(versions, material, token))), versions);

        T Bind<T>(T model, XElement source) where T : OfdStandardNode
        { model.Bind(package, model.InventoryXml); return model; }
    }

    internal static void Validate(OfdDocumentPackage package, OfdModelMaterialization material,
        System.Threading.CancellationToken token)
    {
        var metadata = package.Options.Metadata;
        if (metadata.DocUsage.HasValue && !Enum.IsDefined(typeof(OfdDocUsage), metadata.DocUsage.Value)) throw new ArgumentException("Invalid DocUsage.");
        if (metadata.Cover is not null &&
            (package.SourceEnvelope is null || !package.SourceEnvelope.Metadata.TryGetValue("Cover", out var originalCover) || originalCover != metadata.Cover))
            ExistingPath("Cover", metadata.Cover);
        if (metadata.CustomDatas is { } data)
        {
            CheckOwner(data);
            foreach (var item in data.Items)
                if (string.IsNullOrEmpty(item.Name)) throw new ArgumentException("CustomData Name is required.");
        }
        if (package.Permissions is { } permissions)
        {
            CheckOwner(permissions);
            if (!OfdPermissions.Recognized(permissions.Root)) throw new ArgumentException("Invalid permissions declaration.");
            if (permissions.Print is { Printable: null }) throw new ArgumentException("Print permission requires Printable.");
        }
        if (package.ViewingPreferences is { } preferences)
        {
            CheckOwner(preferences);
            if (!OfdViewingPreferences.Recognized(preferences.Root)) throw new ArgumentException("Invalid viewer preferences.");
        }
        if (package.Versions is { } versions)
        {
            CheckOwner(versions);
            if (!OfdVersions.Recognized(versions.Root)) throw new ArgumentException("Invalid version declaration.");
            if (versions.Items.Count == 0 && versions.Root.Elements().Count() == 0) return;
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var version in versions.Items)
            {
                if (!ids.Add(version.Id)) throw new ArgumentException("Duplicate version ID.");
                try { XmlConvert.VerifyNCName(version.Id); }
                catch (XmlException error) { throw new ArgumentException("Version ID must be an XML ID.", error); }
                if (versions.OriginalXml is null || versions.IsDirty) ExistingPath("Version BaseLoc", version.BaseLocation);
            }
        }
        void CheckOwner(OfdStandardNode value)
        {
            if (value.DocumentOwner is not null && !ReferenceEquals(value.DocumentOwner, package))
                throw new NotSupportedException("Imported standard object belongs to another document.");
            material.GuardFragmentXml(value.InventoryXml, token);
        }
        void ExistingPath(string role, string location)
        {
            if (string.IsNullOrWhiteSpace(location)) throw new ArgumentException(role + " path is empty.");
            var resolved = material.ResolvePath("OFD.xml", location, token);
            if (!package.PreservedEntries.ContainsKey(resolved))
                throw new NotSupportedException(role + " reference is not an existing package entry.");
        }
    }

    internal static void WriteMetadata(OfdDocumentPackage package, XElement info)
    {
        var metadata = package.Options.Metadata;
        Scalar("DocUsage", metadata.DocUsage?.ToString());
        Scalar("Cover", metadata.Cover);
        Scalar("CreatorVersion", metadata.CreatorVersion);
        if (metadata.CustomDatas is { } data)
        {
            if (info.Element(info.Name.Namespace + "CustomDatas") is not null && data.OriginalXml is null)
                throw new NotSupportedException("Cannot replace an unclaimed CustomDatas source.");
            var emitted = data.Emit(info.Name.Namespace, package);
            OfdXmlEnvelope.Singleton(info, "CustomDatas", Empty(emitted) ? null : emitted);
        }
        void Scalar(string name, string? value)
        {
            if (value is null && !(package.SourceEnvelope?.Metadata.TryGetValue(name, out var original) == true && original is not null)) return;
            var nodes = info.Elements(info.Name.Namespace + name).ToArray();
            if (nodes.Length > 1) throw new InvalidDataException("Duplicate DocInfo field.");
            if (nodes.Length == 1)
            {
                var node = nodes[0];
                if (node.Value == value) return;
                if (node.Nodes().Any(n => n is not XText))
                    throw new NotSupportedException("Cannot overwrite mixed DocInfo content.");
                if (value is null)
                {
                    if (node.HasAttributes || node.Nodes().Any(n => n is not XText))
                        throw new NotSupportedException("Cannot remove DocInfo field with unknown content.");
                    node.Remove();
                }
                else { node.Nodes().OfType<XText>().Remove(); node.AddFirst(new XText(value)); }
            }
            else if (value is not null) OfdXmlEnvelope.Singleton(info, name, new XElement(info.Name.Namespace + name, value));
        }
    }

    internal static void WriteDocument(OfdDocumentPackage package, XElement document)
    {
        Write(package.Permissions, "Permissions"); Write(package.ViewingPreferences, "VPreferences");
        void Write(OfdStandardNode? value, string name)
        {
            if (value is null) return;
            if (document.Element(document.Name.Namespace + name) is not null && value.OriginalXml is null)
                throw new NotSupportedException("Cannot replace an unclaimed standard object source.");
            var emitted = value.Emit(document.Name.Namespace, package);
            OfdXmlEnvelope.Singleton(document, name, Empty(emitted) ? null : emitted);
        }
    }

    internal static void WriteVersions(OfdDocumentPackage package, XElement body)
    {
        if (package.Versions is not { } versions) return;
        if (body.Element(body.Name.Namespace + "Versions") is not null && versions.OriginalXml is null)
            throw new NotSupportedException("Cannot replace an unclaimed Versions source.");
        var emitted = versions.Emit(body.Name.Namespace, package);
        OfdXmlEnvelope.Singleton(body, "Versions", Empty(emitted) ? null : emitted);
    }

    internal static bool Empty(XElement element) =>
        !element.Attributes().Any(a => !a.IsNamespaceDeclaration) &&
        !element.Nodes().Any(n => n is not XText t || !string.IsNullOrWhiteSpace(t.Value));

    private static string? Scalar(XElement parent, string name)
    {
        var nodes = parent.Elements(parent.Name.Namespace + name).ToArray();
        return nodes.Length == 1 && !nodes[0].HasElements ? nodes[0].Value : null;
    }

    internal static T? Clone<T>(T? model, object target) where T : OfdStandardNode
    {
        if (model is null) return null;
        var root = new XElement(model.Root);
        OfdStandardNode clone = model switch
        {
            OfdCustomDatas => new OfdCustomDatas(root), OfdPermissions => new OfdPermissions(root),
            OfdViewingPreferences => new OfdViewingPreferences(root), OfdVersions => new OfdVersions(root),
            _ => throw new NotSupportedException("Unsupported standard object clone.")
        };
        if (model.OriginalXml is string source) clone.Bind(target, source);
        return (T)clone;
    }
}
