using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace Ofdrw.Net.Core.Models;

/// <summary>Document usage declared in DocInfo. An omitted value has the standard's Normal default.</summary>
public enum OfdDocUsage
{
    /// <summary>Ordinary document.</summary>
    Normal,
    /// <summary>Electronic book.</summary>
    EBook,
    /// <summary>Electronic newspaper.</summary>
    ENewsPaper,
    /// <summary>Electronic magazine; retains the standard's spelling.</summary>
    EMagzine
}

/// <summary>Base for XML-backed standard objects that retain extension content and document ownership.</summary>
public abstract class OfdStandardNode
{
    private readonly XElement root;
    internal string? OriginalXml { get; private set; }
    internal object? DocumentOwner { get; private set; }
    internal XElement Root => root;
    /// <summary>Creates a new standard object with the given local element name.</summary>
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
    /// <summary>Reads an optional child element's text.</summary>
    protected string? Text(string name) => root.Element(root.Name.Namespace + name)?.Value;
    /// <summary>Sets or removes a child element while refusing mixed source content.</summary>
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
    /// <summary>Reads an optional unqualified attribute.</summary>
    protected string? Attribute(string name) => root.Attribute(name)?.Value;
    /// <summary>Sets or removes an unqualified attribute.</summary>
    protected void Attribute(string name, string? value) => root.SetAttributeValue(name, value);
    /// <summary>Reads an optional Boolean child element.</summary>
    protected bool? Boolean(string name) => TryBoolean(Text(name));
    /// <summary>Sets or removes a Boolean child element.</summary>
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
    internal static void GuardRemoval(XElement element, bool textIsModeled, params string[] modeledAttributes)
    {
        if (element.Attributes().Any(attribute => attribute.IsNamespaceDeclaration || attribute.Name.Namespace != XNamespace.None ||
                !modeledAttributes.Contains(attribute.Name.LocalName, StringComparer.Ordinal)) ||
            element.Nodes().Any(node => node is not XText text || !textIsModeled && !string.IsNullOrWhiteSpace(text.Value)))
            throw new NotSupportedException("Cannot remove standard object entry with unknown source content.");
    }
}

/// <summary>An ordered DocInfo custom metadata entry. Name is required by the OFD schema.</summary>
public sealed class OfdCustomData : OfdStandardNode
{
    /// <summary>Creates a named custom metadata value.</summary>
    public OfdCustomData(string name, string value) : base("CustomData") { Name = name; Value = value; }
    internal OfdCustomData(XElement source) : base(source) { }
    /// <summary>Required name of this custom metadata entry.</summary>
    public string Name { get => Attribute("Name") ?? string.Empty; set => Attribute("Name", value); }
    /// <summary>Text value of this custom metadata entry.</summary>
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
    /// <summary>Creates an empty custom metadata collection.</summary>
    public OfdCustomDatas() : base("CustomDatas") { }
    internal OfdCustomDatas(XElement source) : base(source) { }
    /// <summary>Ordered view of the current custom metadata entries.</summary>
    public IReadOnlyList<OfdCustomData> Items => Root.Elements(Root.Name.Namespace + "CustomData").Select(e => new OfdCustomData(e)).ToArray();
    /// <summary>Adds a named text value and returns its attached entry.</summary>
    public OfdCustomData Add(string name, string value)
    {
        if (string.IsNullOrEmpty(name)) throw new ArgumentException("CustomData Name is required.", nameof(name));
        var item = new OfdCustomData(name, value);
        var attached = item.Emit(Root.Name.Namespace, this); Root.Add(attached); return new OfdCustomData(attached);
    }
    /// <summary>Removes the modeled entry at its zero-based index.</summary>
    public void RemoveAt(int index)
    {
        var item = Root.Elements(Root.Name.Namespace + "CustomData").ElementAt(index);
        GuardRemoval(item, true, "Name"); item.Remove();
    }
    /// <summary>Removes modeled entries while retaining extension children.</summary>
    public void Clear()
    {
        var items = Root.Elements(Root.Name.Namespace + "CustomData").ToArray();
        foreach (var item in items) GuardRemoval(item, true, "Name");
        foreach (var item in items) item.Remove();
    }
}

/// <summary>Document permission declarations. This model does not enforce access restrictions.</summary>
public sealed class OfdPermissions : OfdStandardNode
{
    /// <summary>Creates an empty set of permission declarations.</summary>
    public OfdPermissions() : base("Permissions") { }
    internal OfdPermissions(XElement source) : base(source) { }
    /// <summary>Optional editing permission declaration.</summary>
    public bool? Edit { get => Boolean("Edit"); set => Boolean("Edit", value); }
    /// <summary>Optional annotation permission declaration.</summary>
    public bool? Annot { get => Boolean("Annot"); set => Boolean("Annot", value); }
    /// <summary>Optional export permission declaration.</summary>
    public bool? Export { get => Boolean("Export"); set => Boolean("Export", value); }
    /// <summary>Optional signature permission declaration.</summary>
    public bool? Signature { get => Boolean("Signature"); set => Boolean("Signature", value); }
    /// <summary>Optional watermark permission declaration.</summary>
    public bool? Watermark { get => Boolean("Watermark"); set => Boolean("Watermark", value); }
    /// <summary>Optional screen-print permission declaration.</summary>
    public bool? PrintScreen { get => Boolean("PrintScreen"); set => Boolean("PrintScreen", value); }
    /// <summary>Optional print permission details.</summary>
    public OfdPrintPermission? Print
    {
        get => Root.Element(Root.Name.Namespace + "Print") is XElement e ? BindChild(new OfdPrintPermission(e)) : null;
        set
        {
            GuardReplacement("Print", value, "Printable", "Copies");
            OfdXmlEnvelope.Singleton(Root, "Print", value?.Emit(Root.Name.Namespace, DocumentOwner ?? this));
        }
    }
    /// <summary>Optional period during which the declarations are valid.</summary>
    public OfdValidPeriod? ValidPeriod
    {
        get => Root.Element(Root.Name.Namespace + "ValidPeriod") is XElement e ? BindChild(new OfdValidPeriod(e)) : null;
        set
        {
            GuardReplacement("ValidPeriod", value, "StartDate", "EndDate");
            OfdXmlEnvelope.Singleton(Root, "ValidPeriod", value?.Emit(Root.Name.Namespace, DocumentOwner ?? this));
        }
    }
    private void GuardReplacement(string name, OfdStandardNode? replacement, params string[] modeledAttributes)
    {
        var original = Root.Element(Root.Name.Namespace + name);
        if (original is null || ReferenceEquals(replacement?.Root, original)) return;
        if (original.Attributes().Any(attribute => attribute.IsNamespaceDeclaration ||
                attribute.Name.Namespace != XNamespace.None || !modeledAttributes.Contains(attribute.Name.LocalName, StringComparer.Ordinal)) ||
            original.Nodes().Any(node => node is not XText text || !string.IsNullOrWhiteSpace(text.Value)))
            throw new NotSupportedException("Cannot replace " + name + " with unknown source content.");
    }
    private T BindChild<T>(T child) where T : OfdStandardNode
    { if (DocumentOwner is not null) child.Bind(DocumentOwner, child.InventoryXml); return child; }
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

/// <summary>Print permission declaration; the library does not enforce it.</summary>
public sealed class OfdPrintPermission : OfdStandardNode
{
    /// <summary>Creates empty print permission details.</summary>
    public OfdPrintPermission() : base("Print") { }
    internal OfdPrintPermission(XElement source) : base(source) { }
    /// <summary>Required printability declaration when a Print element is written.</summary>
    public bool? Printable { get => TryBoolean(Attribute("Printable")); set => Attribute("Printable", value?.ToString().ToLowerInvariant()); }
    /// <summary>Optional maximum copy count; the schema uses minus one as its default.</summary>
    public int? Copies { get => int.TryParse(Attribute("Copies"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null; set => Attribute("Copies", value?.ToString(CultureInfo.InvariantCulture)); }
}

/// <summary>Optional validity interval of permission declarations.</summary>
public sealed class OfdValidPeriod : OfdStandardNode
{
    /// <summary>Creates an empty validity interval.</summary>
    public OfdValidPeriod() : base("ValidPeriod") { }
    internal OfdValidPeriod(XElement source) : base(source) { }
    /// <summary>XML dateTime lexical value; retained exactly to preserve timezone omission or offset.</summary>
    public string? StartDate { get => Attribute("StartDate"); set { CheckDate(value); Attribute("StartDate", value); } }
    /// <summary>XML dateTime lexical value; retained exactly to preserve timezone omission or offset.</summary>
    public string? EndDate { get => Attribute("EndDate"); set { CheckDate(value); Attribute("EndDate", value); } }
    private static void CheckDate(string? value) { if (value is not null) _ = XmlConvert.ToDateTime(value, XmlDateTimeSerializationMode.RoundtripKind); }
}

/// <summary>Initial page panel or viewing mode requested from a viewer.</summary>
public enum OfdPageMode
{
    /// <summary>No special panel.</summary>
    None,
    /// <summary>Full-screen viewing.</summary>
    FullScreen,
    /// <summary>Show the outline panel.</summary>
    UseOutlines,
    /// <summary>Show page thumbnails.</summary>
    UseThumbs,
    /// <summary>Show custom tags.</summary>
    UseCustomTags,
    /// <summary>Show layers.</summary>
    UseLayers,
    /// <summary>Show attachments; retains the OFD spelling.</summary>
    UseAttatchs,
    /// <summary>Show bookmarks.</summary>
    UseBookmarks
}

/// <summary>Page arrangement requested from a viewer.</summary>
public enum OfdPageLayout
{
    /// <summary>Display one page.</summary>
    OnePage,
    /// <summary>Display a single page column.</summary>
    OneColumn,
    /// <summary>Display a two-page spread with the first page on the left.</summary>
    TwoPageL,
    /// <summary>Display two columns with the first page on the left.</summary>
    TwoColumnL,
    /// <summary>Display a two-page spread with the first page on the right.</summary>
    TwoPageR,
    /// <summary>Display two columns with the first page on the right.</summary>
    TwoColumnR
}

/// <summary>Label to show for the document tab.</summary>
public enum OfdTabDisplay
{
    /// <summary>Use the document title.</summary>
    DocTitle,
    /// <summary>Use the file name.</summary>
    FileName
}

/// <summary>Initial zoom choice requested from a viewer.</summary>
public enum OfdZoomMode
{
    /// <summary>Use the viewer's default zoom.</summary>
    Default,
    /// <summary>Fit page height.</summary>
    FitHeight,
    /// <summary>Fit page width.</summary>
    FitWidth,
    /// <summary>Fit a rectangle.</summary>
    FitRect
}

/// <summary>Viewer preference declarations; this library does not apply them to a UI.</summary>
public sealed class OfdViewingPreferences : OfdStandardNode
{
    /// <summary>Creates empty viewer preferences.</summary>
    public OfdViewingPreferences() : base("VPreferences") { }
    internal OfdViewingPreferences(XElement source) : base(source) { }
    /// <summary>Optional initial page mode.</summary>
    public OfdPageMode? PageMode { get => EnumValue<OfdPageMode>("PageMode"); set => Text("PageMode", value?.ToString()); }
    /// <summary>Optional page layout.</summary>
    public OfdPageLayout? PageLayout { get => EnumValue<OfdPageLayout>("PageLayout"); set => Text("PageLayout", value?.ToString()); }
    /// <summary>Optional document-tab label choice.</summary>
    public OfdTabDisplay? TabDisplay { get => EnumValue<OfdTabDisplay>("TabDisplay"); set => Text("TabDisplay", value?.ToString()); }
    /// <summary>Optional instruction to hide the toolbar.</summary>
    public bool? HideToolbar { get => Boolean("HideToolbar"); set => Boolean("HideToolbar", value); }
    /// <summary>Optional instruction to hide the menu bar.</summary>
    public bool? HideMenubar { get => Boolean("HideMenubar"); set => Boolean("HideMenubar", value); }
    /// <summary>Optional instruction to hide other window controls.</summary>
    public bool? HideWindowUI { get => Boolean("HideWindowUI"); set => Boolean("HideWindowUI", value); }
    /// <summary>Named zoom choice, mutually exclusive with a numeric Zoom.</summary>
    public OfdZoomMode? ZoomMode { get => EnumValue<OfdZoomMode>("ZoomMode"); set { if (value.HasValue) Text("Zoom", null); Text("ZoomMode", value?.ToString()); } }
    /// <summary>Numeric initial zoom, mutually exclusive with ZoomMode.</summary>
    public double? Zoom { get => double.TryParse(Text("Zoom"), NumberStyles.Float, CultureInfo.InvariantCulture, out var n) ? n : null; set { if (value.HasValue && (double.IsNaN(value.Value) || double.IsInfinity(value.Value))) throw new ArgumentOutOfRangeException(nameof(value)); if (value.HasValue) Text("ZoomMode", null); Text("Zoom", value?.ToString("R", CultureInfo.InvariantCulture)); } }
    private T? EnumValue<T>(string name) where T : struct, Enum
    {
        var token = Text(name);
        return token is not null && Enum.GetNames(typeof(T)).Contains(token, StringComparer.Ordinal) &&
            Enum.TryParse<T>(token, false, out var value) ? value : null;
    }
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
    /// <summary>Creates a version descriptor with an ID, index, and package location.</summary>
    public OfdVersion(string id, int index, string baseLocation) : base("Version") { Id = id; Index = index; BaseLocation = baseLocation; }
    internal OfdVersion(XElement source) : base(source) { }
    /// <summary>XML ID of this version descriptor.</summary>
    public string Id { get => Attribute("ID") ?? string.Empty; set => Attribute("ID", value); }
    /// <summary>Version index declared in OFD.xml.</summary>
    public int Index { get => int.TryParse(Attribute("Index"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0; set => Attribute("Index", value.ToString(CultureInfo.InvariantCulture)); }
    /// <summary>Optional marker for the current version.</summary>
    public bool? Current { get => TryBoolean(Attribute("Current")); set => Attribute("Current", value?.ToString().ToLowerInvariant()); }
    /// <summary>Package-relative location of this version's existing file.</summary>
    public string BaseLocation { get => Attribute("BaseLoc") ?? string.Empty; set => Attribute("BaseLoc", value); }
}

/// <summary>Ordered version descriptors. Referenced version files are preserved, never synthesized.</summary>
public sealed class OfdVersions : OfdStandardNode
{
    private readonly Dictionary<XElement, string?> originalLocations = new();
    /// <summary>Creates an empty version descriptor collection.</summary>
    public OfdVersions() : base("Versions") { }
    internal OfdVersions(XElement source) : base(source)
    {
        foreach (var version in Root.Elements(Root.Name.Namespace + "Version"))
            originalLocations.Add(version, version.Attribute("BaseLoc")?.Value);
    }
    /// <summary>Ordered view of the current version descriptors.</summary>
    public IReadOnlyList<OfdVersion> Items => Root.Elements(Root.Name.Namespace + "Version").Select(e => new OfdVersion(e)).ToArray();
    /// <summary>Adds a descriptor for an existing version file.</summary>
    public OfdVersion Add(string id, int index, string baseLocation)
    {
        if (string.IsNullOrWhiteSpace(id) || string.IsNullOrWhiteSpace(baseLocation)) throw new ArgumentException("Version ID and BaseLoc are required.");
        var item = new OfdVersion(id, index, baseLocation);
        var attached = item.Emit(Root.Name.Namespace, this); Root.Add(attached); return new OfdVersion(attached);
    }
    /// <summary>Removes the modeled descriptor at its zero-based index.</summary>
    public void RemoveAt(int index)
    {
        var item = Root.Elements(Root.Name.Namespace + "Version").ElementAt(index);
        GuardRemoval(item, false, "ID", "Index", "Current", "BaseLoc"); item.Remove();
    }
    /// <summary>Removes modeled descriptors while retaining extension children.</summary>
    public void Clear()
    {
        var items = Root.Elements(Root.Name.Namespace + "Version").ToArray();
        foreach (var item in items) GuardRemoval(item, false, "ID", "Index", "Current", "BaseLoc");
        foreach (var item in items) item.Remove();
    }
    internal bool LocationChanged(OfdVersion version) =>
        !originalLocations.TryGetValue(version.Root, out var original) || original != version.BaseLocation;
    internal void CopyLocationBaselines(OfdVersions source)
    {
        originalLocations.Clear();
        var originals = source.Root.Elements(source.Root.Name.Namespace + "Version").ToArray();
        var copies = Root.Elements(Root.Name.Namespace + "Version").ToArray();
        for (var index = 0; index < originals.Length; index++)
            if (source.originalLocations.TryGetValue(originals[index], out var original))
                originalLocations.Add(copies[index], original);
    }
    internal static bool Recognized(XElement root)
    {
        var versions = root.Elements(root.Name.Namespace + "Version").ToArray();
        return versions.Length > 0 &&
            versions.Select(e => e.Attribute("ID")?.Value).Distinct(StringComparer.Ordinal).Count() == versions.Length &&
            versions.All(e =>
                e.Attribute("ID") is XAttribute id && ValidId(id.Value) &&
                e.Attribute("Index") is XAttribute index && int.TryParse(index.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out _) &&
                e.Attribute("BaseLoc") is XAttribute path && !string.IsNullOrWhiteSpace(path.Value) &&
                (e.Attribute("Current") is not XAttribute current || TryBoolean(current.Value).HasValue));
    }
    private static bool ValidId(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        try { XmlConvert.VerifyNCName(value); return true; }
        catch (XmlException) { return false; }
    }
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
            if (Enum.GetNames(typeof(OfdDocUsage)).Contains(value, StringComparer.Ordinal) &&
                Enum.TryParse<OfdDocUsage>(value, false, out var usage)) metadata.DocUsage = usage;
        }
        metadata.Cover = Scalar(info, "Cover"); metadata.CreatorVersion = Scalar(info, "CreatorVersion");
        if (info.Elements(ns + "CustomDatas").Count() == 1)
        {
            var element = info.Element(ns + "CustomDatas")!;
            if (element.Elements(ns + "CustomData").All(e =>
                e.Attribute("Name") is XAttribute name && !string.IsNullOrEmpty(name.Value) && !e.HasElements))
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
            if (versions.Items.Count == 0) return;
            if (!OfdVersions.Recognized(versions.Root)) throw new ArgumentException("Invalid version declaration.");
            var ids = new HashSet<string>(StringComparer.Ordinal);
            foreach (var version in versions.Items)
            {
                if (!ids.Add(version.Id)) throw new ArgumentException("Duplicate version ID.");
                try { XmlConvert.VerifyNCName(version.Id); }
                catch (XmlException error) { throw new ArgumentException("Version ID must be an XML ID.", error); }
                if (versions.LocationChanged(version))
                    ExistingPath("Version BaseLoc", version.BaseLocation);
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
            if (package.DocumentEntryPath is string selected && package.PreservedEntries.TryGetValue("OFD.xml", out var rootBytes))
            {
                var root = material.ReadXml("OFD.xml", rootBytes, token);
                if (root.Root?.Elements(root.Root.Name.Namespace + "DocBody").Skip(1).Any() == true)
                {
                    var directory = selected.Split('/')[0] + "/";
                    if (!resolved.StartsWith(directory, StringComparison.OrdinalIgnoreCase))
                        throw new NotSupportedException(role + " reference crosses the selected document directory.");
                }
            }
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
        if (model is OfdVersions originalVersions && clone is OfdVersions clonedVersions)
            clonedVersions.CopyLocationBaselines(originalVersions);
        return (T)clone;
    }
}
