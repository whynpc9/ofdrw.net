using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Preflight;

// A handler is deliberately internal until a caller-owned rule registration and trust contract exists.
internal interface IArchivalRuleHandler
{
    ArchivalRuleDefinition Rule { get; }
    bool Apply(XElement document);
}

internal static class ArchivalRuleHandlers
{
    // New catalog rules must not inherit a destructive handler by accident.
    internal static IReadOnlyList<IArchivalRuleHandler> Supported { get; } = Array.AsReadOnly(new IArchivalRuleHandler[]
    {
        new RemoveDocumentDeclarationHandler(ArchivalRuleCatalog.Rules.Single(rule => rule.RuleId == "AR-001")),
        new RemoveDocumentDeclarationHandler(ArchivalRuleCatalog.Rules.Single(rule => rule.RuleId == "AR-002"))
    });
}

internal sealed class RemoveDocumentDeclarationHandler : IArchivalRuleHandler
{
    internal RemoveDocumentDeclarationHandler(ArchivalRuleDefinition rule) { Rule = rule; }
    public ArchivalRuleDefinition Rule { get; }

    public bool Apply(XElement document)
    {
        var nodes = document.Elements(document.Name.Namespace + Rule.Element).ToArray();
        if (nodes.Length > 1) throw new InvalidDataException("Duplicate " + Rule.Element + " declarations require manual review.");
        if (nodes.Length == 0) return false;
        if (!SafeToRemove(nodes[0])) throw new InvalidDataException("Unknown or unsupported " + Rule.Element + " content would be lost.");
        nodes[0].Remove();
        return true;
    }

    private bool SafeToRemove(XElement node)
    {
        if (node.Attributes().Any(x => !x.IsNamespaceDeclaration) ||
            node.Nodes().Any(x => x is not XElement && x is not XText || x is XText text && !string.IsNullOrWhiteSpace(text.Value)))
            return false;
        if (Rule.Element == "Permissions" && !OfdPermissions.Recognized(node)) return false;
        if (Rule.Element == "VPreferences" && !OfdViewingPreferences.Recognized(node)) return false;
        var allowed = Rule.Element == "Permissions"
            ? new[] { "Edit", "Annot", "Export", "Signature", "Watermark", "PrintScreen", "Print", "ValidPeriod" }
            : new[] { "PageMode", "PageLayout", "TabDisplay", "HideToolbar", "HideMenubar", "HideWindowUI", "ZoomMode", "Zoom" };
        foreach (var child in node.Elements())
        {
            if (child.Name.Namespace != node.Name.Namespace || !allowed.Contains(child.Name.LocalName, StringComparer.Ordinal) || child.HasElements ||
                child.Nodes().Any(x => x is not XText) ||
                (child.Name.LocalName == "Print" || child.Name.LocalName == "ValidPeriod") && !string.IsNullOrWhiteSpace(child.Value)) return false;
            var attributes = child.Name.LocalName == "Print" ? new[] { "Printable", "Copies" } :
                child.Name.LocalName == "ValidPeriod" ? new[] { "StartDate", "EndDate" } : Array.Empty<string>();
            if (child.Attributes().Any(x => x.IsNamespaceDeclaration || x.Name.Namespace != XNamespace.None ||
                !attributes.Contains(x.Name.LocalName, StringComparer.Ordinal))) return false;
        }
        return true;
    }
}
