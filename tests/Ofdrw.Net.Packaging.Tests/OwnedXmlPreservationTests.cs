using System.Text;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class OwnedXmlPreservationTests
{
    private const string Source = "<Document xmlns='urn:ofd' xmlns:v='urn:vendor' v:first='1' v:second='2'><DocInfo><Title>Original</Title></DocInfo><Pages><Page ID='1' BaseLoc='Pages/P1.xml'/></Pages><v:A Value='opaque'><![CDATA[visible raw]]></v:A><?vendor keep?><!--keep--><v:B/></Document>";
    [Fact]
    public void OnlyStructuralIndentationCanChangeWithoutAnExplicitPlanPermission()
    {
        using var context = new OfdControlledEditContext(new());
        OfdOwnedXmlPreservation.RequireExactClosedRewrite("Document.xml", XDocument.Parse(Source, LoadOptions.PreserveWhitespace),
            Encoding.UTF8.GetBytes(Source.Replace("<DocInfo>", "\n  <DocInfo>").Replace("</Pages>", "</Pages>\n")), context.Material, default);
    }
    [Theory]
    [InlineData("Original", "Changed")]
    [InlineData("Value='opaque'", "Value='changed'")]
    [InlineData("v:first='1'", "v:first='9'")]
    [InlineData("v:first='1' v:second='2'", "v:second='2' v:first='1'")]
    [InlineData("v:second='2'", "v:second='2' v:added='new'")]
    [InlineData("v:second='2'", "")]
    [InlineData("<![CDATA[visible raw]]>", "visible raw")]
    [InlineData("<?vendor keep?>", "<?vendor changed?>")]
    [InlineData("<?vendor keep?>", "")]
    [InlineData("<!--keep-->", "")]
    [InlineData("<v:B/>", "<v:B/><v:new/>")]
    [InlineData("<!--keep--><v:B/>", "<v:B/><!--keep-->")]
    [InlineData("<Pages><Page ID='1' BaseLoc='Pages/P1.xml'/></Pages>", "<Pages><Page ID='1' BaseLoc='Pages/P2.xml'/></Pages>")]
    public void AttributesKnownScalarsOpaqueNodesAndOrderedRepresentationRequireExactPermission(string from, string to)
    {
        using var context = new OfdControlledEditContext(new());
        Assert.Throws<InvalidDataException>(() => OfdOwnedXmlPreservation.RequireExactClosedRewrite("Document.xml",
            XDocument.Parse(Source, LoadOptions.PreserveWhitespace), Encoding.UTF8.GetBytes(Source.Replace(from, to)), context.Material, default));
    }
    [Fact]
    public void UnselectedTextSizeFillTransformAndTextStayProtected()
    {
        const string source = "<Page xmlns='urn:ofd'><Content><Layer ID='2' Type='Body'><TextObject ID='3' Size='4' CTM='1 0 0 1 0 0'><FillColor Value='0 0 0'/><TextCode X='0' Y='4'>Unselected</TextCode></TextObject></Layer></Content></Page>";
        foreach (var changed in new[] { source.Replace("Size='4'", "Size='9'"), source.Replace("1 0 0 1 0 0", "2 0 0 2 0 0"), source.Replace("0 0 0", "1 1 1"), source.Replace("Unselected", "Wrong") })
        {
            using var context = new OfdControlledEditContext(new());
            Assert.Throws<InvalidDataException>(() => OfdOwnedXmlPreservation.RequireExactClosedRewrite("Page.xml",
                XDocument.Parse(source, LoadOptions.PreserveWhitespace), Encoding.UTF8.GetBytes(changed), context.Material, default));
        }
    }
    [Theory]
    [InlineData("<Content><![CDATA[ ]]><Layer ID='2'/></Content>", "<Content><Layer ID='2'/></Content>")]
    [InlineData("<v:Opaque><Content> <Layer ID='2'/></Content></v:Opaque>", "<v:Opaque><Content><Layer ID='2'/></Content></v:Opaque>")]
    [InlineData("<Extension><Content> <Layer ID='2'/></Content></Extension>", "<Extension><Content><Layer ID='2'/></Content></Extension>")]
    [InlineData("<Content xml:space='preserve'> <Layer ID='2'/></Content>", "<Content xml:space='preserve'><Layer ID='2'/></Content>")]
    public void WhitespaceCdataAndOpaqueOrPreservedWhitespaceStayExact(string original, string changed)
    {
        const string prefix = "<Page xmlns='urn:ofd' xmlns:v='urn:vendor'>";
        using var context = new OfdControlledEditContext(new());
        Assert.Throws<InvalidDataException>(() => OfdOwnedXmlPreservation.RequireExactClosedRewrite("Page.xml",
            XDocument.Parse(prefix + original + "</Page>", LoadOptions.PreserveWhitespace),
            Encoding.UTF8.GetBytes(prefix + changed + "</Page>"), context.Material, default));
    }
    [Fact]
    public void ActualPageAreaIndentationIsStructuralAndItsWhitespaceScanIsCharged()
    {
        const string original = "<Page xmlns='urn:ofd'><Area> <PhysicalBox>5 7 200 200</PhysicalBox> </Area><Content><Layer ID='2'/></Content></Page>";
        using var shortContext = new OfdControlledEditContext(new());
        using var longContext = new OfdControlledEditContext(new());
        var output = Encoding.UTF8.GetBytes(original.Replace("> <", "><"));
        OfdOwnedXmlPreservation.RequireExactClosedRewrite("Page.xml", XDocument.Parse(original, LoadOptions.PreserveWhitespace), output, shortContext.Material, default);
        const int extraCharacters = 10000;
        var longer = original.Replace("<Area> ", "<Area> " + new string(' ', extraCharacters));
        OfdOwnedXmlPreservation.RequireExactClosedRewrite("Page.xml", XDocument.Parse(longer, LoadOptions.PreserveWhitespace), output, longContext.Material, default);
        Assert.True(longContext.Material.ValidationWorkBytes - shortContext.Material.ValidationWorkBytes >= 2L * extraCharacters);
    }
    [Fact]
    public void PageReferenceDoesNotAcquirePageRootContentPermission()
    {
        const string source = "<Document xmlns='urn:ofd'><Pages><Page ID='1'><Content> <Layer ID='2'/></Content></Page></Pages></Document>";
        using var context = new OfdControlledEditContext(new());
        Assert.Throws<InvalidDataException>(() => OfdOwnedXmlPreservation.RequireExactClosedRewrite("Document.xml",
            XDocument.Parse(source, LoadOptions.PreserveWhitespace), Encoding.UTF8.GetBytes(source.Replace("<Content> ", "<Content>")), context.Material, default));
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImportedFragmentDropsOnlyReifiedSourceAncestorBindings(bool defaultNamespace)
    {
        var source = XElement.Parse(defaultNamespace
            ? "<Page xmlns='urn:ofd'><Content><TextObject ID='3'/></Content></Page>"
            : "<ofd:Page xmlns:ofd='urn:ofd'><ofd:Content><ofd:TextObject ID='3'/></ofd:Content></ofd:Page>");
        var root = source.Descendants().Last();
        var fragment = XElement.Parse(root.ToString(SaveOptions.DisableFormatting));
        Assert.Contains(fragment.Attributes(), attribute => attribute.IsNamespaceDeclaration);
        using var context = new OfdControlledEditContext(new());
        OfdOwnedXmlPreservation.RemoveReifiedInheritedNamespaces(fragment, root, context.Material, default);
        Assert.DoesNotContain(fragment.Attributes(), attribute => attribute.IsNamespaceDeclaration);
        Assert.Equal(root.Name, fragment.Name);
        Assert.Equal("3", fragment.Attribute("ID")?.Value);
    }
    [Fact]
    public void OriginalOrNonInheritedNamespaceDeclarationsStayProtected()
    {
        var source = XElement.Parse("<ofd:Page xmlns:ofd='urn:ofd'><ofd:Content><ofd:TextObject ID='3' xmlns:v='urn:vendor'/></ofd:Content></ofd:Page>");
        var root = source.Descendants().Last();
        var fragment = XElement.Parse("<ofd:TextObject xmlns:ofd='urn:ofd' ID='3' xmlns:v='urn:vendor' xmlns:new='urn:new'/>");
        using var context = new OfdControlledEditContext(new());
        OfdOwnedXmlPreservation.RemoveReifiedInheritedNamespaces(fragment, root, context.Material, default);
        Assert.Equal("urn:vendor", fragment.Attribute(XNamespace.Xmlns + "v")?.Value);
        Assert.Equal("urn:new", fragment.Attribute(XNamespace.Xmlns + "new")?.Value);
        Assert.Null(fragment.Attribute(XNamespace.Xmlns + "ofd"));
    }
}
