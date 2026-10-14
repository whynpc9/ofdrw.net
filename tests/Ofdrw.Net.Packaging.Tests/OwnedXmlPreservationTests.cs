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
}
