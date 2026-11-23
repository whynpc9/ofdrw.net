using Ofdrw.Net.Core.Models;
using Xunit;

namespace Ofdrw.Net.Core.Tests;

/// <summary>The one catalogue of how package entries are located; consumers no longer keep their own copies.</summary>
public sealed class OfdReferenceCatalogTests
{
    [Theory]
    [InlineData("Doc_0/Pages/Page_0/Content.xml", "Doc_0/Pages/Page_0/PageRes.xml")]
    [InlineData("Doc_0/Tpls/Tpl_0/Content.xml", "Doc_0/Tpls/Tpl_0/PageRes.xml")]
    // A root-level page used to become "/PageRes.xml" in the pruner and "PageRes.xml" in signing selection.
    [InlineData("Content.xml", "PageRes.xml")]
    public void ImplicitPageResIsTheSiblingOfThePageContent(string page, string expected) =>
        Assert.Equal(expected, OfdReferenceCatalog.ImplicitPageResPath(page));

    [Fact]
    public void DocumentPathPrefersTheSourceEntryAndFallsBackToTheDocumentFolder()
    {
        var package = new OfdDocumentPackage { Options = { DocumentId = "Doc_7" } };
        Assert.Equal("Doc_7/Document.xml", OfdReferenceCatalog.DocumentPath(package));
        package.DocumentEntryPath = "Custom/Doc.xml";
        Assert.Equal("Custom/Doc.xml", OfdReferenceCatalog.DocumentPath(package));
    }

    [Fact]
    public void IndexDeclarationsAndControlLocationsAreDocumentVocabulary()
    {
        foreach (var name in OfdReferenceCatalog.DocumentIndexDeclarations)
        {
            Assert.True(OfdReferenceCatalog.IsDocumentChild(name), name);
            Assert.Contains(name, OfdReferenceCatalog.ControlLocationElements);
        }
        Assert.Equal("Extensions", OfdReferenceCatalog.DocumentChildren[^1]);
    }
}
