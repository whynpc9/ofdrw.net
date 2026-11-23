using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;
using Xunit;

namespace Ofdrw.Net.Packaging.Tests;

/// <summary>
/// The shared package-scope walk, tested over small in-memory entries instead of through a whole assessment:
/// each gap must land in exactly the areas it could contribute to.
/// </summary>
public sealed class OfdPackageScopeTests
{
    private const string Ns = "http://www.ofdspec.org/2016";

    private static Dictionary<string, string> Complete() => new Dictionary<string, string>
    {
        ["OFD.xml"] = $"<ofd:OFD xmlns:ofd='{Ns}' DocType='OFD'><ofd:DocBody><ofd:DocRoot>Doc_0/Document.xml</ofd:DocRoot>" +
                      "<ofd:Signatures>Doc_0/Signs/Signatures.xml</ofd:Signatures></ofd:DocBody></ofd:OFD>",
        ["Doc_0/Document.xml"] = $"<ofd:Document xmlns:ofd='{Ns}'><ofd:CommonData><ofd:PublicRes>PublicRes.xml</ofd:PublicRes>" +
                                 "<ofd:TemplatePage ID='9' BaseLoc='Tpls/Tpl_0/Content.xml'/></ofd:CommonData>" +
                                 "<ofd:Pages><ofd:Page ID='1' BaseLoc='Pages/Page_0/Content.xml'/></ofd:Pages>" +
                                 "<ofd:Annotations>Annots/Annotations.xml</ofd:Annotations>" +
                                 "<ofd:Attachments>Attachs/Attachments.xml</ofd:Attachments>" +
                                 "<ofd:CustomTags>Tags/CustomTags.xml</ofd:CustomTags></ofd:Document>",
        ["Doc_0/PublicRes.xml"] = $"<ofd:Res xmlns:ofd='{Ns}' BaseLoc='Res'><ofd:Fonts><ofd:Font ID='2'><ofd:FontFile>a.ttf</ofd:FontFile></ofd:Font></ofd:Fonts></ofd:Res>",
        ["Doc_0/Pages/Page_0/Content.xml"] = $"<ofd:Page xmlns:ofd='{Ns}'/>",
        ["Doc_0/Tpls/Tpl_0/Content.xml"] = $"<ofd:Page xmlns:ofd='{Ns}'/>",
        ["Doc_0/Annots/Annotations.xml"] = $"<ofd:Annotations xmlns:ofd='{Ns}'><ofd:Page PageID='1'><ofd:FileLoc>Page_0/Annotation.xml</ofd:FileLoc></ofd:Page></ofd:Annotations>",
        ["Doc_0/Attachs/Attachments.xml"] = $"<ofd:Attachments xmlns:ofd='{Ns}'><ofd:Attachment ID='3' Format='json'><ofd:FileLoc>data.json</ofd:FileLoc></ofd:Attachment></ofd:Attachments>",
        ["Doc_0/Tags/CustomTags.xml"] = $"<ofd:CustomTags xmlns:ofd='{Ns}'><ofd:CustomTag><ofd:FileLoc>tag.xml</ofd:FileLoc></ofd:CustomTag></ofd:CustomTags>",
        ["Doc_0/Signs/Signatures.xml"] = $"<ofd:Signatures xmlns:ofd='{Ns}'><ofd:Signature ID='1' BaseLoc='Sign_0/Signature.xml'/></ofd:Signatures>"
    };

    private static OfdPackageScope Walk(Dictionary<string, string> entries, ArchivalPreflightOptions? options = null) =>
        OfdPackageScope.Walk(new BoundedPackageXml(
            new OfdPackageArchive(entries.ToDictionary(x => x.Key, x => Encoding.UTF8.GetBytes(x.Value))),
            options ?? new ArchivalPreflightOptions(), CancellationToken.None));

    private static readonly OfdScopeArea[] Areas =
    {
        OfdScopeArea.Document, OfdScopeArea.Pages, OfdScopeArea.Resources, OfdScopeArea.Signatures,
        OfdScopeArea.Annotations, OfdScopeArea.CustomTags, OfdScopeArea.Attachments
    };

    private static IEnumerable<OfdScopeArea> Incomplete(OfdPackageScope scope) => Areas.Where(x => !scope.Verdict(x).Complete);

    [Fact]
    public void ACompletePackageHasNoGaps()
    {
        var scope = Walk(Complete());
        Assert.Empty(Incomplete(scope));
        Assert.True(scope.Verdict().Complete);
        Assert.Null(scope.Verdict().FirstGap);
        Assert.Equal(new[] { "Doc_0/Pages/Page_0/Content.xml" }, scope.PagePaths);
        Assert.Equal(new[] { false, true }, scope.Pages.Select(x => x.IsTemplate));
        Assert.Equal("Doc_0/Res/a.ttf", scope.Locations.Single(x => x.Kind == OfdScopeLocationKind.ResourcePayload).Resolved);
        Assert.Equal(new[] { "Doc_0/Signs/Sign_0/Signature.xml" }, scope.SignaturePaths);
    }

    [Fact]
    public void AMissingPageLeavesPagesAndResourcesUnknownButNothingElse()
    {
        var entries = Complete();
        entries.Remove("Doc_0/Pages/Page_0/Content.xml");
        var scope = Walk(entries);
        Assert.Equal(new[] { OfdScopeArea.Pages, OfdScopeArea.Resources }, Incomplete(scope));
        Assert.Equal("Doc_0/Pages/Page_0/Content.xml", scope.Verdict(OfdScopeArea.Resources).FirstGap);
    }

    [Fact]
    public void AnUnavailableDocumentLeavesEveryDocumentAreaUnknownWithItsPathAsEvidence()
    {
        var entries = Complete();
        entries["Doc_0/Document.xml"] = "<not-well-formed";
        var scope = Walk(entries);
        Assert.True(scope.DocumentUnavailable);
        Assert.Equal(Areas.Where(x => x != OfdScopeArea.Signatures), Incomplete(scope));
        foreach (var area in Incomplete(scope)) Assert.Equal("Doc_0/Document.xml", scope.Verdict(area).FirstGap);
    }

    [Theory]
    [InlineData("<ofd:Page PageID='1'><ofd:FileLoc>Page_0/%41nnotation.xml</ofd:FileLoc></ofd:Page>")]
    [InlineData("")]
    public void AnnotationGapsStayInTheAnnotationArea(string page)
    {
        var entries = Complete();
        if (page.Length == 0) entries.Remove("Doc_0/Annots/Annotations.xml");
        else entries["Doc_0/Annots/Annotations.xml"] = $"<ofd:Annotations xmlns:ofd='{Ns}'>{page}</ofd:Annotations>";
        Assert.Equal(new[] { OfdScopeArea.Annotations }, Incomplete(Walk(entries)));
    }

    [Fact]
    public void UnresolvedAttachmentKeepsItsDeclaredFormat()
    {
        var entries = Complete();
        entries["Doc_0/Attachs/Attachments.xml"] =
            $"<ofd:Attachments xmlns:ofd='{Ns}'><ofd:Attachment ID='3' Format='json' FileLoc='..\\x.json'/></ofd:Attachments>";
        var gap = Assert.Single(Walk(entries).Verdict(OfdScopeArea.Attachments).Gaps);
        Assert.Equal(OfdScopeLocationKind.Attachment, gap.Kind);
        Assert.Equal("json", gap.Format);
    }

    [Fact]
    public void UnresolvedSignatureBaseLocIsASignatureGap()
    {
        var entries = Complete();
        entries["Doc_0/Signs/Signatures.xml"] = $"<ofd:Signatures xmlns:ofd='{Ns}'><ofd:Signature ID='1' BaseLoc='http://x/s.xml'/></ofd:Signatures>";
        var scope = Walk(entries);
        Assert.Equal(new[] { OfdScopeArea.Signatures }, Incomplete(scope));
        Assert.Equal(new string?[] { null }, scope.SignatureDeclarations);
    }

    [Fact]
    public void TheImplicitSiblingPageResIsAPageResourceTable()
    {
        var entries = Complete();
        entries["Doc_0/Pages/Page_0/PageRes.xml"] = $"<ofd:Res xmlns:ofd='{Ns}'><ofd:Fonts><ofd:Font ID='5'/></ofd:Fonts></ofd:Res>";
        var table = Assert.Single(Walk(entries).Tables, x => x.Scope == OfdResourceScope.Page);
        Assert.Equal("Doc_0/Pages/Page_0/PageRes.xml", table.Path);
    }

    [Fact]
    public void ADeclaredSiblingPageResIsNotCountedTwice()
    {
        var entries = Complete();
        entries["Doc_0/Pages/Page_0/Content.xml"] = $"<ofd:Page xmlns:ofd='{Ns}'><ofd:PageRes>PageRes.xml</ofd:PageRes></ofd:Page>";
        entries["Doc_0/Pages/Page_0/PageRes.xml"] = $"<ofd:Res xmlns:ofd='{Ns}'/>";
        Assert.Single(Walk(entries).Tables, x => x.Scope == OfdResourceScope.Page);
    }

    [Fact]
    public void AmbiguousDocRootIsADocumentGap()
    {
        var entries = Complete();
        entries["OFD.xml"] = entries["OFD.xml"].Replace("</ofd:DocRoot>", "</ofd:DocRoot><ofd:DocRoot>Doc_1/Document.xml</ofd:DocRoot>");
        var scope = Walk(entries);
        Assert.Null(scope.DocRoot);
        Assert.Equal("OFD.xml", scope.Verdict(OfdScopeArea.Pages).FirstGap);
        Assert.True(scope.Verdict(OfdScopeArea.Signatures).Complete);
    }

    [Fact]
    public void TheWalkChargesTheCallersReferenceBudget()
    {
        var ex = Assert.Throws<PackageBudgetExceededException>(() => Walk(Complete(), new ArchivalPreflightOptions { MaxReferences = 2 }));
        Assert.Equal("Reference", ex.Budget);
    }
}
