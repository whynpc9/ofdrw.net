using System.Xml.Linq;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Tests;

public sealed class OfdNavigationTests
{
    [Fact]
    public void RemovingOwnedItemKeepsSurvivorsAfterOpaqueSourceAnchor()
    {
        var package = Read("<Outlines><OutlineElem Title='A'/><v:Anchor/><OutlineElem Title='B'/></Outlines>");
        package.Outlines!.Items.RemoveAt(0);
        var root = Write(package).Single();
        Assert.Equal(new[] { "Anchor", "OutlineElem" }, root.Elements().Select(node => node.Name.LocalName));
        Assert.Equal("B", root.Elements().Last().Attribute("Title")!.Value);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void UnclaimedActionsCannotKeepRemovedOrRenamedBookmark(bool rename)
    {
        var package = Read("<Actions><Action><Goto><Bookmark Name='B'/></Goto></Action></Actions><Bookmarks><Bookmark Name='B'><Dest Type='Fit' PageID='1'/></Bookmark></Bookmarks>");
        package.Actions = null;
        if (rename) package.Bookmarks!.Items[0].Name = "Changed"; else package.Bookmarks!.Items.Clear();
        Assert.Throws<InvalidOperationException>(() => OfdNavigationXml.Validate(package, package.Pages, default));
    }

    [Fact]
    public void ModifiedCollectionCannotBeWithdrawnAsUnclaimed()
    {
        var package = Read("<Outlines><OutlineElem Title='A'/></Outlines>");
        package.Outlines!.Items[0].Title = "Edited";
        Assert.Throws<InvalidOperationException>(() => package.Outlines = null);
    }

    [Fact]
    public void PristineWithdrawalPreservesSource()
    {
        var package = Read("<Outlines v:Value='v:Token'><OutlineElem Title='A'/><v:Anchor/></Outlines>");
        package.Outlines = null;
        var root = Write(package).Single();
        Assert.Equal("v:Token", root.Attribute(XName.Get("Value", "urn:vendor"))!.Value);
        Assert.Equal("urn:vendor", root.GetNamespaceOfPrefix("v")!.NamespaceName);
        Assert.Single(root.Elements(XName.Get("Anchor", "urn:vendor")));
    }

    [Theory]
    [InlineData("<Outlines><OutlineElem/></Outlines>")]
    [InlineData("<Bookmarks><Bookmark Name='B'><Dest Type='0' PageID='1'/></Bookmark></Bookmarks>")]
    public void FailedKnownProjectionStillBlocksPageDeletionWhenCollectionReplaced(string xml)
    {
        var package = Read(xml);
        if (package.Outlines is not null) package.Outlines = new OfdOutlineCollection();
        if (package.Bookmarks is not null) package.Bookmarks = new OfdBookmarkCollection();
        Assert.Throws<NotSupportedException>(() => OfdNavigationXml.Validate(package, Array.Empty<OfdPage>(), default));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ClonePreservesDirtyBaselineAndMapsActualPages(bool edited)
    {
        var source = Read("<Outlines><OutlineElem Title='A'><Actions><Action><Goto><Dest Type='Fit' PageID='1'/></Goto></Action></Actions></OutlineElem></Outlines>");
        if (edited) source.Outlines!.Items[0].Title = "Edited";
        var destination = new OfdDocumentPackage(); var page = new OfdPage(); destination.Pages.Add(page);
        var mapping = new Dictionary<object, object> { [source.Pages[0]] = page };
        OfdNavigationCloner.Clone(source, destination, mapping, new OfdModelMaterialization(new OfdModelLimits()), default);
        Assert.Same(page, Assert.IsType<OfdGoToAction>(destination.Outlines!.Items[0].Actions!.Items[0]).Destination!.Page);
        if (edited) Assert.Throws<InvalidOperationException>(() => destination.Outlines = null);
        else { destination.Outlines = null; Assert.Single(Write(destination)); }
    }

    [Fact]
    public void RawEventCanBeClearedWithoutExecutingOrDroppingUnknownBody()
    {
        var package = Read("<Actions><Action Event='DO'><URI Base='https://example.invalid'/></Action></Actions>");
        Assert.IsType<OfdRawAction>(package.Actions!.Items[0]).Event = null;
        var action = Write(package).Single().Elements().Single();
        Assert.Null(action.Attribute("Event")); Assert.Single(action.Elements(action.Name.Namespace + "URI"));
    }

    [Theory]
    [InlineData("<Goto><?vendor PageID='2'?><Dest Type='Fit' PageID='1'/></Goto>")]
    [InlineData("<Goto>opaque<Dest Type='Fit' PageID='1'/></Goto>")]
    [InlineData("<Goto><Bookmark Name='B'><?vendor PageID='2'?></Bookmark></Goto>")]
    public void OpaqueGoToWrappersPreventDeletion(string gotoXml)
    {
        var package = Read("<Actions><Action>" + gotoXml + "</Action></Actions><Bookmarks><Bookmark Name='B'><Dest Type='Fit' PageID='1'/></Bookmark></Bookmarks>");
        package.Pages.Add(new OfdPage { Id = "2" });
        Assert.Throws<NotSupportedException>(() => OfdNavigationXml.Validate(package, new[] { package.Pages[0] }, default));
    }

    [Fact]
    public void OpaqueBookmarkNameCannotBeClaimedByNewTypedBookmark()
    {
        var package = Read("<Bookmarks><Bookmark Name='B'><Dest Type='UNKNOWN' PageID='1'/></Bookmark></Bookmarks>");
        package.Bookmarks!.Items.Add(new OfdBookmark { Name = "B", Destination = new OfdDestination { Page = package.Pages[0] } });
        Assert.Throws<ArgumentException>(() => OfdNavigationXml.Validate(package, package.Pages, default));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void DuplicateBookmarkNamesAreRejectedEvenWhenOneIsOpaque(bool opaqueFirst)
    {
        var opaque = "<Bookmark Name='B'><Dest Type='UNKNOWN' PageID='1'/></Bookmark>";
        var typed = "<Bookmark Name='B'><Dest Type='Fit' PageID='1'/></Bookmark>";
        Assert.Throws<System.IO.InvalidDataException>(() => Read("<Bookmarks>" + (opaqueFirst ? opaque + typed : typed + opaque) + "</Bookmarks>"));
    }

    private static OfdDocumentPackage Read(string navigation)
    {
        var package = new OfdDocumentPackage(); package.Pages.Add(new OfdPage { Id = "1" });
        var material = new OfdModelMaterialization(new OfdModelLimits());
        OfdNavigationXml.Read(package, XElement.Parse($"<Document xmlns='{package.Options.Namespace}' xmlns:v='urn:vendor'>{navigation}</Document>"), material, default);
        return package;
    }
    private static IReadOnlyList<XElement> Write(OfdDocumentPackage package)
        => OfdNavigationXml.Write(package, package.Pages.ToDictionary(page => page, _ => "1"), package.Options.Namespace, default);
}
