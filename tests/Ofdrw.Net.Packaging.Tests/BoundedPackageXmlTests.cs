using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Xml.Linq;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Preflight;
using Xunit;

namespace Ofdrw.Net.Packaging.Tests;

/// <summary>The one bounded XML reader and reference policy shared by Preflight, Archival and Medical.</summary>
public sealed class BoundedPackageXmlTests
{
    private static OfdPackageArchive Archive(params (string Name, string Xml)[] entries)
    {
        var map = new Dictionary<string, byte[]>();
        foreach (var (name, xml) in entries) map[name] = Encoding.UTF8.GetBytes(xml);
        return new OfdPackageArchive(map);
    }

    [Fact]
    public void DistinguishesAbsentMalformedUnsupportedEncodingAndDtd()
    {
        var reader = new BoundedPackageXml(Archive(
            ("ok.xml", "<a/>"), ("bad.xml", "<a>"), ("dtd.xml", "<!DOCTYPE a [<!ENTITY x 'y'>]><a>&x;</a>"),
            ("enc.xml", "<?xml version=\"1.0\" encoding=\"x-no-such-encoding\"?><a/>")),
            new ArchivalPreflightOptions(), CancellationToken.None);

        Assert.Equal(BoundedXmlStatus.Parsed, reader.Load("ok.xml").Status);
        Assert.Equal(BoundedXmlStatus.Absent, reader.Load("missing.xml").Status);
        Assert.Equal(BoundedXmlStatus.Malformed, reader.Load("bad.xml").Status);
        Assert.Equal(BoundedXmlStatus.Malformed, reader.Load("dtd.xml").Status);
        Assert.Equal(BoundedXmlStatus.UnsupportedEncoding, reader.Load("enc.xml").Status);
    }

    [Fact]
    public void ChargesEachDistinctEntryOnceByCaseInsensitiveIdentity()
    {
        var options = new ArchivalPreflightOptions { MaxXmlBytes = 8 };
        var reader = new BoundedPackageXml(Archive(("A.xml", "<a></a>")), options, CancellationToken.None);
        Assert.NotNull(reader.Read("A.xml"));
        // A lexical re-read of the same bytes and a cached re-parse are free.
        Assert.NotNull(reader.Bytes("A.xml"));
        Assert.NotNull(reader.Read("A.xml"));
        var other = new BoundedPackageXml(Archive(("A.xml", "<a></a>"), ("B.xml", "<b/>")), options, CancellationToken.None);
        other.Read("A.xml");
        var ex = Assert.Throws<PackageBudgetExceededException>(() => other.Read("B.xml"));
        Assert.Equal("XmlBytes", ex.Budget);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void NodeChargingModeDecidesWhetherAttributesCount(bool chargeAttributes, bool exhausted)
    {
        var charging = chargeAttributes ? XmlNodeCharging.NodesAndAttributes : XmlNodeCharging.Nodes;
        var options = new ArchivalPreflightOptions { MaxXmlNodes = 3 };
        var reader = new BoundedPackageXml(Archive(("a.xml", "<a x=\"1\" y=\"2\"><b/></a>")), options, CancellationToken.None, charging);
        if (exhausted) Assert.Equal("XmlNodes", Assert.Throws<PackageBudgetExceededException>(() => reader.Read("a.xml")).Budget);
        else Assert.NotNull(reader.Read("a.xml"));
    }

    [Fact]
    public void DepthAndReferenceBudgetsAreEnforced()
    {
        var reader = new BoundedPackageXml(Archive(("deep.xml", "<a><b><c/></b></a>")),
            new ArchivalPreflightOptions { MaxXmlDepth = 1, MaxReferences = 1 }, CancellationToken.None);
        Assert.Equal("XmlNodes", Assert.Throws<PackageBudgetExceededException>(() => reader.Read("deep.xml")).Budget);
        reader.Visit();
        Assert.Equal("Reference", Assert.Throws<PackageBudgetExceededException>(() => reader.Visit()).Budget);
        Assert.False(reader.TryVisit());
    }

    [Theory]
    [InlineData("Pages/Page_0/Content.xml", true, "Doc_0/Pages/Page_0/Content.xml")]
    [InlineData("  Res/a.ttf  ", true, "Doc_0/Res/a.ttf")]
    [InlineData("/Doc_0/Document.xml", true, "Doc_0/Document.xml")]
    [InlineData("../OFD.xml", true, "OFD.xml")]
    [InlineData("../../escape.xml", false, "")]
    [InlineData("Res/a%20b.ttf", false, "")]
    [InlineData("//server/share.xml", false, "")]
    [InlineData("Res\\a.ttf", false, "")]
    [InlineData("http://example.invalid/a.xml", false, "")]
    [InlineData("a.xml?x=1", false, "")]
    [InlineData("a.xml#f", false, "")]
    [InlineData("", false, "")]
    public void OneSafeReferencePolicyForEveryAssessor(string reference, bool ok, string expected)
    {
        Assert.Equal(ok, OfdPackageReference.TryResolve("Doc_0/Document.xml", reference, out var resolved));
        if (ok) Assert.Equal(expected, resolved);
    }

    [Fact]
    public void DirectoryReferencesMayResolveToTheRoot()
    {
        Assert.True(OfdPackageReference.TryResolveDirectory("Doc_0/Res/PublicRes.xml", "../..", out var root));
        Assert.Equal(string.Empty, root);
        Assert.False(OfdPackageReference.TryResolve("Doc_0/Res/PublicRes.xml", "../..", out _));
    }

    [Fact]
    public void SelectorNamesUnsupportedNamespacesWithoutLeakingThem()
    {
        var root = XElement.Parse("<x:OFD xmlns:x='urn:other'><x:DocBody/><x:DocBody/></x:OFD>");
        Assert.Equal("/{<unsupported-namespace>}OFD[1]/{<unsupported-namespace>}DocBody[2]",
            OfdPackageReference.Selector(((XElement)root.LastNode!)));
    }
}
