using System.Collections;
using System.Text;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Tests;

public sealed class OfdModelTraversalTests
{
    [Fact]
    public void SelectedOrderAndPaintOrderArePreservedWithoutChangingModel()
    {
        var (package, first) = Package(Text("body"));
        var background = new OfdTemplateContent { TemplateId = "bg", BaseLocation = "Templates/bg.xml" };
        background.Elements.Add(Text("background")); first.Templates.Add(background);
        var foreground = new OfdTemplateContent { TemplateId = "fg", ZOrder = "Foreground" };
        foreground.Elements.Add(Text("foreground")); first.Templates.Insert(0, foreground);
        first.AnnotationAppearances.Add(Text("annotation"));
        var second = new OfdPage(); second.Elements.Add(Text("second")); package.Pages.Add(second);

        var result = Read(package, [second, first], new() { IncludeTemplates = true, IncludeAnnotations = true });

        Assert.Equal(new[] { "second", "background", "body", "foreground", "annotation" },
            result.Occurrences.Select(o => ((OfdTextElement)o.Element).Text));
        Assert.Same(first.Elements[0], result.Occurrences[2].Element);
        Assert.Equal(new[] { OfdContentScope.Body, OfdContentScope.Template, OfdContentScope.Body,
            OfdContentScope.Template, OfdContentScope.Annotation }, result.Occurrences.Select(o => o.RootScope));
        Assert.Equal(new double[] { 1, 0, 0, 1, 0, 0 }, result.Occurrences[4].ParentToPage);
        Assert.Null(first.Id); Assert.Null(first.Elements[0].ObjectId);
        Assert.Same(foreground, first.Templates[0]);
    }

    [Fact]
    public void FreshAndImportedBodyAreRealSourcesButSharedTemplatesAndDerivedAnnotationsAreNotEditable()
    {
        var fresh = Text("fresh");
        var imported = Text("imported"); imported.SourceXml = XmlText("imported");
        var (package, page) = Package(fresh, imported);
        page.SourceEntryPath = "Doc_0/Pages/Page_0/Content.xml";
        for (var i = 0; i < 2; i++)
        {
            var template = new OfdTemplateContent { TemplateId = "shared", BaseLocation = "Templates/shared.xml" };
            template.Elements.Add(Text("separately parsed")); page.Templates.Add(template);
        }
        page.AnnotationAppearances.Add(Text("derived"));
        var result = Read(package, [page], new() { IncludeTemplates = true, IncludeAnnotations = true });
        Assert.All(result.Occurrences.Where(o => o.RootScope == OfdContentScope.Body), o =>
        {
            Assert.True(o.IsEditableSource); Assert.Equal(OfdSourceAuthority.ModelDefinition, o.SourceAuthority);
        });
        Assert.All(result.Occurrences.Where(o => o.RootScope == OfdContentScope.Template), o =>
        {
            Assert.False(o.IsEditableSource); Assert.True(o.IsSharedDefinition);
            Assert.Equal("Doc_0/Templates/shared.xml", o.SourceEntryPath);
        });
        var annotation = Assert.Single(result.Occurrences, o => o.RootScope == OfdContentScope.Annotation);
        Assert.False(annotation.IsEditableSource); Assert.Null(annotation.SourceEntryPath);
        Assert.Equal(OfdSourceAuthority.DerivedAnnotationProjection, annotation.SourceAuthority);
        Assert.Same(imported.SourceXml, result.Occurrences.Single(o => ReferenceEquals(o.Element, imported)).SourceXml);
    }

    [Fact]
    public void SharedBodyDefinitionHasDistinctOccurrencesAndIsNotOccurrenceEditable()
    {
        var shared = Text("shared");
        var (package, page) = Package(shared, shared);
        var result = Read(package, [page]);
        Assert.Equal(2, result.Occurrences.Count);
        Assert.All(result.Occurrences, o => { Assert.True(o.IsSharedDefinition); Assert.False(o.IsEditableSource); });
        Assert.Equal(result.Occurrences[0].DefinitionPath, result.Occurrences[1].DefinitionPath);
        Assert.NotEqual(result.Occurrences[0].OccurrencePath, result.Occurrences[1].OccurrencePath);
    }

    [Fact]
    public void RawPageBlockAndCompositeAreDiagnosedWithoutInventingTypedChildren()
    {
        var raw = new OfdRawElement { LocalName = "PageBlock", Xml = "<PageBlock><TextObject><TextCode>private text</TextCode></TextObject></PageBlock>" };
        var composite = new OfdRawElement { LocalName = "CompositeObject", Xml = "<CompositeObject ResourceID='42'/>" };
        var (package, page) = Package(Text("known"), raw, composite);
        var result = Read(package, [page]);
        Assert.Single(result.Occurrences);
        Assert.Equal(2, result.Diagnostics.Count);
        Assert.All(result.Diagnostics, d =>
        {
            Assert.Equal(OfdTraversalDiagnosticKind.RawObject, d.Kind);
            Assert.DoesNotContain("private text", d.Message); Assert.DoesNotContain("<", d.Message);
        });
        Assert.Throws<NotSupportedException>(() => Read(package, [page], new() { RequireSupportedSource = true }));
        Assert.Same(raw, page.Elements[1]); Assert.Contains("private text", raw.Xml);
    }

    [Theory]
    [InlineData("<TextObject xmlns='urn:foreign'><TextCode>x</TextCode></TextObject>")]
    [InlineData("<textobject><TextCode>x</TextCode></textobject>")]
    [InlineData("<TextObject xmlns='http://www.ofdspec.org'><TextCode>x</TextCode><Extension/></TextObject>")]
    [InlineData("<TextObject xmlns='http://www.ofdspec.org' ObjectRef='42'><TextCode>x</TextCode></TextObject>")]
    public void NamespaceLookalikesAndUnsupportedReferencesCannotBeEditableSources(string xml)
    {
        var text = Text("x"); text.SourceXml = xml;
        var (package, page) = Package(text);
        var result = Read(package, [page]);
        Assert.False(Assert.Single(result.Occurrences).IsEditableSource);
        Assert.Equal(OfdTraversalDiagnosticKind.UnsupportedSource, Assert.Single(result.Diagnostics).Kind);
        Assert.Throws<NotSupportedException>(() => Read(package, [page], new() { RequireSupportedSource = true }));
    }

    [Fact]
    public void PrimitiveShapeSupportDoesNotClaimGlyphGeometryValidation()
    {
        var text = Text("x"); text.SourceXml = "<TextObject CTM='not-a-matrix'><TextCode>x</TextCode></TextObject>";
        var (package, page) = Package(text);
        // A0 certifies source shape only. 06 must reject this CTM when computing geometry.
        var occurrence = Assert.Single(Read(package, [page], new() { RequireSupportedSource = true }).Occurrences);
        Assert.Same(text, occurrence.Element);
    }

    [Theory]
    [InlineData("<TextObject xmlns='urn:foreign' Visible='false'><TextCode>x</TextCode></TextObject>")]
    [InlineData("<WrongRoot Visible='0'><TextCode>x</TextCode></WrongRoot>")]
    public void UnknownRootsCannotHideDiagnosticsUsingVisibleFalse(string xml)
    {
        var text = Text("x"); text.SourceXml = xml;
        var (package, page) = Package(text);
        var loose = Read(package, [page], new() { VisibleOnly = true });
        Assert.Single(loose.Diagnostics); Assert.Single(loose.Occurrences);
        Assert.Throws<NotSupportedException>(() => Read(package, [page], new() { VisibleOnly = true, RequireSupportedSource = true }));
        text.SourceXml = "<TextObject xmlns='http://www.ofdspec.org' Visible='false'><TextCode>x</TextCode></TextObject>";
        var knownHidden = Read(package, [page], new() { VisibleOnly = true, RequireSupportedSource = true });
        Assert.Empty(knownHidden.Occurrences); Assert.Empty(knownHidden.Diagnostics);
    }

    [Fact]
    public void SelectedPagesMustBeActualDistinctMembersEvenWhenIdsAndIndicesMatch()
    {
        var (package, page) = Package(Text("x"));
        Assert.Throws<ArgumentException>(() => Read(package, [page, page]));
        Assert.Throws<ArgumentException>(() => Read(package, [new OfdPage { Id = page.Id, Index = page.Index }]));
        Assert.Empty(Read(package, []).Occurrences);
    }

    [Fact]
    public void WholePackagePreflightIsOnceAndIncludesExcludedTemplatesAndPages()
    {
        var (package, page) = Package(Text("one"));
        var second = new OfdPage(); second.Elements.Add(Text("two")); package.Pages.Add(second);
        var hidden = new OfdTemplateContent(); hidden.Elements.Add(Text("not selected")); second.Templates.Add(hidden);
        var one = Read(package, [page]); var both = Read(package, [page, second]);
        Assert.Equal(one.NodeCount, both.NodeCount);
        Assert.Equal(one.ReferenceCount + 1, both.ReferenceCount);
        Assert.Throws<InvalidDataException>(() => Read(package, [page], new()
        {
            Limits = new() { MaxNodeCount = (int)one.NodeCount - 1 }
        }));
    }

    [Fact]
    public void OccurrenceAndDiagnosticLimitsAreSharedAcrossPages()
    {
        var (package, page) = Package(Text("one"));
        var second = new OfdPage(); second.Elements.Add(new OfdRawElement { Xml = "<PageBlock/>" }); package.Pages.Add(second);
        Assert.Throws<InvalidDataException>(() => Read(package, [page, second], new()
        {
            Limits = new() { MaxExpandedOccurrenceCount = 1 }
        }));
        var complete = Read(package, [page, second], new() { Limits = new() { MaxExpandedOccurrenceCount = 2 } });
        Assert.Single(complete.Occurrences); Assert.Single(complete.Diagnostics);
    }

    [Fact]
    public void NodeReferenceAndSourceByteBoundariesAreInclusive()
    {
        var text = Text("中🙂"); text.SourceXml = XmlText("中🙂");
        var (package, page) = Package(text);
        var baseline = Read(package, [page]);
        Assert.Equal(Encoding.UTF8.GetByteCount(text.SourceXml), baseline.SourceXmlBytes);
        var limits = new OfdModelLimits { MaxNodeCount = (int)baseline.NodeCount,
            MaxReferenceCount = (int)baseline.ReferenceCount, MaxSourceXmlBytes = baseline.SourceXmlBytes };
        Assert.Single(Read(package, [page], new() { Limits = limits }).Occurrences);
        limits.MaxReferenceCount--;
        Assert.Throws<InvalidDataException>(() => Read(package, [page], new() { Limits = limits }));
        limits.MaxReferenceCount++; limits.MaxSourceXmlBytes--;
        Assert.Throws<InvalidDataException>(() => Read(package, [page], new() { Limits = limits }));
    }

    [Fact]
    public void SourceByteBudgetRejectsBeforeMalformedXmlIsParsed()
    {
        var text = Text("private"); text.SourceXml = "<" + new string('x', 4096);
        var (package, page) = Package(text);
        var failure = Assert.Throws<InvalidDataException>(() => Read(package, [page], new()
        {
            Limits = new() { MaxSourceXmlBytes = 32 }
        }));
        Assert.Contains("byte budget", failure.Message); Assert.Null(failure.InnerException);
    }

    [Fact]
    public void RawDepthAndHiddenDefinitionsCannotBypassPreflight()
    {
        var nested = string.Concat(Enumerable.Repeat("<a>", 65)) + string.Concat(Enumerable.Repeat("</a>", 65));
        var (package, page) = Package(Text("known"));
        var excluded = new OfdPage(); excluded.Elements.Add(new OfdRawElement { Xml = nested }); package.Pages.Add(excluded);
        Assert.Throws<InvalidDataException>(() => Read(package, [page]));
    }

    [Fact]
    public void DtdAndMalformedSourceAreRejectedWithoutModelMutation()
    {
        var text = Text("x"); text.SourceXml = "<!DOCTYPE TextObject [<!ENTITY e 'private'>]><TextObject><TextCode>&e;</TextCode></TextObject>";
        var (package, page) = Package(text);
        var original = text.SourceXml;
        Assert.Throws<InvalidDataException>(() => Read(package, [page]));
        Assert.Equal(original, text.SourceXml);
    }

    [Fact]
    public void TemplatePathMetadataIsBoundedBeforeResolvingSegments()
    {
        var (package, page) = Package(Text("body"));
        var template = new OfdTemplateContent { BaseLocation = string.Concat(Enumerable.Repeat("x/", 100)) + "Content.xml" };
        template.Elements.Add(Text("template")); page.Templates.Add(template);
        var failure = Assert.Throws<InvalidDataException>(() => Read(package, [page], new()
        {
            IncludeTemplates = true, Limits = new() { MaxReferenceCount = 20 }
        }));
        Assert.Contains("reference budget", failure.Message);
    }

    [Fact]
    public void ResolvedTemplatePathsChargeRepeatedLongDirectoryAndCacheSharedProvenance()
    {
        var (package, page) = Package(Text("body"));
        package.DocumentEntryPath = "Doc_0/" + new string('d', 1000) + "/Document.xml";
        for (var i = 0; i < 3; i++)
        {
            var template = new OfdTemplateContent { BaseLocation = $"T{i}.xml" };
            template.Elements.Add(Text("template")); page.Templates.Add(template);
        }
        Assert.Throws<InvalidDataException>(() => Read(package, [page], new()
        {
            IncludeTemplates = true, Limits = new() { MaxSourceXmlBytes = 3500 }
        }));
        foreach (var template in page.Templates) template.BaseLocation = "shared.xml";
        var shared = Read(package, [page], new() { IncludeTemplates = true, Limits = new() { MaxSourceXmlBytes = 3500 } });
        var occurrences = shared.Occurrences.Where(o => o.RootScope == OfdContentScope.Template).ToList();
        Assert.Equal(3, occurrences.Count);
        Assert.Same(occurrences[0].SourceEntryPath, occurrences[1].SourceEntryPath);
        Assert.Same(occurrences[1].SourceEntryPath, occurrences[2].SourceEntryPath);
        foreach (var template in page.Templates) { template.Elements.Clear(); template.BaseLocation = new string('x', 600); }
        Assert.Single(Read(package, [page], new() { IncludeTemplates = true, Limits = new() { MaxSourceXmlBytes = 3500 } }).Occurrences);
    }

    [Fact]
    public void OptionsAndLimitsAreCapturedBeforeSelectionEnumeration()
    {
        var (package, page) = Package(new OfdRawElement { Xml = "<PageBlock/>" });
        var options = new OfdTraversalOptions();
        var selected = new CallbackPages(page, () => { options.RequireSupportedSource = true; options.Limits.MaxExpandedOccurrenceCount = 0; });
        var result = OfdModelTraversal.ReadGraphics(package, selected, options);
        Assert.Single(result.Diagnostics);
    }

    [Fact]
    public void CancellationBeforeAndDuringSelectionProducesNoResultAndDoesNotMutateModel()
    {
        var (package, page) = Package(Text("unchanged"));
        using var before = new CancellationTokenSource(); before.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfdModelTraversal.ReadGraphics(package, [page], new(), before.Token));
        using var during = new CancellationTokenSource();
        Assert.Throws<OperationCanceledException>(() => OfdModelTraversal.ReadGraphics(package,
            new CallbackPages(page, during.Cancel), new(), during.Token));
        Assert.Equal("unchanged", ((OfdTextElement)Assert.Single(page.Elements)).Text);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(-1)]
    public void NonpositiveLimitsAreRejected(int value)
    {
        var (package, page) = Package(Text("x"));
        Assert.Throws<ArgumentException>(() => Read(package, [page], new() { Limits = new() { MaxNodeCount = value } }));
        Assert.Throws<ArgumentException>(() => Read(package, [page], new() { Limits = new() { MaxDepth = value } }));
        Assert.Throws<ArgumentException>(() => Read(package, [page], new() { Limits = new() { MaxReferenceCount = value } }));
        Assert.Throws<ArgumentException>(() => Read(package, [page], new() { Limits = new() { MaxExpandedOccurrenceCount = value } }));
        Assert.Throws<ArgumentException>(() => Read(package, [page], new() { Limits = new() { MaxSourceXmlBytes = value } }));
    }

    private static OfdTextElement Text(string value) => new() { Text = value };
    private static string XmlText(string value) => $"<TextObject xmlns='http://www.ofdspec.org' Font='1' Size='4'><TextCode X='0' Y='4'>{value}</TextCode></TextObject>";
    private static (OfdDocumentPackage Package, OfdPage Page) Package(params OfdElement[] elements)
    {
        var package = new OfdDocumentPackage(); var page = new OfdPage();
        page.Elements.AddRange(elements); package.Pages.Add(page); return (package, page);
    }
    private static OfdGraphicTraversalResult Read(OfdDocumentPackage package, IReadOnlyList<OfdPage> pages, OfdTraversalOptions? options = null)
        => OfdModelTraversal.ReadGraphics(package, pages, options ?? new());

    private sealed class CallbackPages(OfdPage page, Action callback) : IReadOnlyList<OfdPage>
    {
        public int Count => 1;
        public OfdPage this[int index] => index == 0 ? page : throw new IndexOutOfRangeException();
        public IEnumerator<OfdPage> GetEnumerator() { callback(); yield return page; }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
