using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class TypedClipCompositeTests
{
    internal static OfdDocumentPackage Create()
    {
        var package = new OfdDocumentPackage { CompositeResources = new() };
        var resource = new OfdCompositeResource { WidthMillimeters = 40, HeightMillimeters = 20 };
        resource.Elements.Add(new OfdTextElement { Text = "shared text", WidthMillimeters = 35, HeightMillimeters = 10 });
        resource.Elements.Add(new OfdPathElement { AbbreviatedData = "M 0 0 L 30 0 L 30 15 L 0 15 C", WidthMillimeters = 30, HeightMillimeters = 15, Fill = true });
        package.CompositeResources.Items.Add(resource);
        var page = new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdCompositeElement { Resource = resource, XMillimeters = 10, YMillimeters = 20, WidthMillimeters = 40, HeightMillimeters = 20 });
        page.Elements.Add(new OfdCompositeElement { Resource = resource, XMillimeters = 55, YMillimeters = 50, WidthMillimeters = 40, HeightMillimeters = 20, Transform = new double[] { 1, 0.1, 0, 1, 0, 0 } });
        var path = new OfdPathElement { AbbreviatedData = "M 0 0 L 40 0 L 40 20 L 0 20 C", WidthMillimeters = 40, HeightMillimeters = 20, Clips = new() };
        var clip = new OfdClip(); clip.Areas.Add(new OfdClipArea { Transform = new double[] { 1, 0, 0, 1, 2, 3 },
            Path = new OfdPathElement { Fill = true, Stroke = false, FillRule = OfdFillRule.EvenOdd, AbbreviatedData = "M 0 0 L 20 0 L 20 20 L 0 20 C" } });
        path.Clips.Items.Add(clip); page.Elements.Add(path); package.Pages.Add(page); return package;
    }

    [Fact]
    public async Task PublicConstructionAndEditPreserveSharedDefinitionsAndClips()
    {
        using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), bytes); bytes.Position = 0;
        var loaded = await new OfdReader().ReadAsync(bytes);
        var resource = Assert.Single(loaded.CompositeResources!.Items);
        Assert.All(loaded.Pages[0].Elements.OfType<OfdCompositeElement>(), placement => Assert.Same(resource, placement.Resource));
        var path = Assert.Single(loaded.Pages[0].Elements.OfType<OfdPathElement>());
        Assert.Equal<OfdFillRule?>(OfdFillRule.EvenOdd, path.Clips!.Items[0].Areas[0].Path.FillRule);
        path.Clips.Items[0].Areas[0].Transform![4] = 7;
        ((OfdCompositeElement)loaded.Pages[0].Elements[0]).Transform = new double[] { 1, 0, 0, 1, 3, 4 };
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(loaded, output); output.Position = 0;
        var readback = await new OfdReader().ReadAsync(output);
        Assert.Equal(7, Assert.Single(readback.Pages[0].Elements.OfType<OfdPathElement>()).Clips!.Items[0].Areas[0].Transform![4]);
        Assert.Single(readback.CompositeResources!.Items);
        Assert.Equal(3, ((OfdCompositeElement)readback.Pages[0].Elements[0]).Transform![4]);
    }

    [Fact]
    public async Task CyclesAndMixedClipAuthorityFailBeforeOutput()
    {
        var package = Create(); var resource = package.CompositeResources!.Items[0];
        resource.Elements.Add(new OfdCompositeElement { Resource = resource, WidthMillimeters = 1, HeightMillimeters = 1 });
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
        package = Create(); package.Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml = "<Clips/>";
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task OwnedClonePreservesDagAndIndependentClipMatrices()
    {
        using var bytes = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), bytes); bytes.Position = 0;
        var source = await new OfdReader().ReadAsync(bytes);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var original = source.CompositeResources!.Items[0]; var copied = candidate.Package.CompositeResources!.Items[0];
        Assert.NotSame(original, copied);
        Assert.All(candidate.Package.Pages[0].Elements.OfType<OfdCompositeElement>(), placement => Assert.Same(copied, placement.Resource));
        var path = candidate.Package.Pages[0].Elements.OfType<OfdPathElement>().Single();
        path.Clips!.Items[0].Areas[0].Transform![4] = 11;
        Assert.Equal(2, source.Pages[0].Elements.OfType<OfdPathElement>().Single().Clips!.Items[0].Areas[0].Transform![4]);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(candidate); using var output = new MemoryStream(); await prepared.WriteAsync(output);
        output.Position = 0; Assert.Equal(11, (await new OfdReader().ReadAsync(output)).Pages[0].Elements.OfType<OfdPathElement>().Single().Clips!.Items[0].Areas[0].Transform![4]);
    }

    [Fact]
    public void StrictTraversalRefusesAnEmptyCompositeAndFreshTypedClip()
    {
        var package = Create(); package.CompositeResources!.Items[0].Elements.Clear();
        var result = OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions());
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Kind == OfdTraversalDiagnosticKind.UnsupportedSource);
        Assert.Throws<System.NotSupportedException>(() => OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions { RequireSupportedSource = true }));
        Assert.Throws<System.NotSupportedException>(() => new Ofdrw.Net.Reader.Keywords.OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(package, "missing"));
        package.Pages[0].Elements.RemoveAll(e => e is OfdCompositeElement);
        Assert.Throws<System.NotSupportedException>(() => OfdModelTraversal.ReadGraphics(package, package.Pages, new OfdTraversalOptions { RequireSupportedSource = true }));
    }

    [Fact]
    public void PublicStrictQueryRejectsFreshTypedTextClip()
    {
        var package = KeywordLocatorTests.Model("A");
        var text = package.Pages[0].Elements.OfType<OfdTextElement>().Single(); text.Clips = new();
        Assert.Null(text.SourceXml);
        Assert.Throws<System.NotSupportedException>(() => new Ofdrw.Net.Reader.Keywords.OfdKeywordLocator(new KeywordLocatorTests.Provider()).Find(package, "missing"));
    }

    [Fact]
    public async Task WithdrawingDirtyTypedOwnersAndLateMutationReject()
    {
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), input); input.Position = 0;
        var package = await new OfdReader().ReadAsync(input);
        package.CompositeResources!.Items[0].WidthMillimeters += 1;
        Assert.Throws<System.InvalidOperationException>(() => package.CompositeResources = null);
        var path = package.Pages[0].Elements.OfType<OfdPathElement>().Single();
        path.Clips!.Items[0].Areas[0].Transform![0] = 2;
        Assert.Throws<System.InvalidOperationException>(() => path.Clips = null);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var candidate = OfdOwnedEditCandidate.CloneForControlledEdit(package, context.CaptureSource(package), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(candidate);
        package.CompositeResources.Items[0].Elements.OfType<OfdPathElement>().Single().FillRule = OfdFillRule.EvenOdd;
        using var output = new MemoryStream(); await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task SplitKeepsSharedClosureAndLastDeletedUsePrunesDefinition()
    {
        var package = Create(); var page = new OfdPage { Index = 1, WidthMillimeters = 100, HeightMillimeters = 100 };
        page.Elements.Add(new OfdTextElement { Text = "retained", WidthMillimeters = 40, HeightMillimeters = 10 }); package.Pages.Add(page);
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, input); input.Position = 0;
        package = await new OfdReader().ReadAsync(input);
        var first = Ofdrw.Net.Layout.Editing.OfdDocumentSplitter.Split(package, new[] { 0 });
        Assert.Single(first.CompositeResources!.Items);
        Assert.Same(first.CompositeResources.Items[0], ((OfdCompositeElement)first.Pages[0].Elements[0]).Resource);
        var second = Ofdrw.Net.Layout.Editing.OfdDocumentSplitter.Split(package, new[] { 1 });
        using var output = new MemoryStream(); await new OfdPackageWriter().WriteAsync(second, output); output.Position = 0;
        Assert.Null((await new OfdReader().ReadAsync(output)).CompositeResources);
    }

    [Fact]
    public async Task UnknownCompositeAndClipXmlSurvivesRoundtrip()
    {
        using var input = new MemoryStream(); await new OfdPackageWriter().WriteAsync(Create(), input);
        var entries = new System.Collections.Generic.Dictionary<string, byte[]>();
        input.Position = 0;
        using (var zip = new System.IO.Compression.ZipArchive(input, System.IO.Compression.ZipArchiveMode.Read, true))
            foreach (var entry in zip.Entries) { using var data = entry.Open(); using var bytes = new MemoryStream(); data.CopyTo(bytes); entries.Add(entry.FullName, bytes.ToArray()); }
        foreach (var path in entries.Keys.ToArray())
        {
            if (!path.EndsWith(".xml")) continue;
            var xml = System.Xml.Linq.XDocument.Load(new MemoryStream(entries[path]));
            foreach (var node in xml.Descendants().Where(n => n.Name.LocalName == "CompositeGraphicUnit" || n.Name.LocalName == "Area" && n.Parent?.Name.LocalName == "Clip"))
                node.Add(new System.Xml.Linq.XElement(System.Xml.Linq.XName.Get("Keep", "urn:vendor"), new System.Xml.Linq.XAttribute("ID", "opaque"), new System.Xml.Linq.XCData("preserved")));
            entries[path] = System.Text.Encoding.UTF8.GetBytes(xml.ToString());
        }
        using var modified = new MemoryStream();
        using (var zip = new System.IO.Compression.ZipArchive(modified, System.IO.Compression.ZipArchiveMode.Create, true))
            foreach (var entry in entries) { using var target = zip.CreateEntry(entry.Key).Open(); target.Write(entry.Value); }
        modified.Position = 0; var loaded = await new OfdReader().ReadAsync(modified);
        using var saved = new MemoryStream(); await new OfdPackageWriter().WriteAsync(loaded, saved); saved.Position = 0;
        var readback = await new OfdReader().ReadAsync(saved);
        Assert.Contains("preserved", readback.CompositeResources!.Items[0].SourceXml);
        Assert.Contains("preserved", readback.Pages[0].Elements.OfType<OfdPathElement>().Single().ClippingXml);
    }

    [Fact]
    public void HiddenUnusedDefinitionsStillConsumeDepthAndCancellationBudget()
    {
        var package = Create(); var child = package.CompositeResources!.Items[0];
        for (var i = 0; i < 6; i++)
        {
            var parent = new OfdCompositeResource { WidthMillimeters = 1, HeightMillimeters = 1 };
            parent.Elements.Add(new OfdCompositeElement { Resource = child, WidthMillimeters = 1, HeightMillimeters = 1 });
            package.CompositeResources.Items.Add(parent); child = parent;
        }
        Assert.Throws<InvalidDataException>(() => OfdCompositeGraph.Validate(package, new OfdModelMaterialization(new OfdModelLimits { MaxDepth = 3 }), default));
        using var cancellation = new System.Threading.CancellationTokenSource(); cancellation.Cancel();
        Assert.ThrowsAny<System.OperationCanceledException>(() => OfdCompositeGraph.Validate(package, new OfdModelMaterialization(new OfdModelLimits()), cancellation.Token));
    }
}
