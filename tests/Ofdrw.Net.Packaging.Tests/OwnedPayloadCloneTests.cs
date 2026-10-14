using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class OwnedPayloadCloneTests
{
    [Fact]
    public void WholeClonePreservesAliasesAcrossEveryExistingPayloadRoleWithoutBorrowingSource()
    {
        var bytes = new byte[] { 1, 2, 3, 4 };
        var source = Package(bytes);
        var page = Assert.Single(source.Pages);
        var body = Assert.IsType<OfdImageElement>(Assert.Single(page.Elements));
        var template = new OfdTemplateContent { TemplateId = "template" };
        var templateImage = Image(bytes); template.Elements.Add(templateImage); page.Templates.Add(template);
        var annotation = Image(bytes); page.AnnotationAppearances.Add(annotation);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        var cloneBytes = Assert.Single(owned.Package.Fonts).Data;
        Assert.NotSame(bytes, cloneBytes);
        Assert.Same(cloneBytes, owned.Package.PreservedEntries["opaque.bin"]);
        Assert.Same(cloneBytes, Assert.Single(owned.Package.Attachments).Data);
        Assert.Same(cloneBytes, owned.GetMapped(body).Data);
        Assert.Same(cloneBytes, owned.GetMapped(templateImage).Data);
        Assert.Same(cloneBytes, owned.GetMapped(annotation).Data);
        Assert.Equal(1, context.Material.OwnedPayloadAllocations);
        var storage = context.Material.StorageBytes; context.Material.StoreArray(cloneBytes);
        Assert.Equal(storage, context.Material.StorageBytes);
        cloneBytes[0] = 99;
        Assert.Equal(99, owned.GetMapped(body).Data[0]); Assert.Equal(1, bytes[0]);
        bytes[1] = 88; Assert.Equal(2, cloneBytes[1]);
    }

    [Fact]
    public void EqualDistinctArraysAndDistinctEmptyArraysRemainIndependent()
    {
        var source = Package(new byte[] { 1, 2 });
        source.Fonts.Add(new OfdFontResource { Id = "second", Data = new byte[] { 1, 2 } });
        var emptyOne = new byte[0]; var emptyTwo = new byte[0];
        source.PreservedEntries.Add("empty-one.bin", emptyOne);
        source.PreservedEntries.Add("empty-one-alias.bin", emptyOne);
        source.PreservedEntries.Add("empty-two.bin", emptyTwo);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        Assert.NotSame(owned.Package.Fonts[0].Data, owned.Package.Fonts[1].Data);
        Assert.Equal(owned.Package.Fonts[0].Data, owned.Package.Fonts[1].Data);
        var one = owned.Package.PreservedEntries["empty-one.bin"];
        var two = owned.Package.PreservedEntries["empty-two.bin"];
        Assert.NotSame(emptyOne, one); Assert.NotSame(emptyTwo, two); Assert.NotSame(one, two);
        Assert.Same(one, owned.Package.PreservedEntries["empty-one-alias.bin"]);
        Assert.Equal(4, context.Material.OwnedPayloadAllocations);
    }

    [Fact]
    public async Task DetachedOriginalPageSharesOnlyTheCandidatePayloadCopy()
    {
        var source = new OfdDocumentPackage(); var bytes = new byte[] { 1, 2, 3 };
        var first = new OfdPage { Index = 0 }; first.Elements.Add(Image(bytes)); source.Pages.Add(first);
        var second = new OfdPage { Index = 1 }; second.Elements.Add(Image(bytes)); source.Pages.Add(second);
        using var original = new MemoryStream(); await new OfdPackageWriter().WriteAsync(source, original); original.Position = 0;
        source = await new OfdReader().ReadAsync(original);
        var liveImage = Assert.IsType<OfdImageElement>(source.Pages[0].Elements[0]);
        var removedPage = source.Pages[1]; var removedImage = Assert.IsType<OfdImageElement>(removedPage.Elements[0]);
        Assert.Same(liveImage.Data, removedImage.Data); source.Pages.RemoveAt(1);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        Assert.DoesNotContain(owned.GetMapped(removedPage), owned.Package.Pages);
        Assert.Same(owned.GetMapped(liveImage).Data, owned.GetMapped(removedImage).Data);
        Assert.NotSame(liveImage.Data, owned.GetMapped(liveImage).Data);
        Assert.Contains(owned.Package.PreservedEntries.Values, value => ReferenceEquals(value, owned.GetMapped(removedImage).Data));
    }

    [Fact]
    public void GenericPageCloneAndProjectionKeepTheirExistingPayloadPolicy()
    {
        var bytes = new byte[] { 1, 2 }; var page = new OfdPage();
        page.Elements.Add(Image(bytes)); page.Elements.Add(Image(bytes));
        var clone = OfdModelCloner.ClonePage(page);
        var first = Assert.IsType<OfdImageElement>(clone.Elements[0]).Data;
        var second = Assert.IsType<OfdImageElement>(clone.Elements[1]).Data;
        Assert.NotSame(bytes, first); Assert.NotSame(first, second);
        var projection = OfdModelCloner.ClonePage(page, clonePayloads: false);
        Assert.All(projection.Elements, element => Assert.Same(bytes, Assert.IsType<OfdImageElement>(element).Data));
    }

    [Theory]
    [InlineData(false)] [InlineData(true)]
    public async Task MutationThroughEitherAliasedOwnerStillRefusesBeforeOutput(bool sourceMutation)
    {
        var source = Package(new byte[] { 1, 2, 3 });
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        using var owned = OfdOwnedEditCandidate.CloneForControlledEdit(source, context.CaptureSource(source), context);
        using var prepared = OfdPackageWriter.PrepareOwnedWrite(owned);
        (sourceMutation ? source : owned.Package).PreservedEntries["opaque.bin"][0] = 99;
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => prepared.WriteAsync(output)); Assert.Equal(0, output.Length);
    }

    [Fact]
    public void RetainedStorageQuotaRejectsBeforePayloadAllocationAndConsumesFactory()
    {
        const int payloadLength = 1024 * 1024; const long cap = 16 * 1024 * 1024;
        var source = Package(new byte[payloadLength]);
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = cap });
        var guard = context.CaptureSource(source);
        context.Material.StoreArray(new byte[checked((int)(cap - context.Material.StorageBytes - payloadLength + 1))]);
        var error = Assert.Throws<InvalidDataException>(() => OfdOwnedEditCandidate.CloneForControlledEdit(source, guard, context));
        Assert.Contains("retained-storage", error.Message); Assert.Equal(0, context.Material.OwnedPayloadAllocations);
        Assert.Throws<InvalidOperationException>(() => context.AssertActive());
    }

    [Fact]
    public void CopyQuotaAdmitsNeitherArrayNorUnfundedDictionaryGrowth()
    {
        const int payloadLength = 64 * 1024;
        using var context = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = payloadLength + 1023 });
        using var session = OfdPayloadCloneSession.Create(context, default);
        Assert.Throws<InvalidDataException>(() => session.Copy(new byte[payloadLength], context.Material, default));
        Assert.Equal(0, context.Material.OwnedPayloadAllocations);
        using var tiny = new OfdControlledEditContext(new OfdModelLimits { MaxMaterializedBytes = 895 });
        using var tinySession = OfdPayloadCloneSession.Create(tiny, default);
        Assert.Throws<InvalidDataException>(() => tinySession.Copy(new byte[0], tiny.Material, default));
        Assert.Equal(0, tiny.Material.OwnedPayloadAllocations);
    }

    [Fact]
    public void PayloadSharingCannotMakeRepeatedGraphicOwnershipValid()
    {
        var source = Package(new byte[] { 1 }); var image = source.Pages[0].Elements[0]; source.Pages[0].Elements.Add(image);
        using var context = new OfdControlledEditContext(new OfdModelLimits());
        Assert.Throws<InvalidDataException>(() => context.CaptureSource(source));
    }

    private static OfdDocumentPackage Package(byte[] bytes)
    {
        var source = new OfdDocumentPackage(); var page = new OfdPage(); page.Elements.Add(Image(bytes)); source.Pages.Add(page);
        source.Fonts.Add(new OfdFontResource { Id = "font", FontName = "fixture", Data = bytes });
        source.Attachments.Add(new OfdAttachment { Id = "attachment", Name = "fixture", Data = bytes });
        source.PreservedEntries.Add("opaque.bin", bytes); return source;
    }
    private static OfdImageElement Image(byte[] bytes) => new() { Data = bytes, WidthMillimeters = 10, HeightMillimeters = 10 };
}
