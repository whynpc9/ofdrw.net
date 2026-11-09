using System.Buffers.Binary;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.TestSupport;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class Upstream202611ArchiveTests
{
    [Theory]
    [InlineData("plain")]
    [InlineData("adjusted")]
    public async Task ReadersResolveMixedCaseBranches(string layout)
    {
        var bytes = Upstream202611ArchiveFixture.Create(layout);
        using var seekable = new MemoryStream(bytes, writable: false);
        using var nonSeekable = new NonSeekableInput(bytes);
        var reader = new OfdReader();
        AssertResolved(await reader.ReadAsync(seekable, new OfdPackageLoadOptions()));
        AssertResolved(await reader.ReadAsync(nonSeekable, new OfdPackageLoadOptions()));

        using var ownedSpool = new MemoryStream(bytes, writable: false);
        using var context = new OfdControlledEditContext(new());
        AssertResolved(await reader.ReadControlledAsync(ownedSpool, new(), context));
        Assert.Equal(1, context.Material.ReadMetadataReservations);
        Assert.Equal(8, context.Material.ReadEntryAllocations);
    }

    [Fact]
    public async Task PlainZipAtNonzeroPositionIsBufferedOnlyByPublicReader()
    {
        var bytes = Upstream202611ArchiveFixture.Create("unadjusted");
        using var input = new MemoryStream(bytes, writable: false) { Position = 77 };
        AssertResolved(await new OfdReader().ReadAsync(input, new OfdPackageLoadOptions()));

        input.Position = 77;
        await AssertUnsupportedSpool(input);
        using var nonSeekable = new NonSeekableInput(Upstream202611ArchiveFixture.Create());
        await AssertUnsupportedSpool(nonSeekable);
    }

    [Fact]
    public async Task PrefixLayoutsHaveDifferentPhysicalZipOffsets()
    {
        var plain = Upstream202611ArchiveFixture.Create();
        var adjusted = Upstream202611ArchiveFixture.Create("adjusted");
        var unadjusted = Upstream202611ArchiveFixture.Create("unadjusted");
        Assert.Equal(plain.Length + 77, adjusted.Length);
        Assert.Equal(adjusted.Length, unadjusted.Length);
        Assert.Equal("504B0506" + new string('0', 36), Convert.ToHexString(adjusted.AsSpan(16, 22)));
        Assert.Equal(adjusted.AsSpan(0, 77).ToArray(), unadjusted.AsSpan(0, 77).ToArray());

        var plainCentral = Offset(plain, plain.Length - 6);
        var adjustedCentral = Offset(adjusted, adjusted.Length - 6);
        Assert.Equal(plainCentral + 77, adjustedCentral);
        Assert.Equal(plainCentral, Offset(unadjusted, unadjusted.Length - 6));
        Assert.Equal(0x02014b50, Offset(adjusted, adjustedCentral));
        Assert.Equal(0, Offset(plain, plainCentral + 42));
        Assert.Equal(77, Offset(adjusted, adjustedCentral + 42));
        Assert.Equal(0, Offset(unadjusted, adjustedCentral + 42));
        Assert.Equal(0x04034b50, Offset(adjusted, 77));

        using var seekable = new MemoryStream(unadjusted, writable: false);
        using var nonSeekable = new NonSeekableInput(unadjusted);
        var reader = new OfdReader();
        await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(seekable, new OfdPackageLoadOptions()));
        await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(nonSeekable, new OfdPackageLoadOptions()));

        using var ownedSpool = new MemoryStream(unadjusted, writable: false);
        using var context = new OfdControlledEditContext(new());
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadControlledAsync(ownedSpool, new(), context));
        Assert.Contains("central directory range/count", error.Message);
        Assert.Equal(0, context.Material.ReadMetadataReservations);
        Assert.Equal(0, context.Material.ReadEntryAllocations);
    }

    [Theory]
    [InlineData("collision", "duplicate entry")]
    [InlineData("traversal", "unsafe path")]
    [InlineData("truncated", null)]
    [InlineData("false-local-header", null)]
    public async Task ReadersRejectConflictingOrDamagedArchives(string defect, string? reason)
    {
        var bytes = Upstream202611ArchiveFixture.Create(defect: defect);
        using var seekable = new MemoryStream(bytes, writable: false);
        using var nonSeekable = new NonSeekableInput(bytes);
        using var ownedSpool = new MemoryStream(bytes, writable: false);
        using var context = new OfdControlledEditContext(new());
        var reader = new OfdReader();
        var publicError = await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(seekable, new OfdPackageLoadOptions()));
        await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(nonSeekable, new OfdPackageLoadOptions()));
        if (defect == "truncated")
        {
            var error = await Assert.ThrowsAsync<NotSupportedException>(() => reader.ReadControlledAsync(ownedSpool, new(), context));
            Assert.Contains("final EOCD", error.Message);
            Assert.Equal(0, context.Material.ReadMetadataReservations);
        }
        else
        {
            var error = await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadControlledAsync(ownedSpool, new(), context));
            if (reason is not null)
            {
                Assert.Contains(reason, publicError.Message);
                Assert.Contains(reason, error.Message);
            }
        }
        Assert.Equal(0, context.Material.ReadDomAllocations);
    }

    [Fact]
    public async Task ReadersRejectMissingMandatoryPageWithoutFallingBackToTemplate()
    {
        var bytes = Upstream202611ArchiveFixture.Create(defect: "missing-page");
        using var seekable = new MemoryStream(bytes, writable: false);
        using var nonSeekable = new NonSeekableInput(bytes);
        using var ownedSpool = new MemoryStream(bytes, writable: false);
        using var context = new OfdControlledEditContext(new());
        var reader = new OfdReader();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(seekable, new OfdPackageLoadOptions()));
        Assert.Contains("a/content.xml", Assert.IsType<KeyNotFoundException>(error.InnerException).Message);
        await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(nonSeekable, new OfdPackageLoadOptions()));
        var controlledError = await Assert.ThrowsAsync<KeyNotFoundException>(() => reader.ReadControlledAsync(ownedSpool, new(), context));
        Assert.Contains("a/content.xml", controlledError.Message);
    }

    [Fact]
    public async Task AdjustedPrefixCountsTowardInputBudgetAndHonorsCancellation()
    {
        var bytes = Upstream202611ArchiveFixture.Create("adjusted");
        var options = new OfdPackageLoadOptions { MaxInputBytes = bytes.Length - 77 };
        using var seekable = new MemoryStream(bytes, writable: false);
        using var nonSeekable = new NonSeekableInput(bytes);
        using var ownedSpool = new MemoryStream(bytes, writable: false);
        using var context = new OfdControlledEditContext(new());
        var reader = new OfdReader();
        var error = await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(seekable, options));
        Assert.Contains("compressed input size limit", error.Message);
        await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadAsync(nonSeekable, options));
        await Assert.ThrowsAsync<InvalidDataException>(() => reader.ReadControlledAsync(ownedSpool, options, context));
        Assert.Equal(0, context.Material.ReadMetadataReservations);
        Assert.Equal(0, context.Material.ReadEntryAllocations);

        var canceled = new CancellationToken(canceled: true);
        seekable.Position = 0;
        ownedSpool.Position = 0;
        using var canceledContext = new OfdControlledEditContext(new());
        using var canceledNonSeekable = new NonSeekableInput(bytes);
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadAsync(seekable, new OfdPackageLoadOptions(), canceled));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadAsync(canceledNonSeekable, new OfdPackageLoadOptions(), canceled));
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() => reader.ReadControlledAsync(ownedSpool, new(), canceledContext, canceled));
        Assert.Equal(0, canceledContext.Material.ReadMetadataReservations);
        Assert.Equal(0, canceledContext.Material.ReadEntryAllocations);
    }

    private static async Task AssertUnsupportedSpool(Stream spool)
    {
        using var context = new OfdControlledEditContext(new());
        var error = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdReader().ReadControlledAsync(spool, new(), context));
        Assert.Contains("owned readable seekable spool at position zero", error.Message);
        Assert.Equal(0, context.Material.ReadMetadataReservations);
    }

    private static int Offset(byte[] bytes, int index) => BinaryPrimitives.ReadInt32LittleEndian(bytes.AsSpan(index, 4));

    private static void AssertResolved(OfdDocumentPackage package)
    {
        var page = Assert.Single(package.Pages);
        Assert.Equal("1", page.Id);
        Assert.Equal(100, page.WidthMillimeters);
        Assert.Equal(70, page.HeightMillimeters);
        var image = Assert.IsType<OfdImageElement>(Assert.Single(page.Elements));
        Assert.Equal(10, image.XMillimeters);
        Assert.Equal(Upstream202611ArchiveFixture.Blue, image.Data);
        var template = Assert.Single(page.Templates);
        Assert.Equal("3", template.TemplateId);
        Assert.Equal("Background", template.ZOrder);
        var templateImage = Assert.IsType<OfdImageElement>(Assert.Single(template.Elements));
        Assert.Equal(60, templateImage.XMillimeters);
        Assert.Equal(Upstream202611ArchiveFixture.Red, templateImage.Data);
        Assert.Equal(Upstream202611ArchiveFixture.Red, package.PreservedEntries["Doc_0/A/swatch.png"]);
        Assert.Equal(Upstream202611ArchiveFixture.Blue, package.PreservedEntries["DOC_0/B/swatch.png"]);
    }
}
