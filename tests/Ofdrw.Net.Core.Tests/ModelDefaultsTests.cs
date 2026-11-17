using Ofdrw.Net.Core.Constants;
using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Tests;

public sealed class ModelDefaultsTests
{
    [Fact]
    public void OfdDocumentOptions_HasExpectedDefaults()
    {
        var options = new OfdDocumentOptions();

        Assert.Equal(OfdConstants.DefaultDocType, options.DocType);
        Assert.Equal(OfdConstants.DefaultDocId, options.DocumentId);
        Assert.Equal(OfdConstants.Namespace, options.Namespace);
        Assert.True(options.EnableDeflateCompression);
    }

    [Theory]
    [InlineData("OFD-A", true)]
    [InlineData("OFD", false)]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("ofd-a", false)]
    [InlineData(" OFD-A", false)]
    [InlineData("OFD-A ", false)]
    [InlineData("OFD-H", false)]
    public void OfdDocTypes_IsArchival_IsExactOrdinal(string? declared, bool expected)
    {
        Assert.Equal(expected, OfdDocTypes.IsArchival(declared));
        // The default declaration of a new document is ordinary, not archival.
        Assert.False(OfdDocTypes.IsArchival(new OfdDocumentOptions().DocType));
    }

    [Fact]
    public void OfdArchivalOriginalException_NamesOperationAndDerivationPath()
    {
        var error = new OfdArchivalOriginalException("Example.Operation");
        Assert.IsAssignableFrom<InvalidOperationException>(error);
        Assert.Equal("Example.Operation", error.Operation);
        Assert.Equal("OFD-A", error.DeclaredDocType);
        Assert.StartsWith("Example.Operation ", error.Message);
        Assert.Contains("DocType=\"OFD-A\"", error.Message);
        Assert.Contains("OfdOrdinaryDerivation.CreateAsync", error.Message);
        Assert.Contains("ofdrw derive-ofd", error.Message);
    }

    [Fact]
    public void OfdDocumentPackage_CanStorePagesAndAttachments()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.Attachments.Add(new OfdAttachment { Name = "a.txt", Data = [1, 2, 3] });

        Assert.Single(package.Pages);
        Assert.Single(package.Attachments);
    }
}
