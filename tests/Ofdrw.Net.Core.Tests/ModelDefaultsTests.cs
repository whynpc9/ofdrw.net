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
