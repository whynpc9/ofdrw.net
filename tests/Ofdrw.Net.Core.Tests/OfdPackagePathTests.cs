using Ofdrw.Net.Core.IO;

namespace Ofdrw.Net.Core.Tests;

public sealed class OfdPackagePathTests
{
    [Theory]
    [InlineData("Doc_0/Signs/Signatures.xml", "Sign_0/Signature.xml", "Doc_0/Signs/Sign_0/Signature.xml")]
    [InlineData("Doc_0/Signs/Signatures.xml", "./Sign_0/Signature.xml", "Doc_0/Signs/Sign_0/Signature.xml")]
    [InlineData("Doc_0/Signs/Signatures.xml", "../Res/a.png", "Doc_0/Res/a.png")]
    [InlineData("Doc_0/Signs/Signatures.xml", "/Doc_0/./Pages//Page_0/Content.xml", "Doc_0/Pages/Page_0/Content.xml")]
    [InlineData("Doc_0/Signs/Signatures.xml", @"Sign_0\Signature.xml", "Doc_0/Signs/Sign_0/Signature.xml")]
    [InlineData("OFD.xml", "Doc_0/Document.xml", "Doc_0/Document.xml")]
    [InlineData("", "content.txt", "content.txt")]
    public void ResolvesPackageReferences(string containingEntry, string reference, string expected)
    {
        Assert.True(OfdPackagePath.TryResolve(containingEntry, reference, out var resolved));
        Assert.Equal(expected, resolved);
        Assert.Equal(expected, OfdPackagePath.Resolve(containingEntry, reference));
    }

    [Theory]
    [InlineData("OFD.xml", "../x")]
    [InlineData("Doc_0/a.xml", "../../x")]
    [InlineData("Doc_0/a.xml", "/../x")]
    public void RejectsReferencesThatEscapeThePackage(string containingEntry, string reference)
    {
        Assert.False(OfdPackagePath.TryResolve(containingEntry, reference, out var resolved));
        Assert.Empty(resolved);
        var error = Assert.Throws<InvalidDataException>(() => OfdPackagePath.Resolve(containingEntry, reference));
        Assert.Equal("Package reference escapes the OFD root.", error.Message);
    }
}
