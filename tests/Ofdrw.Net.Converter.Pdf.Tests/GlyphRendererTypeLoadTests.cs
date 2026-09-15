using Ofdrw.Net.Converter.Pdf.Converters;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class GlyphRendererTypeLoadTests
{
    [Fact]
    public void ConverterPdfAssembly_ShouldLoadPdfGlyphOutlineRenderer()
    {
        var types = typeof(OfdToPdfConverter).Assembly.GetTypes();
        Assert.Contains(types, type => type.FullName == "Ofdrw.Net.Converter.Pdf.Internal.PdfGlyphOutlineRenderer");
    }
}
