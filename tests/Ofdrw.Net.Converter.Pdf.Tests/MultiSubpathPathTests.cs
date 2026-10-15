using System.IO;
using System.Linq;
using Ofdrw.Net.Converter.Pdf.Internal;
using Ofdrw.Net.Core.Models;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;

namespace Ofdrw.Net.Converter.Pdf.Tests;

public sealed class MultiSubpathPathTests
{
    [Theory]
    [InlineData("M 0 0 L 170 0 L 170 30 L 0 30 C M 0 10 L 170 10 M 85 0 L 85 30", 3)]
    [InlineData("M 0 0 L 10 0 C m 20 20 l 10 0 m 0 10 B 1 2 3 4 5 6", 3)]
    public void MoveStartsAnIndependentSubpath(string literal, int moves)
    {
        using var document = new PdfDocument();
        var page = document.AddPage();
        using (var graphics = XGraphics.FromPdfPage(page))
            Assert.True(OfdPathRenderer.TryDraw(graphics, new OfdPathElement
            {
                AbbreviatedData = literal, Fill = true,
                Transform = new double[] { 1, 0.2, 0.1, 1, 4, 5 },
                XMillimeters = 10, YMillimeters = 20
            }, 2, 3));
        using var bytes = new MemoryStream();
        document.Save(bytes, false);
        using var parsed = UglyToad.PdfPig.PdfDocument.Open(bytes.ToArray());
        var operations = parsed.GetPage(1).Operations;
        Assert.Equal(moves, operations.Count(operation => operation.Operator == "m"));
        Assert.Single(operations, operation => operation.Operator is "B" or "B*");
    }
}
