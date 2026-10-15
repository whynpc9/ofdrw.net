using Ofdrw.Net.Core.Models;

namespace Ofdrw.Net.Core.Tests;

public sealed class WriterNumericTests
{
    [Fact]
    public void ScalarWriterContractUsesFormattedValuesAtActualFloatBoundary()
    {
        // .0005f lies above the double boundary; its adjacent float lies below.
        // Keep the existing writer format rather than substitute Math.Round.
        var half = .0005f;
        Assert.Equal(0, OfdNumericFormat.WriterValue((double)float.BitDecrement(half)));
        Assert.False(OfdNumericFormat.WriterPositive(float.BitDecrement(half)));
        Assert.Equal(.001, OfdNumericFormat.WriterValue(half));
        Assert.True(OfdNumericFormat.WriterPositive(half));
        Assert.Equal(.001, OfdNumericFormat.WriterValue(double.BitDecrement(.0005)));
        Assert.True(OfdNumericFormat.WriterPositive(double.BitDecrement(.0005)));
        foreach (var value in new[] { double.NaN, double.PositiveInfinity, double.NegativeInfinity, -1d, 0d, double.Epsilon, .0004 })
            Assert.False(OfdNumericFormat.WriterPositive(value));
        // The legacy fixed format can round the largest finite double beyond
        // its range. Preserve 04's rejection of a nonfinite written result.
        Assert.False(OfdNumericFormat.WriterPositive(double.MaxValue));
        Assert.True(double.IsPositiveInfinity(OfdNumericFormat.WriterValue(double.MaxValue)));
        var prior = System.Globalization.CultureInfo.CurrentCulture;
        try
        {
            System.Globalization.CultureInfo.CurrentCulture = System.Globalization.CultureInfo.GetCultureInfo("fr-FR");
            Assert.Equal(.001, OfdNumericFormat.WriterValue(half));
            Assert.True(OfdNumericFormat.WriterPositive(half));
        }
        finally { System.Globalization.CultureInfo.CurrentCulture = prior; }
    }
}
