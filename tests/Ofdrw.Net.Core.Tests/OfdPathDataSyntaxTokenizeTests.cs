using Ofdrw.Net.Core.Models;
using Xunit;

namespace Ofdrw.Net.Core.Tests;

/// <summary>PDF and SVG interpret AbbreviatedData through one tokenizer next to the admission grammar.</summary>
public sealed class OfdPathDataSyntaxTokenizeTests
{
    [Fact]
    public void SplitsCommandsAndNumbersAndDropsSeparators() =>
        Assert.Equal(new[] { "M", "0", "-1.5", "L", "2e3", ".5", "B", "1", "2", "3", "4", "5", "6", "C" },
            OfdPathDataSyntax.Tokenize("M0-1.5 L 2e3,.5B1 2 3 4 5 6 C"));
}
