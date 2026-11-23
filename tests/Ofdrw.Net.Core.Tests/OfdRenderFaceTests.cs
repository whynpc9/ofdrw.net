using System;
using System.IO;
using System.Threading;
using Ofdrw.Net.Core.Fonts;
using Ofdrw.Net.Core.Models;
using Xunit;

namespace Ofdrw.Net.Core.Tests;

/// <summary>The face policy PDF and SVG/HTML share: which bytes a resource supplies and which text may use them.</summary>
public sealed class OfdRenderFaceTests
{
    private static byte[] NotoSans() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "testdata", "NotoSans-Regular.ttf"));

    [Fact]
    public void EmbeddedFaceLoadsWithCoverageAndValidatesCoveredText()
    {
        var face = OfdRenderFace.Load(new OfdFontResource { FontName = "Noto Sans", Data = NotoSans() }, null, new FontParseBudget());
        Assert.False(face.NameOnlyBinding);
        Assert.NotNull(face.Coverage);
        face.ValidateText(new OfdTextElement { Text = "Ofdrw" }, "Noto Sans", new FontParseBudget(), CancellationToken.None);
        // An uncovered scalar is a data error the caller must resolve with a fallback font.
        Assert.Throws<InvalidDataException>(() =>
            face.ValidateText(new OfdTextElement { Text = "中" }, "Noto Sans", new FontParseBudget(), CancellationToken.None));
    }

    [Fact]
    public void NameOnlyResourceWithoutBindingIsRefused() =>
        Assert.Throws<NotSupportedException>(() =>
            OfdRenderFace.Load(new OfdFontResource { FontName = "Unbound" }, null, new FontParseBudget()));

    [Fact]
    public void NameOnlyBindingsAreLimitedToBmpText() =>
        Assert.Throws<NotSupportedException>(() => OfdRenderFace.ValidateText(new OfdTextElement { Text = "\U0001F600" },
            null, nameOnlyBinding: true, "Bound", new FontParseBudget(), CancellationToken.None));

    [Fact]
    public void UnmodeledCoverageRefusesText() =>
        Assert.Throws<NotSupportedException>(() => OfdRenderFace.ValidateText(new OfdTextElement { Text = "A" },
            null, nameOnlyBinding: false, "X", new FontParseBudget(), CancellationToken.None));
}
