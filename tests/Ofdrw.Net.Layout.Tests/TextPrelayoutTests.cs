using System;
using System.Buffers.Binary;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Xml.Linq;
using Ofdrw.Net.Layout;
using Ofdrw.Net.Layout.Text;
using Ofdrw.Net.Text.HarfBuzz;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Layout.Tests;

public sealed class TextPrelayoutTests
{
    [Fact]
    public void FixedFaceSnapshotAndNativeShapeUseDesignUnitsAndUtf16Clusters()
    {
        var data = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var face = new TextFontFace("Noto Sans", data);
        data[0] = 0xff;
        var copied = face.GetFaceBytes();
        copied[0] = 0xff;
        Assert.NotEqual(0xff, face.GetFaceBytes()[0]);
        Assert.Equal("b85c38ecea8a7cfb39c24e395a4007474fa5a4fc864f6ee33309eb4948d232d5", face.SourceSha256);
        var shaper = new HarfBuzzTextShaper();
        var result = shaper.Shape(new TextShapeRequest("AV fi", 12, face));
        Assert.Equal(face, result.Request.Face);
        Assert.NotEmpty(result.Glyphs);
        Assert.Contains("libHarfBuzzSharp/14.2.1", result.EngineVersion);
        Assert.True(result.XAdvance > 0);
        Assert.All(result.Glyphs, glyph =>
        {
            Assert.InRange(glyph.ClusterStart, 12, 16);
            Assert.InRange(glyph.ClusterLength, 1, 5);
            Assert.NotEqual(0u, glyph.GlyphId);
        });
        Assert.Contains(result.Glyphs, glyph => glyph.Ink is not null);
    }

    [Fact]
    public void FullRunFeaturesChangeLigatureAndDoNotSplitIntoGraphemeShaping()
    {
        var face = Noto();
        var shaper = new HarfBuzzTextShaper();
        var normal = shaper.Shape(new TextShapeRequest("ffi", 0, face));
        var disabled = shaper.Shape(new TextShapeRequest("ffi", 0, face,
            features: new[] { new TextFeature("liga", 0, 0, 3) }));
        Assert.True(normal.Glyphs.Count < disabled.Glyphs.Count);
        Assert.Equal(3, normal.Glyphs[0].ClusterLength);
        Assert.Equal("ffi", normal.Request.Text);
    }

    [Fact]
    public void DirectShapeReportsTheFullSupplementaryMissingScalar()
    {
        var error = Assert.Throws<TextLayoutException>(() => new HarfBuzzTextShaper().Shape(
            new TextShapeRequest("\ud83d\ude00", 11, Noto())));
        Assert.Equal(TextLayoutError.MissingGlyph, error.Error);
        Assert.Equal(11, error.SourceStart);
        Assert.Equal(2, error.SourceLength);
    }

    [Fact]
    public void SameDisplayNameWithDifferentFaceBytesChangesMeasuredWidth()
    {
        var noto = new TextFontFace("Alias", File.ReadAllBytes(Font("NotoSans-Regular.ttf")));
        var wenkai = new TextFontFace("Alias", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var engine = HarfBuzzTextPrelayout.Create();
        var first = engine.Layout(new TextPrelayoutRequest("AV To", new[] { noto }, 4, 100));
        var second = engine.Layout(new TextPrelayoutRequest("AV To", new[] { wenkai }, 4, 100));
        Assert.NotEqual(first.Face.FaceSha256, second.Face.FaceSha256);
        Assert.NotEqual(first.WidthMillimeters, second.WidthMillimeters);
        Assert.NotEqual(first.Lines[0].ShapedRun.XAdvance, second.Lines[0].ShapedRun.XAdvance);
    }

    [Fact]
    public void RtlOutputOrderDoesNotCorruptAbsoluteSourceClusterIntervals()
    {
        var run = new HarfBuzzTextShaper().Shape(new TextShapeRequest("ABC", 8, Noto(),
            TextDirection.RightToLeft));
        Assert.Equal(new[] { 10, 9, 8 }, run.Glyphs.Select(glyph => glyph.ClusterStart));
        Assert.All(run.Glyphs, glyph => Assert.Equal(1, glyph.ClusterLength));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(2)]
    public void UndefinedTextDirectionFailsBeforeShaping(int value)
    {
        var error = Assert.Throws<TextLayoutException>(() =>
            new TextShapeRequest("ABC", 0, Noto(), (TextDirection)value));
        Assert.Equal(TextLayoutError.InvalidRequest, error.Error);
    }

    [Fact]
    public void ReplacementShaperCannotOmitInteriorOrTrailingSource()
    {
        var request = new TextShapeRequest("ABC", 7, Noto());
        TextShapedGlyph Glyph(int start, int length) => new(1, start, length, 500, 0, 0, 0, null, false);
        foreach (var glyphs in new[]
        {
            new[] { Glyph(7, 1), Glyph(9, 1) },
            new[] { Glyph(7, 2) },
        })
            Assert.Equal(TextLayoutError.InvalidRequest,
                Assert.Throws<TextLayoutException>(() => new TextShapedRun(request, glyphs, "replacement")).Error);
        var shared = new TextShapedRun(request, new[] { Glyph(7, 3), Glyph(7, 3) }, "replacement");
        Assert.Equal(2, shared.Glyphs.Count);

        var supplementary = new TextShapeRequest("\ud83d\ude00", 0, Noto());
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            new TextShapedRun(supplementary, new[] { Glyph(0, 1), Glyph(1, 1) }, "replacement")).Error);
        Assert.Single(new TextShapedRun(supplementary, new[] { Glyph(0, 2) }, "replacement").Glyphs);
    }

    [Fact]
    public void ReplacementShaperCannotReturnGlyphOutsideSelectedFace()
    {
        var face = Noto();
        var request = new TextShapeRequest("A", 0, face);
        TextShapedGlyph Glyph(uint id) => new(id, 0, 1, 500, 0, 0, 0, null, false);
        Assert.Single(new TextShapedRun(request, new[] { Glyph((uint)face.GlyphCount - 1) }, "replacement").Glyphs);
        Assert.Equal(TextLayoutError.InvalidRequest, Assert.Throws<TextLayoutException>(() =>
            new TextShapedRun(request, new[] { Glyph((uint)face.GlyphCount) }, "replacement")).Error);
    }

    [Fact]
    public void FinalLinesExposeHeightBaselineSourceAndAdvanceSeparatelyFromInk()
    {
        var face = Noto();
        var engine = HarfBuzzTextPrelayout.Create();
        var wide = engine.Layout(new TextPrelayoutRequest("AV fi\nCafe\u0301", new[] { face }, 4, 100));
        Assert.Equal(2, wide.Lines.Count);
        Assert.Equal((0, 5, 1), (wide.Lines[0].SourceStart, wide.Lines[0].SourceLength, wide.Lines[0].TerminatorLength));
        Assert.Equal((6, 5, 0), (wide.Lines[1].SourceStart, wide.Lines[1].SourceLength, wide.Lines[1].TerminatorLength));
        Assert.True(wide.Lines[0].BaselineMillimeters > wide.Lines[0].TopMillimeters);
        Assert.Equal(wide.Lines[0].HeightMillimeters, wide.Lines[1].TopMillimeters, 6);
        Assert.NotNull(wide.Lines[0].InkBoundsMillimeters);
        Assert.True(wide.WidthMillimeters > 0);
        Assert.False(wide.Overflows);
        Assert.Same(wide.Lines[0].ShapedRun.Request.Face, wide.Face);
        var narrow = engine.Layout(new TextPrelayoutRequest("AV fi", new[] { face }, 4, 5));
        Assert.True(narrow.Lines.Count >= 2);
        Assert.Equal("AV fi", string.Concat(narrow.Lines.Select(line =>
            narrow.Request.Text.Substring(line.SourceStart, line.SourceLength))));
    }

    [Fact]
    public void NoBreakSpacesStayAttachedWhenTheLineOverflows()
    {
        var face = Noto();
        foreach (var space in new[] { '\u00a0', '\u2007', '\u202f' })
        {
            var text = "A" + space + "B";
            Assert.True(face.Covers(space));
            var result = HarfBuzzTextPrelayout.Create().Layout(
                new TextPrelayoutRequest(text, new[] { face }, 4, 1));
            Assert.True(result.Overflows);
            Assert.Single(result.Lines);
            Assert.Equal(text, result.Lines[0].ShapedRun.Request.Text);
        }
    }

    [Fact]
    public void TrailingBreakSpaceDoesNotOverflowFittingWords()
    {
        var face = Noto();
        var result = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV To", new[] { face }, 4, 5));
        Assert.Equal(new[] { "AV ", "To" }, result.Lines.Select(line => line.ShapedRun.Request.Text));
        Assert.True(result.Lines[0].AdvanceMillimeters > 5);
        Assert.False(result.Lines[0].OverflowsWidth);
        Assert.False(result.Overflows);
        Assert.Single(TextPrelayoutOfdWriter.CreateSinglePage(result, 100, 100, 10, 10).Pages);
    }

    [Fact]
    public void LaterShapedOpportunityCanFitAfterAnEarlierOneOverflows()
    {
        var shaper = new NonMonotonicShaper();
        var result = new TextPrelayoutEngine(shaper).Layout(
            new TextPrelayoutRequest("A A", new[] { Noto() }, 4, 5));
        Assert.Equal(new[] { "A ", "A", "A A" }, shaper.Requests);
        Assert.Single(result.Lines);
        Assert.Equal("A A", result.Lines[0].ShapedRun.Request.Text);
        Assert.False(result.Overflows);
    }

    [Fact]
    public void ComplexCmapCoverageIsBoundedBeforeNativeShaping()
    {
        var bytes = File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf"));
        var tableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(4, 2));
        var cmap = -1;
        for (var i = 0; i < tableCount; i++)
        {
            var record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(bytes, record, 4) == "cmap")
                cmap = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4));
        }
        Assert.True(cmap >= 0);
        var subtableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(cmap + 2, 2));
        for (var i = 0; i < subtableCount; i++)
        {
            var record = cmap + 4 + i * 8;
            var offset = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 4, 4));
            if (BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(cmap + offset, 2)) == 12)
                BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(record, 2), 1);
        }
        var first = new TextFontFace("WenKai 1", bytes);
        var second = new TextFontFace("WenKai 2", bytes);
        var covered = Enumerable.Range(0x4e00, 0x5200).Where(first.Covers).Take(4095).ToArray();
        Assert.Equal(4095, covered.Length);
        Assert.False(first.Covers('\u0378'));
        var text = new string(covered.Select(scalar => (char)scalar).ToArray()) + '\u0378';
        var error = Assert.Throws<TextLayoutException>(() => HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest(text, new[] { first, second }, 4, 100,
                script: "Hani", language: "zh")));
        Assert.Equal(TextLayoutError.BudgetExceeded, error.Error);
    }

    [Fact]
    public void ExplicitWholeRunFallbackAndFailureRemainObservable()
    {
        var noto = Noto();
        var wenkai = new TextFontFace("WenKai", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var request = new TextPrelayoutRequest("中文", new[] { noto, wenkai }, 4, 30, script: "Hani", language: "zh");
        var result = HarfBuzzTextPrelayout.Create().Layout(request);
        Assert.Same(wenkai, result.Face);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "TEXT_FACE_FALLBACK");
        Assert.Equal(2, result.Lines[0].ShapedRun.Glyphs.Count);
        var missing = Assert.Throws<TextLayoutException>(() => HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("中文", new[] { noto }, 4, 30, script: "Hani", language: "zh")));
        Assert.Equal(TextLayoutError.MissingGlyph, missing.Error);
        Assert.Equal(0, missing.SourceStart);
    }

    [Fact]
    public void SameDisplayNameStyleAndTtcIndexNeverAliasTheSelectedFace()
    {
        var regular = new TextFontFace("Alias", File.ReadAllBytes(OtherFont("regular.ttf")));
        var bold = new TextFontFace("Alias", File.ReadAllBytes(OtherFont("bold.ttf")));
        var italic = new TextFontFace("Alias", File.ReadAllBytes(OtherFont("italic.ttf")));
        Assert.NotEqual(regular.FaceSha256, bold.FaceSha256);
        Assert.NotEqual(regular.FaceSha256, italic.FaceSha256);
        Assert.False(regular.Bold);
        Assert.True(bold.Bold);
        Assert.True(italic.Italic);
        var collection = File.ReadAllBytes(CollectionFont());
        var first = new TextFontFace("Alias", collection, 0);
        Assert.Equal(0, first.CollectionFaceIndex);
        Assert.NotEqual(first.SourceSha256, first.FaceSha256);
        Assert.Equal(TextLayoutError.InvalidFont, Assert.Throws<TextLayoutException>(() =>
            new TextFontFace("Alias", collection, 1)).Error);
    }

    [Fact]
    public void InvalidInputAndLegacyFlowBehaviorStaySeparate()
    {
        var face = Noto();
        Assert.Equal(TextLayoutError.InvalidText, Assert.Throws<TextLayoutException>(() =>
            new TextShapeRequest("A\ud800", 0, face)).Error);
        Assert.Equal(TextLayoutError.InvalidText, Assert.Throws<TextLayoutException>(() =>
            new TextPrelayoutRequest("A\rB", new[] { face }, 4, 20)).Error);
        var old = new FlowDocument();
        old.Blocks.Add(new Paragraph("AV fi"));
        Assert.Single(old.Render().Pages);
        Assert.Single(old.Render().Pages[0].Elements);
    }

    [Theory]
    [InlineData(0xd800)]
    [InlineData(0xdc00)]
    public void UnpairedSurrogateErrorsPreserveAbsoluteSourceInterval(int codeUnit)
    {
        var text = new string(new[] { 'A', (char)codeUnit });
        var direct = Assert.Throws<TextLayoutException>(() => new TextShapeRequest(text, 7, Noto()));
        Assert.Equal(TextLayoutError.InvalidText, direct.Error);
        Assert.Equal(8, direct.SourceStart);
        Assert.Equal(1, direct.SourceLength);

        var paragraph = Assert.Throws<TextLayoutException>(() =>
            new TextPrelayoutRequest(text, new[] { Noto() }, 4, 50));
        Assert.Equal(TextLayoutError.InvalidText, paragraph.Error);
        Assert.Equal(1, paragraph.SourceStart);
        Assert.Equal(1, paragraph.SourceLength);
    }

    [Fact]
    public void PrelayoutRejectsUnsupportedControlsBeforeFontCoverage()
    {
        var face = Noto();
        foreach (var text in new[] { "A\ufe0f", "A\udb40\udd00", "A\u200d" })
        {
            var error = Assert.Throws<TextLayoutException>(() =>
                new TextPrelayoutRequest(text, new[] { face }, 4, 20));
            Assert.Equal(TextLayoutError.UnsupportedFontProfile, error.Error);
            Assert.Equal(1, error.SourceStart);
        }
    }

    [Fact]
    public void MandatoryUnicodeSeparatorsRequireAnExplicitSupportedParagraphProfile()
    {
        var face = Noto();
        foreach (var separator in new[] { '\u000b', '\u000c', '\u0085', '\u2028', '\u2029' })
        {
            var text = "A" + separator + "B";
            var layoutError = Assert.Throws<TextLayoutException>(() =>
                new TextPrelayoutRequest(text, new[] { face }, 4, 100));
            Assert.Equal(TextLayoutError.InvalidText, layoutError.Error);
            Assert.Equal(1, layoutError.SourceStart);
            var shapeError = Assert.Throws<TextLayoutException>(() =>
                new TextShapeRequest(text, 7, face));
            Assert.Equal(TextLayoutError.InvalidText, shapeError.Error);
            Assert.Equal(8, shapeError.SourceStart);
        }
    }

    [Fact]
    public void SupplementaryHanUsesScalarSafeLineOpportunities()
    {
        var face = new TextFontFace("WenKai", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        const string text = "\U00020003\U0002000B\U0002000E";
        Assert.True(face.Covers(0x20003));
        Assert.True(face.Covers(0x2000b));
        Assert.True(face.Covers(0x2000e));
        var result = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest(text, new[] { face }, 4, 5, script: "Hani", language: "zh"));
        Assert.Equal(new[] { 0, 2, 4 }, result.Lines.Select(line => line.SourceStart));
        Assert.All(result.Lines, line => Assert.Equal(2, line.SourceLength));
        Assert.False(result.Overflows);
    }

    [Fact]
    public void HanBreaksKeepClosingPunctuationAndNoBreakSpacesAttached()
    {
        var face = new TextFontFace("WenKai", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var engine = HarfBuzzTextPrelayout.Create();
        TextPrelayoutResult Layout(string text, double width) => engine.Layout(
            new TextPrelayoutRequest(text, new[] { face }, 4, width, script: "Hani", language: "zh"));
        var firstLineWidth = Layout("中、", 100).WidthMillimeters;
        var punctuation = Layout("中、文", firstLineWidth + 0.01);
        Assert.Equal(new[] { "中、", "文" }, punctuation.Lines.Select(line => line.ShapedRun.Request.Text));
        Assert.False(punctuation.Overflows);

        Assert.True(face.Covers('\u00a0'));
        var glued = Layout("中\u00a0文", 5);
        Assert.Single(glued.Lines);
        Assert.Equal("中\u00a0文", glued.Lines[0].ShapedRun.Request.Text);
        Assert.True(glued.Overflows);
    }

    [Fact]
    public void MixedLatinHanCanBreakImmediatelyBeforeIdeograph()
    {
        var face = new TextFontFace("WenKai", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var engine = HarfBuzzTextPrelayout.Create();
        TextPrelayoutResult Layout(string text, double width) => engine.Layout(
            new TextPrelayoutRequest(text, new[] { face }, 4, width, script: "Hani", language: "zh"));
        var width = Math.Max(Layout("A", 100).WidthMillimeters,
            Layout("中", 100).WidthMillimeters) + 0.01;
        Assert.True(Layout("A中", 100).WidthMillimeters > width);
        var result = Layout("A中", width);
        Assert.Equal(new[] { "A", "中" }, result.Lines.Select(line => line.ShapedRun.Request.Text));
        Assert.False(result.Overflows);
    }

    [Fact]
    public void UnsupportedCmapFormatHasStructuredFontProfileFailure()
    {
        var bytes = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var tableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(4, 2));
        var cmap = -1;
        for (var i = 0; i < tableCount; i++)
        {
            var record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(bytes, record, 4) == "cmap")
                cmap = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4));
        }
        Assert.True(cmap >= 0);
        var subtableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(cmap + 2, 2));
        var changed = 0;
        for (var i = 0; i < subtableCount; i++)
        {
            var offset = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap + 4 + i * 8 + 4, 4));
            var subtable = cmap + offset;
            var format = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable, 2));
            if (format != 4 && format != 12) continue;
            bytes[subtable] = 0;
            bytes[subtable + 1] = 6;
            changed++;
        }
        Assert.True(changed > 0);
        var error = Assert.Throws<TextLayoutException>(() => new TextFontFace("Noto Sans", bytes));
        Assert.Equal(TextLayoutError.UnsupportedFontProfile, error.Error);
    }

    [Fact]
    public void DeferredCmapFailureRemainsStructuredDuringCoverageLookup()
    {
        var bytes = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var tableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(4, 2));
        var cmap = -1;
        for (var i = 0; i < tableCount; i++)
        {
            var record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(bytes, record, 4) == "cmap")
                cmap = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4));
        }
        Assert.True(cmap >= 0);
        var subtableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(cmap + 2, 2));
        var mutated = false;
        for (var i = 0; i < subtableCount; i++)
        {
            var offset = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(cmap + 4 + i * 8 + 4, 4));
            var subtable = cmap + offset;
            if (BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable, 2)) != 4) continue;
            var segmentCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable + 6, 2)) / 2;
            for (var segment = 0; segment < segmentCount; segment++)
            {
                var end = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable + 14 + segment * 2, 2));
                var start = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(subtable + 16 + segmentCount * 2 + segment * 2, 2));
                if (start > 'A' || end < 'A') continue;
                BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(subtable + 16 + segmentCount * 6 + segment * 2, 2), 0xfffe);
                mutated = true;
                break;
            }
        }
        Assert.True(mutated);
        var face = new TextFontFace("Noto Sans", bytes);
        var error = Assert.Throws<TextLayoutException>(() => face.Covers('A'));
        Assert.Equal(TextLayoutError.InvalidFont, error.Error);
        Assert.IsType<System.IO.InvalidDataException>(error.InnerException);
    }

    [Fact]
    public void ReducedMaxpCannotPassFaceValidation()
    {
        var bytes = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var tableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(4, 2));
        var maxp = -1;
        for (var i = 0; i < tableCount; i++)
        {
            var record = 12 + i * 16;
            if (System.Text.Encoding.ASCII.GetString(bytes, record, 4) == "maxp")
                maxp = (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4));
        }
        Assert.True(maxp >= 0);
        BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(maxp + 4, 2), 2);
        var error = Assert.Throws<TextLayoutException>(() => new TextFontFace("Invalid", bytes));
        Assert.Equal(TextLayoutError.InvalidFont, error.Error);
    }

    [Theory]
    [InlineData("maxp-version")]
    [InlineData("maxp-count")]
    [InlineData("hhea-count")]
    [InlineData("loca-offset")]
    public void MalformedGlyphTableDimensionsFailBeforeShaping(string mutation)
    {
        var bytes = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var tableCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(4, 2));
        var tables = new System.Collections.Generic.Dictionary<string, (int Offset, int Length)>();
        for (var i = 0; i < tableCount; i++)
        {
            var record = 12 + i * 16;
            tables.Add(System.Text.Encoding.ASCII.GetString(bytes, record, 4),
                ((int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 8, 4)),
                 (int)BinaryPrimitives.ReadUInt32BigEndian(bytes.AsSpan(record + 12, 4))));
        }
        var maxp = tables["maxp"].Offset;
        var glyphCount = BinaryPrimitives.ReadUInt16BigEndian(bytes.AsSpan(maxp + 4, 2));
        switch (mutation)
        {
            case "maxp-version":
                BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(maxp, 4), 0x00005000);
                break;
            case "maxp-count":
                BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(maxp + 4, 2), (ushort)(glyphCount + 1));
                break;
            case "hhea-count":
                BinaryPrimitives.WriteUInt16BigEndian(bytes.AsSpan(tables["hhea"].Offset + 34, 2), 0);
                break;
            case "loca-offset":
                var location = tables["loca"].Offset + glyphCount * 4;
                BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(location, 4), (uint)(tables["glyf"].Length + 1));
                break;
            default: throw new ArgumentOutOfRangeException(nameof(mutation));
        }
        var error = Assert.Throws<TextLayoutException>(() => new TextFontFace("Invalid", bytes));
        Assert.Equal(TextLayoutError.InvalidFont, error.Error);
    }

    [Fact]
    public void CumulativeFontCopiesStopWideHanCandidateScan()
    {
        var face = new TextFontFace("WenKai", File.ReadAllBytes(Font("LXGWWenKai-Regular.ttf")));
        var shaper = new CountingShaper();
        var error = Assert.Throws<TextLayoutException>(() => new TextPrelayoutEngine(shaper).Layout(
            new TextPrelayoutRequest(new string('中', 4096), new[] { face }, 4, 1_000_000,
                script: "Hani", language: "zh", maximumShapeCalls: 8192)));
        Assert.Equal(TextLayoutError.BudgetExceeded, error.Error);
        Assert.InRange(shaper.Calls, 1, 50);
    }

    [Fact]
    public void CancellationAndShapeBudgetFailWithoutAnyPartialLayout()
    {
        var face = Noto();
        var engine = HarfBuzzTextPrelayout.Create();
        using var canceled = new CancellationTokenSource();
        canceled.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => engine.Layout(
            new TextPrelayoutRequest("AV To", new[] { face }, 4, 3), canceled.Token));
        var error = Assert.Throws<TextLayoutException>(() => engine.Layout(
            new TextPrelayoutRequest("AV To", new[] { face }, 4, 3, maximumShapeCalls: 1)));
        Assert.Equal(TextLayoutError.BudgetExceeded, error.Error);
    }

    [Fact]
    public async Task RepresentableAsciiWriterReusesMeasuredRunAndExactEmbeddedFace()
    {
        var face = Noto();
        var layout = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV A\nTo", new[] { face }, 4, 50));
        var package = TextPrelayoutOfdWriter.CreateSinglePage(layout, 100, 100, 10, 15);
        var text = package.Pages[0].Elements.OfType<OfdTextElement>().ToArray();
        Assert.Equal(new[] { "AV A", "To" }, text.Select(element => element.Text));
        for (var i = 0; i < text.Length; i++)
        {
            Assert.Equal(15 + layout.Lines[i].TopMillimeters, text[i].YMillimeters, 6);
            Assert.Equal(layout.Lines[i].BaselineMillimeters - layout.Lines[i].TopMillimeters,
                text[i].Runs[0].YMillimeters, 6);
            var deltas = text[i].Runs[0].DeltaX!.Split(' ').Select(value =>
                double.Parse(value, System.Globalization.CultureInfo.InvariantCulture)).ToArray();
            Assert.Equal(layout.Lines[i].ShapedRun.Glyphs.Count - 1, deltas.Length);
            for (var j = 0; j < deltas.Length; j++) Assert.Equal(
                layout.Lines[i].ShapedRun.Glyphs[j].XAdvance * 4d / face.UnitsPerEm, deltas[j], 10);
        }
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var read = await new OfdReader().ReadAsync(stream);
        Assert.Equal(new[] { "AV A", "To" }, read.Pages[0].Elements.OfType<OfdTextElement>().Select(element => element.Text));
        var embedded = Assert.Single(read.Fonts).Data;
        Assert.Equal(face.FaceSha256, Convert.ToHexString(SHA256.HashData(embedded)).ToLowerInvariant());
        var ligature = HarfBuzzTextPrelayout.Create().Layout(new TextPrelayoutRequest("ffi", new[] { face }, 4, 50));
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePage(ligature, 100, 100, 10, 15)).Error);
    }

    [Fact]
    public async Task PositionedWriterPreservesFfiLigatureAndActualGidInSubset()
    {
        var face = Noto();
        var engine = HarfBuzzTextPrelayout.Create();
        var enabled = engine.Layout(new TextPrelayoutRequest("ffi", new[] { face }, 4, 50));
        var disabled = engine.Layout(new TextPrelayoutRequest("ffi", new[] { face }, 4, 50,
            features: new[] { new TextFeature("liga", 0, 0, 3) }));
        Assert.Single(enabled.Lines[0].ShapedRun.Glyphs);
        Assert.Equal(3, disabled.Lines[0].ShapedRun.Glyphs.Count);

        foreach (var layout in new[] { enabled, disabled })
        {
            var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout,
                100, 100, 10, 15);
            var original = Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>());
            Assert.Equal("ffi", original.Runs[0].PositionedGlyphs!.LogicalText);
            Assert.Equal(face.FaceSha256, original.Runs[0].PositionedGlyphs!.FaceSha256);
            Assert.Equal(layout.Lines[0].ShapedRun.Glyphs.Select(glyph => glyph.GlyphId),
                original.Runs[0].PositionedGlyphs!.Glyphs.Select(glyph => glyph.GlyphId));
            using var stream = new MemoryStream();
            var write = await new OfdPackageWriter().WriteWithResultAsync(package, stream);
            Assert.True(Assert.Single(write.FontEmbedding).IsSubset);
            stream.Position = 0;
            using (var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true))
            {
                var page = zip.Entries.Single(entry => entry.FullName.EndsWith("Content.xml", StringComparison.Ordinal));
                using var pageStream = page.Open();
                var xml = XDocument.Load(pageStream);
                var mapping = Assert.Single(xml.Descendants(), node => node.Name.LocalName == "CGTransform");
                Assert.Equal("0", mapping.Attribute("CodePosition")?.Value);
                Assert.Equal("3", mapping.Attribute("CodeCount")?.Value);
                Assert.Equal(layout.Lines[0].ShapedRun.Glyphs.Count.ToString(), mapping.Attribute("GlyphCount")?.Value);
                Assert.Equal(string.Join(" ", layout.Lines[0].ShapedRun.Glyphs.Select(glyph => glyph.GlyphId)),
                    Assert.Single(mapping.Elements()).Value);
            }
            stream.Position = 0;
            var read = await new OfdReader().ReadAsync(stream);
            Assert.Equal("ffi", Assert.Single(read.Pages[0].Elements.OfType<OfdTextElement>()).Text);
            Assert.True(Assert.Single(read.Fonts).Data.Length < face.GetFaceBytes().Length);
        }
    }

    [Fact]
    public async Task PositionedWriterRejectsFaceMismatchAndImportedSourceBeforeZip()
    {
        var face = Noto();
        var layout = HarfBuzzTextPrelayout.Create().Layout(new TextPrelayoutRequest("AV", new[] { face }, 4, 50));
        var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout,
            100, 100, 10, 15);
        var text = Assert.Single(package.Pages[0].Elements.OfType<OfdTextElement>());
        var run = Assert.Single(text.Runs);
        run.PositionedGlyphs = new OfdPositionedGlyphRun(run.Text, 0, new string('0', 64),
            run.PositionedGlyphs!.Glyphs);
        using var mismatch = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, mismatch));
        Assert.Equal(0, mismatch.Length);

        run.PositionedGlyphs = new OfdPositionedGlyphRun(run.Text, 0, face.FaceSha256,
            run.PositionedGlyphs.Glyphs);
        text.SourceXml = "<TextObject><TextCode X='0' Y='0'>AV</TextCode></TextObject>";
        using var imported = new MemoryStream();
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageWriter().WriteAsync(package, imported));
        Assert.Equal(0, imported.Length);

        var nonAscii = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("é", new[] { face }, 4, 50));
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(nonAscii,
                100, 100, 10, 15)).Error);
    }

    [Fact]
    public async Task PositionedWriterValidatesOriginalOutlineEvenWhenFontIsPreserved()
    {
        var face = Noto();
        var layout = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV", new[] { face }, 4, 50));
        var valid = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout,
            100, 100, 10, 15);
        valid.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
        using (var stream = new MemoryStream())
        {
            var write = await new OfdPackageWriter().WriteWithResultAsync(valid, stream);
            Assert.False(Assert.Single(write.FontEmbedding).IsSubset);
            Assert.True(stream.Length > 0);
        }

        foreach (var preserved in new[] { false, true })
        {
            var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout,
                100, 100, 10, 15);
            if (preserved) package.Pages[0].PreservedPageElements.Add("<Unknown/>");
            else package.Options.FontEmbedding.Mode = OfdFontEmbeddingMode.Full;
            var run = package.Pages[0].Elements.OfType<OfdTextElement>().Single().Runs.Single();
            var current = run.PositionedGlyphs!;
            var first = current.Glyphs[0];
            var missing = new OfdPositionedGlyph((uint)face.GlyphCount, first.ClusterStart,
                first.ClusterLength, first.OriginXMillimeters, first.OriginYMillimeters,
                first.XAdvanceMillimeters, first.YAdvanceMillimeters);
            run.PositionedGlyphs = new OfdPositionedGlyphRun(current.LogicalText,
                current.SourceStart, current.FaceSha256,
                new[] { missing }.Concat(current.Glyphs.Skip(1)).ToArray());
            using var output = new MemoryStream();
            await Assert.ThrowsAsync<NotSupportedException>(() =>
                new OfdPackageWriter().WriteAsync(package, output));
            Assert.Equal(0, output.Length);
        }
    }

    [Fact]
    public async Task PositionedWriterRejectsOneGidForDistinctSourceCharactersBeforeZip()
    {
        var face = Noto();
        var ambiguous = new TextPrelayoutEngine(new AmbiguousGidShaper()).Layout(
            new TextPrelayoutRequest("AV", new[] { face }, 4, 50));
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(ambiguous,
                100, 100, 10, 15)).Error);
        var layout = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV", new[] { face }, 4, 50));
        var package = TextPrelayoutOfdWriter.CreateSinglePageWithPositionedGlyphs(layout,
            100, 100, 10, 15);
        var run = package.Pages[0].Elements.OfType<OfdTextElement>().Single().Runs.Single();
        var original = run.PositionedGlyphs!;
        var first = original.Glyphs[0];
        var second = original.Glyphs[1];
        run.PositionedGlyphs = new OfdPositionedGlyphRun(original.LogicalText, original.SourceStart,
            original.FaceSha256, new[] { first, new OfdPositionedGlyph(first.GlyphId,
                second.ClusterStart, second.ClusterLength, second.OriginXMillimeters,
                second.OriginYMillimeters, second.XAdvanceMillimeters, second.YAdvanceMillimeters) });
        using var output = new MemoryStream();
        await Assert.ThrowsAsync<NotSupportedException>(() =>
            new OfdPackageWriter().WriteAsync(package, output));
        Assert.Equal(0, output.Length);
    }

    [Fact]
    public async Task WriterRejectsXmlInvalidFaceNameBeforeCreatingOutput()
    {
        var bytes = File.ReadAllBytes(Font("NotoSans-Regular.ttf"));
        var invalid = new TextFontFace("Noto\0Sans", bytes);
        var layout = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV", new[] { invalid }, 4, 50));
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePage(layout, 100, 100, 10, 10)).Error);

        var valid = new TextFontFace("Noto & Sans", bytes);
        var validLayout = HarfBuzzTextPrelayout.Create().Layout(
            new TextPrelayoutRequest("AV", new[] { valid }, 4, 50));
        using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(
            TextPrelayoutOfdWriter.CreateSinglePage(validLayout, 100, 100, 10, 10), stream);
        stream.Position = 0;
        var read = await new OfdReader().ReadAsync(stream);
        Assert.Equal("Noto & Sans", Assert.Single(read.Fonts).FontName);
    }

    [Fact]
    public void WriterChecksActualAdvanceAndInkAtPageEdges()
    {
        var face = Noto();
        var engine = HarfBuzzTextPrelayout.Create();
        var shortLine = engine.Layout(new TextPrelayoutRequest("AV", new[] { face }, 4, 100));
        Assert.True(shortLine.WidthMillimeters < 50);
        Assert.Single(TextPrelayoutOfdWriter.CreateSinglePage(shortLine, 50, 50, 1, 10).Pages);

        var overhanging = engine.Layout(new TextPrelayoutRequest("J", new[] { face }, 4, 100));
        var jInk = overhanging.Lines[0].InkBoundsMillimeters!;
        Assert.True(jInk.X < 0);
        Assert.Equal(TextLayoutError.UnsupportedOutput, Assert.Throws<TextLayoutException>(() =>
            TextPrelayoutOfdWriter.CreateSinglePage(overhanging, 50, 50, 0, 10)).Error);
        var j = Assert.IsType<OfdTextElement>(Assert.Single(
            Assert.Single(TextPrelayoutOfdWriter.CreateSinglePage(overhanging, 50, 50, 1, 10).Pages).Elements));
        Assert.True(j.XMillimeters < 1);
        Assert.Equal(1, j.XMillimeters + j.Runs[0].XMillimeters, 6);
        Assert.True(j.XMillimeters <= 1 + jInk.X);
        Assert.True(j.XMillimeters + j.WidthMillimeters >= 1 + jInk.X + jInk.Width);

        var rightOverhang = engine.Layout(new TextPrelayoutRequest("f", new[] { face }, 4, 100));
        var ink = rightOverhang.Lines[0].InkBoundsMillimeters!;
        Assert.True(ink.X + ink.Width > rightOverhang.Lines[0].AdvanceMillimeters);
        var f = Assert.IsType<OfdTextElement>(Assert.Single(
            Assert.Single(TextPrelayoutOfdWriter.CreateSinglePage(rightOverhang, 50, 50, 1, 10).Pages).Elements));
        Assert.Equal(1, f.XMillimeters + f.Runs[0].XMillimeters, 6);
        Assert.True(f.XMillimeters + f.WidthMillimeters >= 1 + ink.X + ink.Width);
    }

    private static TextFontFace Noto() => new("Noto Sans", File.ReadAllBytes(Font("NotoSans-Regular.ttf")));

    private sealed class AmbiguousGidShaper : ITextShaper
    {
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
            => new(request, request.Text.Select((_, index) => new TextShapedGlyph(36,
                request.SourceStart + index, 1, 600, 0, 0, 0, null, false)).ToArray(),
                "ambiguous-test");
    }

    private sealed class CountingShaper : ITextShaper
    {
        public int Calls { get; private set; }
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
        {
            Calls++;
            return new TextShapedRun(request, new[] { new TextShapedGlyph(1, request.SourceStart,
                request.Text.Length, request.Text.Length * 500, 0, 0, 0, null, false) }, "counting-test");
        }
    }

    private sealed class NonMonotonicShaper : ITextShaper
    {
        public System.Collections.Generic.List<string> Requests { get; } = new();
        public TextShapedRun Shape(TextShapeRequest request, CancellationToken cancellationToken = default)
        {
            Requests.Add(request.Text);
            var advance = request.Text.Length == 3 ? 1000 : 2000;
            return new TextShapedRun(request, new[] { new TextShapedGlyph(1, request.SourceStart,
                request.Text.Length, advance, 0, 0, 0, null, false) }, "nonmonotonic-test");
        }
    }

    private static string Font(string file) => Fixture("Ofdrw.Net.FontSubset.E2E", file);
    private static string OtherFont(string file) => Fixture("Ofdrw.Net.SkiaSharp.E2E", file);
    private static string CollectionFont() => Fixture("Ofdrw.Net.Converter.Pdf.E2E", "style-collection.ttc");

    private static string Fixture(string project, string file)
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            var candidate = Path.Combine(directory.FullName, "e2e", project, "testdata", "fonts", file);
            if (File.Exists(candidate)) return candidate;
            directory = directory.Parent;
        }
        throw new FileNotFoundException("Materialize the licensed font fixture through Git LFS.", file);
    }
}
