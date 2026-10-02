using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using PdfSharpCore.Fonts;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace Ofdrw.Net.Converter.Pdf.Tests;

[CollectionDefinition("Font registry", DisableParallelization = true)]
public sealed class FontRegistryCollection;

[Collection("Font registry")]
public sealed class FontIsolationTests
{
    public FontIsolationTests() => PdfFontRegistry.EnsureInstalled();

    [Theory]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void NameOnlyFallback_ShouldDetectMissingStylesEvenWhenHostResolverOmitsSimulationFlags(bool bold, bool italic)
    {
        var data = File.ReadAllBytes(FontPath("narrow"));
        var host = new HostResolver(data);
        var resource = new OfdFontResource { Id = "12", FontName = "host-CJK-substitute", Bold = bold, Italic = italic };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
        var family = context.Resolve(new OfdTextElement { FontResourceId = "12" }, out var resolved);
        Assert.Same(resource, resolved);
        var face = GlobalFontSettings.FontResolver.ResolveTypeface(family, bold, italic);
        Assert.Equal(bold, face.MustSimulateBold);
        Assert.Equal(italic, face.MustSimulateItalic);
        Assert.Equal(data, PdfFontRegistry.GetOriginalFont(face.FaceName));
        Assert.Equal(1, host.Requests);
    }

    [Theory]
    [InlineData("resolve")]
    [InlineData("read")]
    [InlineData("ttc")]
    [InlineData("woff")]
    [InlineData("broken-sfnt")]
    public void NameOnlyFallback_ShouldIgnoreFailedOrUnsupportedHostProbes(string failure)
    {
        var host = new ProbeResolver(failure);
        var resource = new OfdFontResource { Id = "probe", FontName = "optional-host", Bold = true };
        var retained = PdfFontRegistry.RegisteredFontBytes;
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
        Assert.Equal("optional-host", context.Resolve(new OfdTextElement { FontResourceId = "probe", FontName = "optional-host" }, out _));
        Assert.Equal(retained, PdfFontRegistry.RegisteredFontBytes);
    }

    [Theory]
    [InlineData(false)][InlineData(true)]
    public async Task ActualPdfSkipsInvisibleControlsWithoutChangingExplicitDeltaSlots(bool positioned)
    {
        var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("narrow")));
        var cmap = new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face);
        face.Tables["cmap"] = Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int> { ['A'] = cmap.Glyph('A'), ['B'] = cmap.Glyph('B') });
        var data = face.Build();
        async Task<(string Text, double Left)> Export(string value)
        {
            var package = new OfdDocumentPackage(); package.Fonts.Add(new OfdFontResource { Id="10", FontName="control-probe", Data=data });
            var text = new OfdTextElement { FontResourceId="10", Text=value, XMillimeters=10, YMillimeters=10, FontSizeMillimeters=6, WidthMillimeters=80, HeightMillimeters=20 };
            if (positioned) text.Runs.Add(new OfdTextRun { Text=value, YMillimeters=6, DeltaX="5" });
            package.Pages.Add(new OfdPage { WidthMillimeters=100, HeightMillimeters=50, Elements={text} });
            using var ofd=new MemoryStream(); await new OfdPackageWriter().WriteAsync(package,ofd); ofd.Position=0;
            using var pdf=new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd,pdf);
            using var parsed=PdfPigDocument.Open(pdf.ToArray()); var page=parsed.GetPage(1);
            return (page.Text, page.Letters.Single(letter => letter.Value=="B").BoundingBox.Left);
        }
        var baseline=await Export("AB");
        foreach (var value in new[] { "AB" })
        {
            var result=await Export(value); Assert.Equal("AB",result.Text); Assert.Equal(baseline.Left,result.Left,5);
        }
    }

    [Theory]
    [InlineData(false)][InlineData(true)]
    public async Task ActualPdfKeepsMappedHangulFillerAdvanceInImplicitAndPositionedRuns(bool positioned)
    {
        var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("narrow")));
        var cmap = new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face); var space=cmap.Glyph(' '); Assert.NotEqual(0,space);
        face.Tables["cmap"] = Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int> { ['A']=cmap.Glyph('A'), ['B']=cmap.Glyph('B'), [' ']=space, [0x3164]=space });
        var data=face.Build();
        async Task<double> Export(string value)
        {
            var package=new OfdDocumentPackage(); package.Fonts.Add(new OfdFontResource { Id="10", FontName="mapped-filler", Data=data });
            var text=new OfdTextElement { FontResourceId="10", Text=value, XMillimeters=10, YMillimeters=10, FontSizeMillimeters=6, WidthMillimeters=80, HeightMillimeters=20 };
            if (positioned) text.Runs.Add(new OfdTextRun { Text=value, YMillimeters=6, DeltaX="5 5" });
            package.Pages.Add(new OfdPage { WidthMillimeters=100,HeightMillimeters=50,Elements={text} });
            using var ofd=new MemoryStream();await new OfdPackageWriter().WriteAsync(package,ofd);ofd.Position=0;
            using var pdf=new MemoryStream();await new OfdToPdfConverter().ConvertAsync(ofd,pdf);
            using var doc=PdfPigDocument.Open(pdf.ToArray());return doc.GetPage(1).Letters.Single(letter=>letter.Value=="B").BoundingBox.Left;
        }
        Assert.Equal(await Export("A B"),await Export("A\u3164B"),5);
    }

    [Fact]
    public void NameOnlyMissingOrdinaryGlyphUsesConfiguredCoveringFallbackOrFailsClearly()
    {
        var primary=File.ReadAllBytes(FontPath("narrow"));
        var fallback=new Ofdrw.Net.Core.Fonts.OpenTypeFace(primary);var cmap=new Ofdrw.Net.Core.Fonts.OpenTypeCmap(fallback);
        fallback.Tables["cmap"]=Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int>{[0x4E00]=cmap.Glyph('A')});
        var resource=new OfdFontResource { Id="10",FontName="primary-probe",Bold=true };
        var context=new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource],new CoveringResolver(primary,fallback.Build()));
        var text=new OfdTextElement { FontResourceId="10",Text="一" };var family=context.Resolve(text,out var selected);
        Assert.Same(resource,selected);Assert.StartsWith("ofd-font-",family);Assert.NotEqual(0,context.Coverage(resource,family)!.Glyph(0x4E00));
        var insufficient=new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource],new CoveringResolver(primary,primary));
        Assert.Throws<InvalidDataException>(()=>insufficient.Resolve(text,out _));
    }
    [Fact]
    public async Task SymbolCmapUnusedExportsAndSelectedRefusesUnmodeledSemantics()
    {
        var face=new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("narrow")));
        var cmap=new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face);var rebuilt=Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int>{[0xF041]=cmap.Glyph('A')});
        var offset=(int)Ofdrw.Net.Core.Fonts.OpenTypeFace.U32(rebuilt,8);var length=Ofdrw.Net.Core.Fonts.OpenTypeFace.U16(rebuilt,offset+2);
        var symbol=new byte[12+length];Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(symbol,2,1);Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(symbol,4,3);Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(symbol,8,12);
        Buffer.BlockCopy(rebuilt,offset,symbol,12,length);face.Tables["cmap"]=symbol;
        Assert.True(new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face).IsSymbol);Assert.Equal(cmap.Glyph('A'),new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face).Glyph('A'));
        foreach(var selected in new[]{false,true})
        {
            var package=new OfdDocumentPackage();package.Fonts.Add(new OfdFontResource{Id="10",FontName="symbol-probe",Data=face.Build()});
            package.Pages.Add(new OfdPage{WidthMillimeters=100,HeightMillimeters=50});
            if(selected)package.Pages[0].Elements.Add(new OfdTextElement{FontResourceId="10",Text="A"});
            using var ofd=new MemoryStream();var report=await new OfdPackageWriter().WriteWithResultAsync(package,ofd);
            Assert.False(report.FontEmbedding[0].IsSubset);ofd.Position=0;using var svg=new MemoryStream();
            if (selected)
            {
                await Assert.ThrowsAsync<NotSupportedException>(() => new Ofdrw.Net.Converter.Svg.Converters.OfdToSvgConverter().ConvertAsync(ofd, svg));
                Assert.Equal(0, svg.Length);
            }
            else
            {
                await new Ofdrw.Net.Converter.Svg.Converters.OfdToSvgConverter().ConvertAsync(ofd, svg);
                Assert.Contains("@font-face", System.Text.Encoding.UTF8.GetString(svg.ToArray()));
            }
            ofd.Position = 0; using var pdf = new MemoryStream();
            if (selected)
            {
                await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
                Assert.Equal(0, pdf.Length);
            }
            else await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        }
    }
    [Theory]
    [InlineData("null")][InlineData("resolve")][InlineData("read")][InlineData("cmap")][InlineData("ttc")][InlineData("default-name")]
    public void NameOnlyFallbackProbesFailedDefaultsAndCollectionFaces(string kind)
    {
        var primary = File.ReadAllBytes(FontPath("narrow"));
        var covering = new Ofdrw.Net.Core.Fonts.OpenTypeFace(primary);
        var cmap = new Ofdrw.Net.Core.Fonts.OpenTypeCmap(covering);
        covering.Tables["cmap"] = Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int> { [0x4E00] = cmap.Glyph('A') });
        var bytes = covering.Build();
        var resource = new OfdFontResource { Id = "10", FontName = "primary-probe", Bold = true };
        var host = new FailureDefaultResolver(primary, bytes, kind);
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "一" }, out _);
        Assert.NotEqual(0, context.Coverage(resource, family)!.Glyph(0x4E00));
        Assert.Equal(kind == "ttc" ? 0 : 1, host.ArialRequests);
    }

    private sealed class FailureDefaultResolver(byte[] primary, byte[] covering, string kind) : IFontResolver
    {
        public int ArialRequests { get; private set; }
        public string DefaultFontName => kind == "default-name" ? throw new ArgumentException("Optional default name unavailable.") : "optional-default";
        public FontResolverInfo ResolveTypeface(string family, bool bold, bool italic)
        {
            if (family == "Arial") { ArialRequests++; return new("arial-covering"); }
            if (family != "optional-default") return new("primary-face");
            if (kind == "null") return null!;
            if (kind == "resolve") throw new ArgumentException("Optional host lookup failed.");
            return new("default-face");
        }
        public byte[] GetFont(string face)
        {
            if (face == "primary-face") return primary;
            if (face == "arial-covering") return covering;
            if (kind == "read") throw new IOException("Optional host bytes unavailable.");
            if (kind == "ttc")
            {
                var collection = new byte[16 + covering.Length];
                Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(collection, 0, 0x74746366);
                Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(collection, 4, 0x00010000);
                Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(collection, 8, 1);
                Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(collection, 12, 16);
                Buffer.BlockCopy(covering, 0, collection, 16, covering.Length);
                for (var i = 0; i < Ofdrw.Net.Core.Fonts.OpenTypeFace.U16(covering, 4); i++)
                {
                    var offset = 16 + 12 + i * 16 + 8;
                    Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(collection, offset, Ofdrw.Net.Core.Fonts.OpenTypeFace.U32(collection, offset) + 16);
                }
                return collection;
            }
            var unsupported = new Ofdrw.Net.Core.Fonts.OpenTypeFace(covering);
            unsupported.Tables["cmap"] = Format0Cmap();
            return unsupported.Build();
        }
    }

    private static byte[] Format0Cmap()
    {
        var table = new byte[12 + 262];
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(table, 2, 1);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(table, 4, 3);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(table, 6, 1);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(table, 8, 12);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(table, 14, 262);
        return table;
    }

    [Fact]
    public async Task UnusedUnmodeledCmapDoesNotAbortPdfButSelectedFontRefuses()
    {
        var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("narrow")));
        face.Tables["cmap"] = Format0Cmap();
        var resource = new OfdFontResource { Id = "unsupported", FontName = "format-zero", Data = face.Build() };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource]);
        Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement { FontResourceId = resource.Id, Text = "A" }, out _));
        var package = new OfdDocumentPackage(); package.Fonts.Add(resource);
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50 });
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
        using var pdf = new MemoryStream(); await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
        using var parsed = PdfPigDocument.Open(pdf.ToArray()); Assert.Equal(1, parsed.NumberOfPages);
    }

    [Theory]
    [InlineData("A\u200CB")][InlineData("A\u200DB")]
    public async Task ActualPdfRefusesJoinControlsBeforeWritingBytes(string value)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", Data = File.ReadAllBytes(FontPath("narrow")) });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { new OfdTextElement { FontResourceId = "10", FontName = "symbol-unmapped", Text = value } } });
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
        using var pdf = new MemoryStream(); await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
        Assert.Equal(0, pdf.Length);
    }

    [Fact]
    public async Task RealNotoLigatureContextRetainsOfdTextAndRefusesJoinSemantics()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln"))) directory = directory.Parent;
        var bytes = File.ReadAllBytes(Path.Combine(directory!.FullName, "e2e/Ofdrw.Net.FontSubset.E2E/testdata/fonts/NotoSans-Regular.ttf"));
        foreach (var value in new[] { "office ffi", "of\u200Cfice ffi", "of\u200Dfice ffi" })
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "Noto Sans", Data = bytes });
            package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
                Elements = { new OfdTextElement { FontResourceId = "10", FontName = "Noto Sans", Text = value,
                    XMillimeters = 10, YMillimeters = 10, WidthMillimeters = 80, HeightMillimeters = 20, FontSizeMillimeters = 6 } } });
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            var read = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd);
            Assert.Equal(value, Assert.IsType<OfdTextElement>(read.Pages[0].Elements[0]).Text); ofd.Position = 0;
            using var pdf = new MemoryStream();
            if (value == "office ffi")
            {
                await new OfdToPdfConverter().ConvertAsync(ofd, pdf);
                using var parsed = PdfPigDocument.Open(pdf.ToArray()); Assert.Equal(value, parsed.GetPage(1).Text);
            }
            else
            {
                var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
                Assert.Contains("default-ignorable", exception.Message); Assert.Equal(0, pdf.Length);
            }
        }
    }

    [Theory]
    [InlineData("A ")][InlineData("AB")]
    public async Task SymbolCmapWithUnmappedTypedScalarsPreservesOriginalBytes(string value)
    {
        var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("narrow")));
        var cmap = new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face);
        var rebuilt = Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int> { [0xF041] = cmap.Glyph('A') });
        var offset = (int)Ofdrw.Net.Core.Fonts.OpenTypeFace.U32(rebuilt, 8);
        var length = Ofdrw.Net.Core.Fonts.OpenTypeFace.U16(rebuilt, offset + 2);
        var symbol = new byte[12 + length]; Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(symbol, 2, 1);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(symbol, 4, 3); Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(symbol, 8, 12);
        Buffer.BlockCopy(rebuilt, offset, symbol, 12, length); face.Tables["cmap"] = symbol;
        var bytes = face.Build(); var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "symbol-unmapped", Data = bytes });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50, Elements = { new OfdTextElement { FontResourceId = "10", FontName = "symbol-unmapped", Text = value } } });
        for (var round = 0; round < 2; round++)
        {
            using var ofd = new MemoryStream(); var result = await new OfdPackageWriter().WriteWithResultAsync(package, ofd);
            Assert.False(result.FontEmbedding[0].IsSubset); Assert.Contains(result.Diagnostics, d => d.Contains("FONT_COVERAGE_UNVERIFIED"));
            ofd.Position = 0; package = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd);
            Assert.Equal(bytes, package.Fonts.Single().Data);
            ofd.Position = 0; using var pdf = new MemoryStream();
            var pdfFailure = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
            Assert.Contains("unmodeled", pdfFailure.Message); Assert.Equal(0, pdf.Length);
            ofd.Position = 0; using var svg = new MemoryStream();
            var svgFailure = await Assert.ThrowsAsync<NotSupportedException>(() => new Ofdrw.Net.Converter.Svg.Converters.OfdToSvgConverter().ConvertAsync(ofd, svg));
            Assert.Contains("unmodeled", svgFailure.Message); Assert.Equal(0, svg.Length);
        }
    }

    private static byte[] SymbolFace(byte[] source)
    {
        var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(source);
        var cmap = new Ofdrw.Net.Core.Fonts.OpenTypeCmap(face);
        var rebuilt = Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int> { [0xF041] = cmap.Glyph('A') });
        var offset = (int)Ofdrw.Net.Core.Fonts.OpenTypeFace.U32(rebuilt, 8);
        var length = Ofdrw.Net.Core.Fonts.OpenTypeFace.U16(rebuilt, offset + 2);
        var symbol = new byte[12 + length]; Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(symbol, 2, 1);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(symbol, 4, 3); Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(symbol, 8, 12);
        Buffer.BlockCopy(rebuilt, offset, symbol, 12, length); face.Tables["cmap"] = symbol;
        return face.Build();
    }

    [Theory]
    [InlineData("A")][InlineData("A ")][InlineData("AB")]
    public void SelectedNameOnlySymbolProbeRefusesBeforeCoverageAndFallback(string value)
    {
        var resource = new OfdFontResource { Id = "10", FontName = "symbol-host-probe", Bold = true };
        var host = new HostResolver(SymbolFace(File.ReadAllBytes(FontPath("narrow"))));
        var before = PdfFontRegistry.RegisteredFontBytes;
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
        Assert.Equal(before, PdfFontRegistry.RegisteredFontBytes);
        var exception = Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement { FontResourceId = "10", Text = value }, out _));
        Assert.Contains("unmodeled", exception.Message); Assert.Equal(1, host.Requests);
    }

    [Fact]
    public void SymbolDefaultCannotHalfVerifyMappedAAsUnicodeFallback()
    {
        var source = File.ReadAllBytes(FontPath("narrow"));
        var primary = new Ofdrw.Net.Core.Fonts.OpenTypeFace(source);
        var cmap = new Ofdrw.Net.Core.Fonts.OpenTypeCmap(primary);
        primary.Tables["cmap"] = Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int> { ['B'] = cmap.Glyph('B') });
        var symbol = SymbolFace(source);
        var resource = new OfdFontResource { Id = "10", FontName = "primary-unicode", Bold = true };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], new SymbolDefaultResolver(primary.Build(), symbol, source));
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "A" }, out _);
        var coverage = context.Coverage(resource, family)!;
        Assert.False(coverage.IsSymbol); Assert.Equal(cmap.Glyph('A'), coverage.Glyph('A'));
        var insufficient = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], new SymbolDefaultResolver(primary.Build(), symbol, symbol));
        Assert.Throws<InvalidDataException>(() => insufficient.Resolve(new OfdTextElement { FontResourceId = "10", Text = "A" }, out _));
    }

    private sealed class SymbolDefaultResolver(byte[] primary, byte[] symbol, byte[] arial) : IFontResolver
    {
        public string DefaultFontName => "symbol-default";
        public FontResolverInfo ResolveTypeface(string family, bool bold, bool italic) => new(family == DefaultFontName ? "symbol-face" : family == "Arial" ? "unicode-arial" : "primary-face");
        public byte[] GetFont(string face) => face == "symbol-face" ? symbol : face == "unicode-arial" ? arial : primary;
    }

    [Theory]
    [InlineData(false, "A")][InlineData(false, "A ")][InlineData(false, "AB")][InlineData(true, "A")]
    public void RegularAndUnboundSelectedSymbolFontsRefuseLazily(bool unbound, string value)
    {
        var host = new HostResolver(SymbolFace(File.ReadAllBytes(FontPath("narrow"))));
        var resource = new OfdFontResource { Id = "10", FontName = "regular-symbol" };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext(unbound ? [] : [resource], host);
        Assert.Equal(0, host.Requests);
        Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement { FontResourceId = unbound ? null : "10", FontName = "regular-symbol", Text = value }, out _));
        Assert.Equal(1, host.Requests);
    }

    [Theory]
    [InlineData(false)][InlineData(true)]
    public void CachedNameOnlyCoverageIsRecheckedAndEmphasisUsesSamePhysicalFace(bool resourceBold)
    {
        var bytes = File.ReadAllBytes(FontPath("narrow"));
        var host = new HostResolver(bytes);
        var resource = new OfdFontResource { Id = "10", FontName = "regular-unicode", Bold = resourceBold };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "A", Weight = 700, Italic = true }, out _);
        var face = GlobalFontSettings.FontResolver.ResolveTypeface(family, true, true);
        Assert.StartsWith("ofd:", face.FaceName); Assert.Equal(bytes, PdfFontRegistry.GetOriginalFont(face.FaceName));
        Assert.True(face.MustSimulateBold); Assert.True(face.MustSimulateItalic);
        Assert.Throws<InvalidDataException>(() => context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "一" }, out _));
    }

    [Fact]
    public void AddedEmphasisUsesImmutableSnapshotOfPreviouslySelectedPhysicalFont()
    {
        var bytes = File.ReadAllBytes(FontPath("narrow")); var original = bytes.ToArray();
        var resource = new OfdFontResource { Id = "10", FontName = "snapshot-host", Bold = true };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], new HostResolver(bytes));
        Array.Clear(bytes);
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "A", Italic = true }, out _);
        var face = GlobalFontSettings.FontResolver.ResolveTypeface(family, true, true);
        Assert.StartsWith("ofd:", face.FaceName); Assert.Equal(original, PdfFontRegistry.GetOriginalFont(face.FaceName));
    }

    private static byte[] TestCollection(params byte[][] faces)
    {
        var offset = 12 + faces.Length * 4;
        var output = new byte[offset + faces.Sum(face => face.Length)];
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(output, 0, 0x74746366);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(output, 4, 0x00010000);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(output, 8, (uint)faces.Length);
        for (var index = 0; index < faces.Length; index++)
        {
            Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(output, 12 + index * 4, (uint)offset);
            Buffer.BlockCopy(faces[index], 0, output, offset, faces[index].Length);
            for (var table = 0; table < Ofdrw.Net.Core.Fonts.OpenTypeFace.U16(faces[index], 4); table++)
            {
                var record = offset + 12 + table * 16 + 8;
                Ofdrw.Net.Core.Fonts.OpenTypeFace.Put32(output, record, Ofdrw.Net.Core.Fonts.OpenTypeFace.U32(output, record) + (uint)offset);
            }
            offset += faces[index].Length;
        }
        return output;
    }

    [Fact]
    public void NamedHostCollectionSelectsNonzeroFaceAndRefusesAmbiguousOrOpaqueNames()
    {
        var first = Ofdrw.Net.Core.Fonts.OpenTypeFontIdentity.WithUniqueNames(File.ReadAllBytes(FontPath("narrow")), "first");
        var bold = new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("wide")));
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(bold.Table("OS/2"), 62, 0x20);
        var second = Ofdrw.Net.Core.Fonts.OpenTypeFontIdentity.WithUniqueNames(bold.Build(), "second");
        var collection = TestCollection(first, second);
        Assert.Equal(second, Ofdrw.Net.Core.Fonts.OpenTypeCollection.SelectNamedFace(collection, "Ofdrw-second"));
        Assert.Equal(second, Ofdrw.Net.Core.Fonts.OpenTypeCollection.SelectNamedFace(collection, "Osecond"));
        Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Core.Fonts.OpenTypeCollection.SelectNamedFace(collection, "opaque-face"));
        Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Core.Fonts.OpenTypeCollection.SelectNamedFace(TestCollection(first, first), "Ofdrw-first"));
        Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Core.Fonts.OpenTypeCollection.SelectNamedFace(TestCollection(first), "opaque-face"));
        Assert.Equal(first, Ofdrw.Net.Core.Fonts.OpenTypeCollection.SelectNamedFace(TestCollection(first), "opaque-face", allowSingleFace: true));
        var resource = new OfdFontResource { Id = "10", FontName = "collection-host", Bold = true };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], new NamedCollectionResolver(collection, "Ofdrw-second", first));
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "AB" }, out _);
        var info = GlobalFontSettings.FontResolver.ResolveTypeface(family, true, false);
        Assert.Equal(second, PdfFontRegistry.GetOriginalFont(info.FaceName)); Assert.False(info.MustSimulateBold);
    }

    [Fact]
    public void UnidentifiedCollectionProbeUsesValidatedStandaloneFallbackForNonemptyText()
    {
        var primary = File.ReadAllBytes(FontPath("narrow")); var fallback = File.ReadAllBytes(FontPath("wide"));
        var resource = new OfdFontResource { Id = "10", FontName = "optional-ttc", Bold = true };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], new NamedCollectionResolver(TestCollection(primary), "opaque-face", fallback));
        var family = context.Resolve(new OfdTextElement { FontResourceId = "10", Text = "ABCD" }, out _);
        var info = GlobalFontSettings.FontResolver.ResolveTypeface(family, true, false);
        Assert.Equal(fallback, PdfFontRegistry.GetOriginalFont(info.FaceName)); Assert.NotEqual(primary, PdfFontRegistry.GetOriginalFont(info.FaceName));
    }

    private sealed class NamedCollectionResolver(byte[] collection, string selectedName, byte[] fallback) : IFontResolver
    {
        public string DefaultFontName => "standalone-default";
        public FontResolverInfo ResolveTypeface(string family, bool bold, bool italic) => new(family == DefaultFontName || family == "Arial" ? "default-face" : selectedName);
        public byte[] GetFont(string face) => face == "default-face" ? fallback : collection;
    }

    [Theory]
    [InlineData("A\u200EB")][InlineData("abc\u200Eאב")]
    public async Task LrmSemanticsRefusePdfWhileOriginalOfdTextRemainsComplete(string value)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "name-only-bidi" });
        package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50,
            Elements = { new OfdTextElement { FontResourceId = "10", FontName = "name-only-bidi", Text = value } } });
        using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
        var read = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd);
        Assert.Equal(value, Assert.IsType<OfdTextElement>(read.Pages[0].Elements[0]).Text); ofd.Position = 0;
        using var pdf = new MemoryStream();
        var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
        Assert.Contains("bidi", exception.Message); Assert.Equal(0, pdf.Length);
    }

    [Theory]
    [InlineData("x\u2060y")][InlineData("f\u2061(x)")][InlineData("2\u2062x")][InlineData("1\u20632")][InlineData("2\u2064x")]
    [InlineData("of\u00ADfice")][InlineData("A\u200BB")][InlineData("A\u034FB")][InlineData("A\u180EB")]
    [InlineData("\uFEFFAB")][InlineData("A\uFEFFB")][InlineData("A\u206AB")][InlineData("A\u206FB")]
    public async Task SemanticDefaultIgnorablesNeverSilentlyDisappearFromPdf(string value)
    {
        foreach (var positioned in new[] { false, true })
        {
            var package = new OfdDocumentPackage();
            package.Fonts.Add(new OfdFontResource { Id = "10", FontName = "semantic-controls" });
            var text = new OfdTextElement { FontResourceId = "10", FontName = "semantic-controls", Text = value };
            if (positioned) text.Runs.Add(new OfdTextRun { Text = value, DeltaX = "5 5 5 5 5" });
            package.Pages.Add(new OfdPage { WidthMillimeters = 100, HeightMillimeters = 50, Elements = { text } });
            using var ofd = new MemoryStream(); await new OfdPackageWriter().WriteAsync(package, ofd); ofd.Position = 0;
            var read = await new Ofdrw.Net.Reader.Readers.OfdReader().ReadAsync(ofd); var saved = Assert.IsType<OfdTextElement>(read.Pages[0].Elements[0]);
            Assert.Equal(value, saved.Runs.Count == 0 ? saved.Text : string.Concat(saved.Runs.Select(run => run.Text))); ofd.Position = 0;
            using var pdf = new MemoryStream(); var exception = await Assert.ThrowsAsync<NotSupportedException>(() => new OfdToPdfConverter().ConvertAsync(ofd, pdf));
            Assert.Contains("default-ignorable", exception.Message); Assert.Equal(0, pdf.Length);
        }
    }

    private sealed class CoveringResolver(byte[] primary,byte[] fallback):IFontResolver
    {
        public string DefaultFontName=>"covering-default";
        public FontResolverInfo ResolveTypeface(string family,bool bold,bool italic)=>new(family==DefaultFontName?"covering-face":"primary-face");
        public byte[] GetFont(string face)=>face=="covering-face"?fallback:primary;
    }

    [Fact]
    public void RegisteredNameOnlySourceBytesAlsoProvideCoverageForMissingFillers()
    {
        var data=File.ReadAllBytes(FontPath("narrow")); var host=new HostResolver(data);
        var resource=new OfdFontResource { Id="10",FontName="host-CJK-substitute",Bold=true };
        var context=new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource],host);
        var coverage=context.Coverage(resource); Assert.NotNull(coverage);
        Assert.Equal(0,coverage!.Glyph(0x3164));
        Assert.Equal("AB",Ofdrw.Net.Converter.Pdf.Internal.PdfTextControlPolicy.VisibleText("A\u3164B",coverage));
        var mapped=new Ofdrw.Net.Core.Fonts.OpenTypeFace(data);var original=new Ofdrw.Net.Core.Fonts.OpenTypeCmap(mapped);
        mapped.Tables["cmap"]=Ofdrw.Net.Core.Fonts.OpenTypeCmap.Build(new Dictionary<int,int>{['A']=original.Glyph('A'),['B']=original.Glyph('B'),[0x3164]=original.Glyph(' ')});
        var mappedContext=new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource],new HostResolver(mapped.Build()));
        Assert.Equal("A\u3164B",Ofdrw.Net.Converter.Pdf.Internal.PdfTextControlPolicy.VisibleText("A\u3164B",mappedContext.Coverage(resource)));
    }

    [Fact]
    public void NonRenderingControlsAreNotPaintedAndSemanticBidiOrUvsAreRefused()
    {
        Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Converter.Pdf.Internal.PdfTextControlPolicy.VisibleText("A\u200CB"));
        Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Converter.Pdf.Internal.PdfTextControlPolicy.VisibleText("A\u200DB"));
        Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Converter.Pdf.Internal.PdfTextControlPolicy.VisibleText("\u200D"));
        foreach (var value in new[] { "A\u206AB", "A\u206FB", "A\u200EB", "A\u200CB", "A\u200DB", "A\u202EB", "A\u202CB", "A\u200FB", "A\u061CB", "A\u2067B", "A\uFE00", "A\u180BB", "A\u180FB" })
            Assert.Throws<NotSupportedException>(() => Ofdrw.Net.Converter.Pdf.Internal.PdfTextControlPolicy.Validate(value));
    }

    [Fact]
    public void Os2StyleFlagsDoNotCauseDoubleBoldOrItalicSimulation()
    {
        var face = new Ofdrw.Net.Core.Fonts.OpenTypeFace(File.ReadAllBytes(FontPath("narrow")));
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(face.Table("head"), 44, 0);
        Ofdrw.Net.Core.Fonts.OpenTypeFace.Put16(face.Table("OS/2"), 62, 0x21);
        var family = PdfFontRegistry.RegisterFontFace(face.Build(), bold: true, italic: true);
        var info = PdfSharpCore.Fonts.GlobalFontSettings.FontResolver.ResolveTypeface(family, true, true);
        Assert.False(info.MustSimulateBold); Assert.False(info.MustSimulateItalic);
    }

    [Fact]
    public void SupplementaryUnicodeAndExplicitGlyphSourceFailRatherThanCorruptPdfText()
    {
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([]);
        var exception = Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement { Text = "A\U000107A5B" }, out _));
        Assert.Contains("supplementary Unicode", exception.Message);
        foreach (var xml in new[] { "<CGTransform><Glyphs>1</Glyphs></CGTransform>", "<CGTransform" })
            Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement { Text = "中文", SourceXml = xml }, out _));
        Assert.Throws<NotSupportedException>(() => context.Resolve(new OfdTextElement { Text = "A", SourceXml = "<TextObject><CGTransform><Glyphs>1</Glyphs></CGTransform></TextObject>" }, out _));
    }

    [Fact]
    public void NameOnlyFallback_ShouldNotProbeRegularHostFaces()
    {
        var host = new HostResolver(File.ReadAllBytes(FontPath("narrow")));
        var resource = new OfdFontResource { Id = "probe", FontName = "ordinary-host" };
        var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
        Assert.Equal("ordinary-host", context.Resolve(new OfdTextElement { FontResourceId = "probe", FontName = "ordinary-host" }, out _));
        Assert.Equal(0, host.Requests);
    }

    [Fact]
    public void NameOnlyFallback_ShouldSkipOptionalRegistrationWhenBudgetIsFull()
    {
        PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("narrow")));
        var retained = PdfFontRegistry.RegisteredFontBytes;
        var budget = PdfFontRegistry.MaximumRegisteredFontBytes;
        try
        {
            PdfFontRegistry.MaximumRegisteredFontBytes = retained;
            var host = new HostResolver(File.ReadAllBytes(FontPath("budget")));
            var resource = new OfdFontResource { Id = "probe", FontName = "budget-host", Bold = true };
            var context = new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], host);
            Assert.Equal("budget-host", context.Resolve(new OfdTextElement { FontResourceId = "probe", FontName = "budget-host" }, out _));
            Assert.Equal(retained, PdfFontRegistry.RegisteredFontBytes);
        }
        finally { PdfFontRegistry.MaximumRegisteredFontBytes = budget; }
    }

    [Fact]
    public void NameOnlyFallback_ShouldNotSwallowCancellation()
    {
        var resource = new OfdFontResource { Id = "probe", FontName = "cancel-host", Bold = true };
        Assert.Throws<OperationCanceledException>(() =>
            new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource], new ProbeResolver("cancel")));
    }

    [Fact]
    public void EmbeddedFont_ShouldStillRejectInvalidPayload()
    {
        var resource = new OfdFontResource { Id = "embedded", Data = new byte[12] };
        Assert.ThrowsAny<Exception>(() => new Ofdrw.Net.Converter.Pdf.Internal.DocumentFontContext([resource]));
    }

    private sealed class ProbeResolver(string failure) : IFontResolver
    {
        public string DefaultFontName => "probe";
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic) => failure switch
        {
            "resolve" => throw new ArgumentException("Host cannot resolve this family."),
            "cancel" => throw new OperationCanceledException(),
            _ => new FontResolverInfo("probe")
        };
        public byte[] GetFont(string faceName) => failure switch
        {
            "read" => throw new IOException("Host font file is unavailable."),
            "ttc" => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "fonts", "style-collection.ttc")),
            "woff" => [119, 79, 70, 70, 0, 0, 0, 0, 0, 0, 0, 0],
            _ => [0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0]
        };
    }

    [Fact]
    public void ComposedResolver_ShouldPreserveHostRequestsAndResolveEmbeddedFaces()
    {
        var hostData = File.ReadAllBytes(FontPath("narrow"));
        var host = new HostResolver(hostData);
        var combined = PdfFontRegistry.CreateResolver(host);
        Assert.Equal("host-face", combined.ResolveTypeface("host", false, false).FaceName);
        Assert.Equal(hostData, combined.GetFont("host-face"));
        var family = PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("wide")));
        var face = combined.ResolveTypeface(family, false, false);
        Assert.StartsWith("ofd:", face.FaceName);
        using var data = new MemoryStream(combined.GetFont(face.FaceName));
        Assert.StartsWith("Ofdrw-", SixLabors.Fonts.FontDescription.LoadDescription(data).FontFamilyInvariantCulture);
        Assert.Equal(1, host.Requests);
    }

    [Fact]
    public async Task SameNamedFonts_ShouldRemainDistinctAcrossSequentialAndConcurrentDocuments()
    {
        var narrow = await ConvertAsync("narrow");
        var wide = await ConvertAsync("wide");
        Assert.True(wide > narrow * 1.5, $"Font widths were conflated: narrow={narrow}, wide={wide}");
        var widths = await Task.WhenAll(Enumerable.Range(0, 8).Select(index => ConvertAsync(index % 2 == 0 ? "narrow" : "wide")));
        for (var index = 0; index < widths.Length; index++) Assert.Equal(index % 2 == 0 ? narrow : wide, widths[index], precision: 3);
    }

    [Fact]
    public void Registration_ShouldSnapshotBytesAndRejectNameRebinding()
    {
        var original = File.ReadAllBytes(FontPath("narrow"));
        var mutable = (byte[])original.Clone();
        var family = PdfFontRegistry.RegisterFontFace(mutable);
        var retainedBytes = PdfFontRegistry.RegisteredFontBytes;
        Array.Clear(mutable);
        var face = GlobalFontSettings.FontResolver.ResolveTypeface(family, false, false);
        Assert.Equal(original, PdfFontRegistry.GetOriginalFont(face.FaceName));
        Assert.Equal(family, PdfFontRegistry.RegisterFontFace(original));
        Assert.Equal(retainedBytes, PdfFontRegistry.RegisteredFontBytes);
        var name = "legacy-test-" + Guid.NewGuid();
        PdfFontRegistry.RegisterFont(name, original);
        Assert.Throws<InvalidOperationException>(() => PdfFontRegistry.RegisterFont(name, File.ReadAllBytes(FontPath("wide"))));
        Assert.Equal(face.FaceName, GlobalFontSettings.FontResolver.ResolveTypeface(name, false, false).FaceName);
    }

    [Fact]
    public void Registration_ShouldBoundRetainedBytesWithoutEvictingActiveFaces()
    {
        PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("narrow")));
        var originalBudget = PdfFontRegistry.MaximumRegisteredFontBytes;
        var retained = PdfFontRegistry.RegisteredFontBytes;
        try
        {
            PdfFontRegistry.MaximumRegisteredFontBytes = retained;
            Assert.Throws<InvalidOperationException>(() => PdfFontRegistry.RegisterFontFace(File.ReadAllBytes(FontPath("budget"))));
            Assert.Equal(retained, PdfFontRegistry.RegisteredFontBytes);
        }
        finally { PdfFontRegistry.MaximumRegisteredFontBytes = originalBudget; }
    }

    private static async Task<double> ConvertAsync(string variant)
    {
        var package = new OfdDocumentPackage();
        package.Fonts.Add(new OfdFontResource { Id = "12", FontName = "Ofdrw Test Face", Data = File.ReadAllBytes(FontPath(variant)) });
        var page = new OfdPage { WidthMillimeters = 100, HeightMillimeters = 60 };
        page.Elements.Add(new OfdTextElement
        {
            FontName = "Ofdrw Test Face", FontResourceId = "12", FontSizeMillimeters = 5,
            XMillimeters = 10, YMillimeters = 10, Text = "ABCD"
        });
        package.Pages.Add(page);
        using var ofd = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, ofd);
        ofd.Position = 0;
        using var output = new MemoryStream();
        await new OfdToPdfConverter().ConvertAsync(ofd, output);
        output.Position = 0;
        using var pdf = PdfPigDocument.Open(output);
        Assert.Equal("ABCD", pdf.GetPage(1).Text);
        var letters = pdf.GetPage(1).Letters;
        var artifactDirectory = Environment.GetEnvironmentVariable("OFDRW_TEST_ARTIFACTS");
        if (!string.IsNullOrEmpty(artifactDirectory))
        {
            Directory.CreateDirectory(artifactDirectory);
            await File.WriteAllBytesAsync(Path.Combine(artifactDirectory, $"font-{variant}.pdf"), output.ToArray());
            await File.WriteAllLinesAsync(Path.Combine(artifactDirectory, $"font-{variant}-letters.txt"), letters.Select(letter =>
                $"{letter.Value}: {letter.Location}; bounds={letter.BoundingBox}; width={letter.Width}"));
        }
        return letters.Max(letter => letter.BoundingBox.Right) - letters.Min(letter => letter.BoundingBox.Left);
    }

    private static string FontPath(string variant) => Path.Combine(AppContext.BaseDirectory, "fonts", variant + ".ttf");

    private sealed class HostResolver(byte[] data) : IFontResolver
    {
        internal int Requests { get; private set; }
        public string DefaultFontName => "host";
        public FontResolverInfo ResolveTypeface(string familyName, bool isBold, bool isItalic)
        {
            Requests++;
            return new FontResolverInfo("host-face");
        }
        public byte[] GetFont(string faceName) => data;
    }
}
