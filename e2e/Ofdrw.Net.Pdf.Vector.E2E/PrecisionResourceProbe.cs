using System.Text;
using System.Text.Json;
using System.Security.Cryptography;
using Ofdrw.Net.Converter.Pdf;
using Ofdrw.Net.Converter.Pdf.Vector;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Core.Models;

internal static class PrecisionResourceProbe
{
    internal static async Task Run(string output, byte[] font)
    {
        using var fixture = new PdfFixture(font); var reports = new List<object>();
        string Header(string title) => "0.12 0.365 0.65 rg " + fixture.Text(title + " / 中文 A  B", 34, 535, 15) + "0.12 0.365 0.65 RG 12 w ";
        var precision = fixture.Create(new[] {
            Header("Large coordinate line") + "q 1 0 0 1 -16777156 0 cm 16777216 350 m 16777217 350 l S Q",
            Header("Controls and signed rectangle") + "q 1 0 0 1 -16777156 0 cm 16777217 330 -1 60 re f 16777216 280 m 16777217 300 16777219 260 16777220 280 c S Q",
            Header("Representable native controls") + "q 1 0 0 1 -16777156 0 cm 16777216 350 m 16777218 350 l S Q 1 w 10.1 20.2 m 30.3 40.4 l S" });
        await Check("precision", precision, new[] { false, false, true }, "PATH_FLOAT_PRECISION");
        var cmap = Encoding.ASCII.GetBytes("/CIDInit /ProcSet findresource begin 12 dict begin begincmap /CIDSystemInfo << /Registry (Adobe) /Ordering (Identity) /Supplement 0 >> def /CMapName /ProbeIdentityH def /CMapType 1 def /WMode 0 def 1 begincodespacerange <0000> <FFFF> endcodespacerange 1 begincidrange <0000> <FFFF> 0 endcidrange endcmap CMapName currentdict /CMap defineresource pop end end");
        var cidMap = Enumerable.Range(0,65536).SelectMany(value => new[]{(byte)(value>>8),(byte)value}).ToArray();
        foreach (var kind in new[] { "encoding-stream", "cid-map-stream" })
        {
            var body = "0.12 0.365 0.65 rg " + fixture.Text(kind + " / 字体映射", 34, 535, 15) + "0.12 0.17 0.23 rg " + fixture.Text("Original A  B中文 remains visible in whole-page fallback", 34, 508, 9);
            var source = fixture.Create(new[] { "0.12 0.365 0.65 rg 34 330 350 120 re f", body }, encodingStream: kind == "encoding-stream" ? cmap : null, cidMapStream: kind == "cid-map-stream" ? cidMap : null);
            await Check(kind, source, new[] { true, false }, kind == "encoding-stream" ? "FONT_PROFILE" : "CID_MAPPING");
        }
        var pixels = Enumerable.Range(0,840*1190).Select(index => (byte)((index%840/80 + index/840/80)%2)).ToArray();
        var indexed = fixture.Create(new[] { "q 420 0 0 595 0 0 cm /Im1 Do Q" }, imageWidth:840,imageHeight:1190,imageBytes:pixels,imageColorSpace:"[/Indexed /DeviceRGB 1 <1f5da6e9f1f9>]");
        await Check("indexed-image", indexed, new[] { false }, "PDFV_RASTER_PAGE");
        File.WriteAllText(Path.Combine(output,"precision-resource-report.json"),JsonSerializer.Serialize(reports,new JsonSerializerOptions{WriteIndented=true}));

        async Task Check(string name, byte[] source, bool[] native, string code)
        {
            File.WriteAllBytes(Path.Combine(output,name+"-source.pdf"),source);
            using(var failInput=new MemoryStream(source))using(var failOutput=new MemoryStream())
            {
                failOutput.Write(new byte[]{7,8});
                try{await new PdfVectorToOfdConverter().ConvertAsync(failInput,failOutput);throw new Exception("Expected unsupported page.");}
                catch(NotSupportedException){if(!failOutput.ToArray().SequenceEqual(new byte[]{7,8}))throw new Exception("Fail modified destination.");}
            }
            using var input=new MemoryStream(source);using var ofd=new MemoryStream();
            var result=await new PdfVectorToOfdConverter(new(){UnsupportedPagePolicy=PdfVectorUnsupportedPagePolicy.RasterizePage,Compatibility=new(){PreferExternalPdfToPpm=true}}).ConvertWithResultAsync(input,ofd);
            if(!result.Pages.Select(p=>p.IsNative).SequenceEqual(native))throw new Exception("Wrong precision/resource page disposition.");
            ofd.Position=0;var package=await new OfdReader().ReadAsync(ofd);
            for(var i=0;i<native.Length;i++)if(!native[i])
            {
                if(result.Pages[i].ImageObjects!=1||result.Pages[i].PathObjects!=0||!result.Pages[i].Diagnostic.Contains(code)||package.Pages[i].Elements.OfType<OfdTextElement>().Any(t=>t.FillColor.Alpha!=0))throw new Exception("Visible partial native fallback leak.");
            }
            if(name=="indexed-image"&&result.Pages[0].Diagnostic.Contains("ORIGINAL_IMAGE"))throw new Exception("Indexed samples interpreted as RGB.");
            File.WriteAllBytes(Path.Combine(output,name+".ofd"),ofd.ToArray());ofd.Position=0;
            using(var pdf=File.Create(Path.Combine(output,name+".pdf")))await new OfdToPdfConverter().ConvertAsync(ofd,pdf);
            reports.Add(new{name,result,sourceSha256=Convert.ToHexStringLower(SHA256.HashData(source)),ofdBytes=ofd.Length});
        }
    }
}
