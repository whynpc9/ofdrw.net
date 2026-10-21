using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Reader.Readers;
using Ofdrw.Net.Converter.Pdf.Converters;
using Ofdrw.Net.Converter.Svg.Converters;

if (args.Length != 1) throw new ArgumentException("Pass one output directory under artifacts/.");
var dir = Path.GetFullPath(args[0]);
Directory.CreateDirectory(dir);
var package = new OfdDocumentPackage();
var page = new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 100 };
package.Pages.Add(page);
page.Elements.Add(new OfdPathElement { XMillimeters = 0, YMillimeters = 0, WidthMillimeters = 100, HeightMillimeters = 100,
    Fill = true, Stroke = false, FillColor = new OfdColor(255, 255, 255), AbbreviatedData = "M 0 0 L 100 0 L 100 100 L 0 100 C" });
page.Elements.Add(new OfdRawElement { LocalName = "PageBlock", Xml = "<PageBlock xmlns='http://www.ofdspec.org' ID='500'><PathObject ID='501' Boundary='10 10 50 50' Fill='true' Stroke='false'><FillColor Value='0 0 255'/><AbbreviatedData>M 0 0 L 50 0 L 50 50 L 0 50 C</AbbreviatedData></PathObject></PageBlock>" });
var ofdPath = Path.Combine(dir, "pageblock.ofd");
await using (var output = File.Create(ofdPath)) await new OfdPackageWriter().WriteAsync(package, output);
await using (var input = File.OpenRead(ofdPath)) package = await new OfdReader().ReadAsync(input);
Console.WriteLine($"Reader: {string.Join(",", package.Pages[0].Elements.Select(e => e.GetType().Name))}");
await using (var input = File.OpenRead(ofdPath)) await using (var output = File.Create(Path.Combine(dir, "pageblock.pdf")))
    await new OfdToPdfConverter().ConvertAsync(input, output);
await using (var input = File.OpenRead(ofdPath)) await using (var output = File.Create(Path.Combine(dir, "pageblock.svg")))
    await new OfdToSvgConverter().ConvertAsync(input, output);
await using var htmlInput = File.OpenRead(ofdPath);
await using var htmlOutput = new MemoryStream();
try { await new OfdToHtmlConverter().ConvertAsync(htmlInput, htmlOutput); Console.WriteLine("HTML: success"); }
catch (NotSupportedException ex) { Console.WriteLine($"HTML: rejected ({ex.Message}), bytes={htmlOutput.Length}"); }
Console.WriteLine($"SVG contains blue: {File.ReadAllText(Path.Combine(dir, "pageblock.svg")).Contains("0,0,255", StringComparison.Ordinal)}");
