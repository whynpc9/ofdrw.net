using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Ofdrw.Net.Converter.Svg;
using Ofdrw.Net.Converter.Svg.Converters;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Packaging;

var destination = Path.GetFullPath(args[0]); Directory.CreateDirectory(destination);
var package = new OfdDocumentPackage();
for (var i = 0; i < 3; i++)
{
    var page = new OfdPage { Index = i, WidthMillimeters = i == 1 ? 150 : 100, HeightMillimeters = 120 };
    page.Elements.Add(new OfdTextElement { Text = $"Package page {i + 1} 中文 Wi & <literal>", XMillimeters = 8, YMillimeters = 10, WidthMillimeters = 90, HeightMillimeters = 20 });
    package.Pages.Add(page);
}
var source = Path.Combine(destination, "package-three-pages.ofd");
await using (var stream = File.Create(source)) await new OfdPackageWriter().WriteAsync(package, stream);
var html = Path.Combine(destination, "package-api.html");
await using (var input = File.OpenRead(source))
await using (var output = File.Create(html))
    await new OfdToHtmlConverter(new OfdHtmlOptions { Title = "Package <literal> & title" }).ConvertAsync(input, output, new[] { 2, 0, 2 });
var document = XDocument.Load(html);
if (!document.Descendants("section").Select(n => n.Attribute("data-source-page")!.Value).SequenceEqual(new[] { "3", "1", "3" })) throw new Exception("Package HTML page order failed.");
if (document.Descendants("title").Single().Value != "Package <literal> & title") throw new Exception("Package HTML escaping failed.");
if (document.Descendants().Any(n => n.Name.LocalName == "script")) throw new Exception("Package HTML policy failed.");
Console.WriteLine("HTML API consumed from local nupkg: three page occurrences, ordering and escaping passed.");
