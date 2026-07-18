using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Layout.Editing;
using Ofdrw.Net.Packaging;
using Ofdrw.Net.Packaging.Archive;
using Ofdrw.Net.Packaging.Validation;
using Ofdrw.Net.Reader.Extraction;
using Ofdrw.Net.Reader.Readers;

namespace Ofdrw.Net.Packaging.Tests;

public sealed class PackageRoundTripTests
{
    [Fact]
    public async Task Writer_ShouldGenerateRequiredEntries_AndReaderCanParse()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "hello ofd",
                    XMillimeters = 10,
                    YMillimeters = 12,
                    WidthMillimeters = 60,
                    HeightMillimeters = 8,
                    FontSizeMillimeters = 4
                }
            }
        });

        var package = builder.Build();
        var writer = new OfdPackageWriter();

        await using var ms = new MemoryStream();
        await writer.WriteAsync(package, ms);
        ms.Position = 0;

        var loader = new OfdPackageLoader();
        var archive = await loader.LoadAsync(ms);

        Assert.Contains("OFD.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Document.xml", archive.EntryNames);
        Assert.Contains("Doc_0/Pages/Page_0/Content.xml", archive.EntryNames);

        var pageXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Pages/Page_0/Content.xml"));
        var ns = pageXml.Root!.Name.Namespace;
        var textObject = pageXml.Descendants(ns + "TextObject").Single();
        var fontReference = textObject.Attribute("Font")?.Value;
        Assert.False(string.IsNullOrWhiteSpace(fontReference));
        Assert.Equal("hello ofd", textObject.Element(ns + "TextCode")?.Value);

        var publicResXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/PublicRes.xml"));
        var publicResNs = publicResXml.Root!.Name.Namespace;
        var font = publicResXml.Descendants(publicResNs + "Font").Single();
        Assert.Equal(fontReference, font.Attribute("ID")?.Value);
        Assert.Equal("SimSun", font.Attribute("FontName")?.Value);

        ms.Position = 0;
        var reader = new OfdReader();
        var parsed = await reader.ReadAsync(ms);

        Assert.Single(parsed.Pages);
        var firstText = parsed.Pages[0].Elements.OfType<OfdTextElement>().First();
        Assert.Equal("hello ofd", firstText.Text);
        Assert.Equal("SimSun", firstText.FontName);
    }

    [Fact]
    public async Task Writer_ShouldStoreImagesInDocumentResources_AndReaderCanResolveThem()
    {
        var imageData = new byte[] { 0x89, 0x50, 0x4E, 0x47 };
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdImageElement
                {
                    XMillimeters = 0,
                    YMillimeters = 0,
                    WidthMillimeters = 210,
                    HeightMillimeters = 297,
                    Data = imageData,
                    MediaType = "image/png"
                }
            }
        });

        var writer = new OfdPackageWriter();
        await using var ms = new MemoryStream();
        await writer.WriteAsync(builder.Build(), ms);
        ms.Position = 0;

        var loader = new OfdPackageLoader();
        var archive = await loader.LoadAsync(ms);

        Assert.Contains("Doc_0/DocumentRes.xml", archive.EntryNames);
        var imageEntry = Assert.Single(
            archive.EntryNames,
            x => x.StartsWith(
                "Doc_0/Res/Image_",
                System.StringComparison.Ordinal));
        Assert.DoesNotContain("Doc_0/Pages/Page_0/PageRes.xml", archive.EntryNames);

        var documentResXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/DocumentRes.xml"));
        var ns = documentResXml.Root!.Name.Namespace;
        Assert.Equal("Res", documentResXml.Root.Attribute("BaseLoc")?.Value);
        var media = documentResXml.Descendants(ns + "MultiMedia").Single();
        var resourceId = media.Attribute("ID")?.Value;
        Assert.False(string.IsNullOrWhiteSpace(resourceId));
        Assert.Equal("PNG", media.Attribute("Format")?.Value);
        Assert.Equal(Path.GetFileName(imageEntry), media.Element(ns + "MediaFile")?.Value);

        var pageXml = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Pages/Page_0/Content.xml"));
        var pageNs = pageXml.Root!.Name.Namespace;
        var imageObject = pageXml.Descendants(pageNs + "ImageObject").Single();
        Assert.Equal("210 0 0 297 0 0", imageObject.Attribute("CTM")?.Value);
        Assert.Equal(resourceId, imageObject.Attribute("ResourceID")?.Value);

        ms.Position = 0;
        var reader = new OfdReader();
        var parsed = await reader.ReadAsync(ms);
        var image = Assert.Single(parsed.Pages[0].Elements.OfType<OfdImageElement>());
        Assert.Equal(imageData, image.Data);
        Assert.Equal("image/png", image.MediaType);
        Assert.Equal(Path.GetFileName(imageEntry), image.FileName);
    }

    [Fact]
    public async Task Writer_ShouldUseUniqueDocumentUnitIds()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement { Text = "A", FontName = "SimSun", FontSizeMillimeters = 4 },
                new OfdImageElement
                {
                    WidthMillimeters = 10,
                    HeightMillimeters = 10,
                    Data = [0x89, 0x50, 0x4E, 0x47]
                }
            }
        });

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(stream);

        var xmlEntries = archive.EntryNames
            .Where(x => x.EndsWith(".xml", System.StringComparison.OrdinalIgnoreCase))
            .Select(x => XDocument.Parse(archive.ReadUtf8Text(x)))
            .ToList();
        var ids = xmlEntries
            .SelectMany(x => x.Descendants().Attributes("ID"))
            .Select(x => long.Parse(x.Value, System.Globalization.CultureInfo.InvariantCulture))
            .ToList();

        Assert.Equal(ids.Count, ids.Distinct().Count());
        var document = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Document.xml"));
        var maxUnitId = long.Parse(
            document.Descendants().Single(x => x.Name.LocalName == "MaxUnitID").Value,
            System.Globalization.CultureInfo.InvariantCulture);
        Assert.Equal(ids.Max(), maxUnitId);
    }

    [Fact]
    public async Task Writer_ShouldUseStandardAttachmentAndCustomTagReferences()
    {
        var package = new OfdDocumentPackage();
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 210, HeightMillimeters = 297 });
        package.Attachments.Add(new OfdAttachment
        {
            Id = "report",
            Name = "report.txt",
            Format = "text/plain",
            MediaType = "text/plain",
            Data = [1, 2, 3]
        });
        package.CustomTags["record-id"] = "42";

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(stream);

        var document = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Document.xml"));
        Assert.Equal("Attachs/Attachments.xml", document.Descendants().Single(x => x.Name.LocalName == "Attachments").Value);
        Assert.Equal("Tags/CustomTags.xml", document.Descendants().Single(x => x.Name.LocalName == "CustomTags").Value);

        var attachments = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Attachs/Attachments.xml"));
        var attachment = attachments.Descendants().Single(x => x.Name.LocalName == "Attachment");
        Assert.Null(attachment.Attribute("FileLoc"));
        Assert.Equal("Attach_1_report.txt", attachment.Elements().Single(x => x.Name.LocalName == "FileLoc").Value);

        var customTags = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Tags/CustomTags.xml"));
        var customTag = customTags.Descendants().Single(x => x.Name.LocalName == "CustomTag");
        Assert.Equal("EMR", customTag.Attribute("TypeID")?.Value);
        Assert.Null(customTag.Attribute("FileLoc"));
        Assert.Equal("CustomTag_EMR.xml", customTag.Elements().Single(x => x.Name.LocalName == "FileLoc").Value);

        stream.Position = 0;
        var parsed = await new OfdReader().ReadAsync(stream);
        Assert.Equal([1, 2, 3], Assert.Single(parsed.Attachments).Data);
        Assert.Equal("42", parsed.CustomTags["record-id"]);
    }

    [Fact]
    public async Task ReaderAndWriter_ShouldPreserveUnknownPageObjectsAndPackageEntries()
    {
        const string ns = "http://www.ofdspec.org/2016";
        var package = new OfdDocumentPackage();
        package.PreservedEntries["Doc_0/Extensions/opaque.bin"] = [9, 8, 7];
        package.Pages.Add(new OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdPathElement
                {
                    ObjectId = "50",
                    LayerId = "40",
                    LayerType = "Body",
                    XMillimeters = 1,
                    YMillimeters = 2,
                    WidthMillimeters = 3,
                    HeightMillimeters = 4,
                    AbbreviatedData = "M 0 0 L 1 1",
                    StrokeColor = new OfdColor(10, 20, 30),
                    SourceXml = $"<ofd:PathObject xmlns:ofd=\"{ns}\" ID=\"50\" Boundary=\"1 2 3 4\" VendorStyle=\"retained\"><ofd:AbbreviatedData>M 0 0 L 1 1</ofd:AbbreviatedData></ofd:PathObject>"
                },
                new OfdRawElement
                {
                    ObjectId = "60",
                    LayerId = "41",
                    LayerType = "Foreground",
                    LocalName = "CompositeObject",
                    Xml = $"<ofd:CompositeObject xmlns:ofd=\"{ns}\" ID=\"60\" Boundary=\"5 6 7 8\" ResourceID=\"70\" />"
                }
            }
        });

        await using var first = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, first);
        first.Position = 0;
        var parsed = await new OfdReader().ReadAsync(first);

        var parsedPath = Assert.Single(parsed.Pages[0].Elements.OfType<OfdPathElement>());
        Assert.Equal("M 0 0 L 1 1", parsedPath.AbbreviatedData);
        Assert.Equal(10, parsedPath.StrokeColor.Red);
        Assert.Single(parsed.Pages[0].Elements.OfType<OfdRawElement>());
        Assert.Equal(2, parsed.Pages[0].Elements.Select(x => x.LayerId).Distinct().Count());
        Assert.Equal([9, 8, 7], parsed.PreservedEntries["Doc_0/Extensions/opaque.bin"]);

        await using var second = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(parsed, second);
        second.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(second);
        Assert.Equal([9, 8, 7], archive.GetBytes("Doc_0/Extensions/opaque.bin"));
        var content = XDocument.Parse(archive.ReadUtf8Text("Doc_0/Pages/Page_0/Content.xml"));
        var writtenPath = Assert.Single(content.Descendants(), x => x.Name.LocalName == "PathObject");
        Assert.Equal("retained", writtenPath.Attribute("VendorStyle")?.Value);
        Assert.Single(content.Descendants(), x => x.Name.LocalName == "CompositeObject");
        Assert.Equal(2, content.Descendants().Count(x => x.Name.LocalName == "Layer"));
    }

    [Fact]
    public async Task Loader_ShouldRejectEntriesAboveConfiguredLimit()
    {
        await using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            var entry = zip.CreateEntry("large.bin");
            await using var entryStream = entry.Open();
            await entryStream.WriteAsync(new byte[32]);
        }

        stream.Position = 0;
        var options = new OfdPackageLoadOptions
        {
            MaxEntryCount = 10,
            MaxEntryUncompressedBytes = 16,
            MaxTotalUncompressedBytes = 64,
            MaxCompressionRatio = 1_000
        };
        await Assert.ThrowsAsync<InvalidDataException>(() => new OfdPackageLoader().LoadAsync(stream, options));
    }

    [Fact]
    public async Task StructureChecker_ShouldFollowConfiguredDocumentRoot()
    {
        var package = new OfdDocumentPackage();
        package.Options.DocumentId = "Document_A";
        package.Pages.Add(new OfdPage { Index = 0, WidthMillimeters = 100, HeightMillimeters = 100 });

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var archive = await new OfdPackageLoader().LoadAsync(stream);

        Assert.DoesNotContain(OfdPackageStructureChecker.Check(archive), x => x.IsError);
    }

    [Fact]
    public async Task EditorAndTextExtractor_ShouldReorderRemoveCropAndPersistPages()
    {
        var package = new OfdDocumentPackage();
        foreach (var label in new[] { "Alpha", "Beta", "Gamma" })
        {
            package.Pages.Add(new OfdPage
            {
                Index = package.Pages.Count,
                WidthMillimeters = 210,
                HeightMillimeters = 297,
                Elements =
                {
                    new OfdTextElement
                    {
                        Text = label,
                        XMillimeters = 20,
                        YMillimeters = 30,
                        WidthMillimeters = 30,
                        HeightMillimeters = 8
                    }
                }
            });
        }

        OfdDocumentEditor.ReorderPages(package, new[] { 2, 0, 1 });
        Assert.Equal(
            new[] { "Gamma", "Alpha", "Beta" },
            new OfdTextExtractor().ExtractPages(package));

        OfdDocumentEditor.RemovePages(package, new[] { 1 });
        OfdDocumentEditor.CropPage(package.Pages[0], 10, 15, 80, 90);
        Assert.Equal(new[] { "Gamma", "Beta" }, new OfdTextExtractor().ExtractPages(package));

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(package, stream);
        stream.Position = 0;
        var parsed = await new OfdReader().ReadAsync(stream);

        Assert.Equal(2, parsed.Pages.Count);
        Assert.Equal(10, parsed.Pages[0].XMillimeters);
        Assert.Equal(15, parsed.Pages[0].YMillimeters);
        Assert.Equal(80, parsed.Pages[0].WidthMillimeters);
        Assert.Equal(90, parsed.Pages[0].HeightMillimeters);
        Assert.Equal(new[] { "Gamma", "Beta" }, new OfdTextExtractor().ExtractPages(parsed));
    }

    [Fact]
    public async Task DocumentMerger_ShouldCreateSelfContainedOrderedPages()
    {
        static OfdDocumentPackage CreateSource(string label)
        {
            var source = new OfdDocumentPackage();
            source.Pages.Add(new OfdPage
            {
                Index = 0,
                WidthMillimeters = 100,
                HeightMillimeters = 120,
                Elements =
                {
                    new OfdTextElement
                    {
                        Text = label,
                        FontName = "Arial",
                        XMillimeters = 10,
                        YMillimeters = 15,
                        WidthMillimeters = 40,
                        HeightMillimeters = 8
                    },
                    new OfdPathElement
                    {
                        XMillimeters = 5,
                        YMillimeters = 20,
                        WidthMillimeters = 60,
                        HeightMillimeters = 1,
                        AbbreviatedData = "M 0 0 L 60 0"
                    }
                }
            });
            return source;
        }

        var merged = OfdDocumentMerger.Merge(new[]
        {
            CreateSource("First"),
            CreateSource("Second")
        });

        Assert.Equal(2, merged.Pages.Count);
        Assert.Equal(new[] { "First", "Second" }, new OfdTextExtractor().ExtractPages(merged));
        Assert.All(merged.Pages, page => Assert.Single(page.Elements.OfType<OfdPathElement>()));

        await using var stream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(merged, stream);
        stream.Position = 0;
        var parsed = await new OfdReader().ReadAsync(stream);

        Assert.Equal(new[] { "First", "Second" }, new OfdTextExtractor().ExtractPages(parsed));
        Assert.Equal(2, parsed.Pages.Sum(page => page.Elements.OfType<OfdPathElement>().Count()));
    }

    [Fact]
    public async Task UpstreamComplexSample_ShouldRetainLayersPathsTemplatesAndAnnotations()
    {
        var samplePath = Path.Combine(
            ResolveRepositoryRoot(),
            "e2e",
            "Ofdrw.Net.Converter.Pdf.E2E",
            "testdata",
            "upstream-ofdrw",
            "999.ofd");

        await using var stream = File.OpenRead(samplePath);
        var parsed = await new OfdReader().ReadAsync(stream);

        Assert.Equal(560, parsed.Pages.Sum(page => page.Elements.OfType<OfdTextElement>().Count()));
        Assert.Equal(121, parsed.Pages.Sum(page => page.Templates.Sum(template =>
            template.Elements.OfType<OfdTextElement>().Count())));
        Assert.Equal(79, parsed.Pages.Sum(page => page.Templates.Sum(template =>
            template.Elements.OfType<OfdPathElement>().Count())));
        var firstPath = parsed.Pages.SelectMany(page => page.Templates)
            .SelectMany(template => template.Elements)
            .OfType<OfdPathElement>()
            .First();
        Assert.False(string.IsNullOrWhiteSpace(firstPath.AbbreviatedData));
        Assert.NotNull(firstPath.SourceXml);
        Assert.Equal(10, parsed.Pages.Sum(page =>
            page.Elements.Concat(page.Templates.SelectMany(template => template.Elements))
                .Select(element => element.LayerId)
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct()
                .Count()));
        Assert.Equal(5, parsed.Pages.Sum(page => page.PreservedPageElements.Count(xml => xml.Contains("Template"))));
        Assert.Contains(parsed.PreservedEntries.Keys, path => path.Contains("/Annots/", System.StringComparison.OrdinalIgnoreCase));
        Assert.Contains(parsed.PreservedCommonDataElements, xml => xml.Contains("TemplatePage"));
        Assert.Contains(parsed.PreservedDocumentElements, xml => xml.Contains("Annotations"));

        await using var roundTripStream = new MemoryStream();
        await new OfdPackageWriter().WriteAsync(parsed, roundTripStream);
        roundTripStream.Position = 0;
        var roundTripArchive = await new OfdPackageLoader().LoadAsync(roundTripStream);
        Assert.DoesNotContain(OfdPackageStructureChecker.Check(roundTripArchive), issue => issue.IsError);

        roundTripStream.Position = 0;
        var roundTrip = await new OfdReader().ReadAsync(roundTripStream);
        Assert.Equal(5, roundTrip.Pages.Count);
        Assert.Equal(560, roundTrip.Pages.Sum(page => page.Elements.OfType<OfdTextElement>().Count()));
        Assert.Equal(121, roundTrip.Pages.Sum(page => page.Templates.Sum(template =>
            template.Elements.OfType<OfdTextElement>().Count())));
        Assert.Equal(79, roundTrip.Pages.Sum(page => page.Templates.Sum(template =>
            template.Elements.OfType<OfdPathElement>().Count())));
        Assert.Contains(roundTrip.PreservedEntries.Keys, path =>
            path.Contains("/Annots/", System.StringComparison.OrdinalIgnoreCase));
    }

    private static string ResolveRepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null)
        {
            if (File.Exists(Path.Combine(directory.FullName, "Ofdrw.Net.sln")))
            {
                return directory.FullName;
            }

            directory = directory.Parent;
        }

        throw new DirectoryNotFoundException("Unable to locate the Ofdrw.Net repository root.");
    }
}
