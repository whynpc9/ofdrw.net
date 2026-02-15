using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Http.Json;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc.Testing;
using Ofdrw.Net.Core.Models;
using Ofdrw.Net.Layout.Builders;
using Ofdrw.Net.Packaging;

namespace Ofdrw.Net.Service.Tests;

public sealed class ApiTests : IClassFixture<WebApplicationFactory<Program>>
{
    private readonly WebApplicationFactory<Program> _factory;

    public ApiTests(WebApplicationFactory<Program> factory)
    {
        _factory = factory;
    }

    [Fact]
    public async Task GetProfile_ShouldReturnOk()
    {
        var client = _factory.CreateClient();
        var response = await client.GetAsync("/api/v1/ofd/validate/profiles/emr-ofd-h-202x");

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var body = await response.Content.ReadAsStringAsync();
        Assert.Contains("emr-ofd-h-202x", body);
    }

    [Fact]
    public async Task ValidateEndpoint_ShouldReturnReport()
    {
        var client = _factory.CreateClient();
        var content = new MultipartFormDataContent();

        var bytes = await CreateSampleOfdAsync();
        var file = new ByteArrayContent(bytes);
        file.Headers.ContentType = new MediaTypeHeaderValue("application/ofd");
        content.Add(file, "ofd", "sample.ofd");

        var response = await client.PostAsync("/api/v1/ofd/validate/emr-tech-spec", content);

        Assert.Equal(HttpStatusCode.OK, response.StatusCode);
        var json = await response.Content.ReadAsStringAsync();
        Assert.Contains("profileVersion", json, System.StringComparison.OrdinalIgnoreCase);
    }

    private static async Task<byte[]> CreateSampleOfdAsync()
    {
        var builder = new OfdDocumentBuilder();
        builder.AddPage(new Ofdrw.Net.Core.Models.OfdPage
        {
            Index = 0,
            WidthMillimeters = 210,
            HeightMillimeters = 297,
            Elements =
            {
                new OfdTextElement
                {
                    Text = "sample",
                    XMillimeters = 10,
                    YMillimeters = 10,
                    WidthMillimeters = 50,
                    HeightMillimeters = 8,
                    FontSizeMillimeters = 4
                }
            }
        });

        await using var ms = new MemoryStream();
        var writer = new OfdPackageWriter();
        await writer.WriteAsync(builder.Build(), ms);
        return ms.ToArray();
    }
}
